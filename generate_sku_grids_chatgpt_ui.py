#!/usr/bin/env python3
"""
generate_sku_grids_chatgpt_ui.py

UI-based version for users who DO NOT have an API key.

It uses a real browser with Playwright to:
1. Read sku_code + sku_name from a CSV/XLSX/XLS file.
2. Split records into batches of 25.
3. Open a NEW ChatGPT chat per batch (or reuse one conversation if
   --chat-url points to /c/<id>).
4. Paste the 5x5 image-generation prompt and wait for the generated image.
5. Download the raw image.
6. (label-mode "overlay", default) Cut the raw image into 25 cells and
   compose a clean final grid where the SKU code / SKU name labels are drawn
   by Pillow using the EXACT text from the Excel file. Image models often
   misspell small text, so this guarantees correct labels.
7. Save the prompt, raw image, final grid, optional per-SKU crops and a
   manifest. Completed batches are skipped on re-run.
8. Detect human-verification/CAPTCHA pages and pause for manual completion.

Final grid layout (matches the reference contact sheet):

    +-----------------+-----------------+ ...
    |   [product]     |   [product]     |
    |  SKU-EL00051    |  SKU-EL00052    |   <- bold
    | VoltBrand Tablets| PixelMax Audio |   <- sku_name before "Model-"
    | Model-0051 - Grey / 64 GB | ...    |   <- sku_name from "Model-"
    +-----------------+-----------------+ ...

Install:
    pip install playwright pandas openpyxl pillow
    playwright install chromium

Run (new chat per batch - recommended):
    python generate_sku_grids_chatgpt_ui.py ^
        --input "products.xlsx" ^
        --output "./product_grids" ^
        --start-batch 1 ^
        --end-batch 2

Reuse one existing conversation instead:
    python generate_sku_grids_chatgpt_ui.py --input products.xlsx ^
        --chat-url "https://chatgpt.com/c/YOUR_CONVERSATION_ID"

Re-draw labels on already downloaded raw images (no browser):
    python generate_sku_grids_chatgpt_ui.py --input products.xlsx --compose-only

The script uses a persistent local browser profile:
    ./chatgpt_browser_profile
Log in once in that browser profile. Future runs reuse the session.
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import logging
import re
import sys
import time
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import (
    Page,
    sync_playwright,
)


DEFAULT_BATCH_SIZE = 25
DEFAULT_WAIT_SECONDS = 300
DEFAULT_POLL_SECONDS = 3
GRID_COLS = 5
GRID_ROWS = 5
NEW_CHAT_URL = "https://chatgpt.com/"

REQUIRED_COLUMNS = ["sku_code", "sku_name"]


# ---------------------------------------------------------------------------
# SKU NAME HELPERS
# ---------------------------------------------------------------------------

def split_sku_name(sku_name: str) -> tuple[str, str]:
    """Split 'VoltBrand Tablets Model-0051 - Grey / 64 GB' into
    ('VoltBrand Tablets', 'Model-0051 - Grey / 64 GB').

    Text is never changed, only split. If there is no 'Model-' part the
    whole name is returned as the first line.
    """
    match = re.match(r"^(.*?)\s+(Model-.*)$", sku_name.strip())
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return sku_name.strip(), ""


def safe_filename_part(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())


# ---------------------------------------------------------------------------
# IMAGE GENERATION PROMPTS
# ---------------------------------------------------------------------------

COMMON_RULES = r"""
VISUAL STYLE:
- photorealistic professional ecommerce studio product photography
- realistic materials, proportions and lighting
- clean pure white background, soft grounding shadow
- consistent front three-quarter camera angle across all cells
- one product per cell, centered, fully visible, similar scale in every cell

PRODUCT RULES:
- Each record's SKU_NAME describes the physical product: brand, product type,
  model, color and configuration. The color MUST match exactly.
- Records sharing the same "Model-XXXX" are the SAME physical design; change
  only the explicitly stated color / configuration.
- The brand name may appear as a small logo on the product itself, spelled
  exactly as given.

DO NOT ADD: people, hands, lifestyle scenes, props, packaging, duplicate
products, illustrations, cartoons, banners, price tags, watermarks.
""".strip()


OVERLAY_LAYOUT = r"""
LAYOUT (STRICT):
- ONE square image containing exactly 5 columns x 5 rows = 25 EQUAL square cells.
- Thin light-grey divider lines between all cells, full width and full height.
- In every cell, place the product in the UPPER 75% of the cell.
- Leave the BOTTOM 25% of every cell completely EMPTY pure white
  (text labels will be added later by software).
- Do NOT write any text, captions, SKU codes, numbers or labels in the image
  (the only allowed text is the brand logo printed on a product).
- Cell order: left-to-right, top-to-bottom. Cell 01 = top-left,
  Cell 05 = top-right, Cell 21 = bottom-left, Cell 25 = bottom-right.
- If fewer than 25 records are given, leave the remaining cells empty white.
""".strip()


AI_LABEL_LAYOUT = r"""
LAYOUT (STRICT):
- ONE square image containing exactly 5 columns x 5 rows = 25 EQUAL square cells.
- Thin light-grey divider lines between all cells.
- In every cell: the product in the upper ~75%, and below it three centered
  text lines in a clean sans-serif font, black on white:
    line 1 (bold): the SKU_CODE
    line 2: LABEL_LINE_2
    line 3: LABEL_LINE_3
- Copy label text EXACTLY, character by character. Do not invent text.
- Cell order: left-to-right, top-to-bottom. Cell 01 = top-left,
  Cell 05 = top-right, Cell 21 = bottom-left, Cell 25 = bottom-right.
""".strip()


def build_prompt(
    rows: list[dict[str, str]],
    batch_number: int,
    label_mode: str,
) -> str:
    """Build a self-contained prompt containing the actual SKU records."""

    layout = OVERLAY_LAYOUT if label_mode == "overlay" else AI_LABEL_LAYOUT

    parts = [
        "Create an image: a square photorealistic ecommerce product catalog "
        f"contact sheet showing {len(rows)} products in a 5 x 5 grid. "
        "Use the image generation tool now; no reference image is needed "
        f"(batch {batch_number}).",
        "",
        COMMON_RULES,
        "",
        layout,
        "",
        "========== PRODUCTS (in exact cell order) ==========",
    ]

    for position, row in enumerate(rows, start=1):
        sku_code = row["sku_code"]
        sku_name = row["sku_name"]
        line2, line3 = split_sku_name(sku_name)

        if label_mode == "overlay":
            parts.append(f"Cell {position:02d}: {sku_name}")
        else:
            parts.extend([
                "",
                f"Cell {position:02d}:",
                f"  SKU_CODE: {sku_code}",
                f"  SKU_NAME: {sku_name}",
                f"  LABEL_LINE_2: {line2}",
                f"  LABEL_LINE_3: {line3}",
            ])

    parts.extend([
        "========== END PRODUCTS ==========",
        "",
        f"Create the image now with exactly {len(rows)} products.",
    ])

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# INPUT
# ---------------------------------------------------------------------------

def load_input(path: Path, sheet: str | int = 0) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(
            path,
            dtype=str,
            keep_default_na=False,
        )
    elif suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(
            path,
            sheet_name=sheet,
            dtype=str,
        ).fillna("")
    else:
        raise ValueError(
            f"Unsupported file: {suffix}. Use CSV/XLSX/XLS."
        )

    df.columns = [str(c).strip() for c in df.columns]

    missing = [
        c for c in REQUIRED_COLUMNS
        if c not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Missing columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df = df[
        (df["sku_code"] != "") &
        (df["sku_name"] != "")
    ].copy()

    if df.empty:
        raise ValueError("No valid sku_code + sku_name rows found.")

    return df


# ---------------------------------------------------------------------------
# BATCHING
# ---------------------------------------------------------------------------

def get_total_batches(row_count: int, batch_size: int) -> int:
    return (row_count + batch_size - 1) // batch_size


def get_batch(
    df: pd.DataFrame,
    batch_number: int,
    batch_size: int,
):
    start = (batch_number - 1) * batch_size
    end = min(start + batch_size, len(df))

    if start >= len(df):
        return None, start, end

    batch = df.iloc[start:end]

    records = [
        {
            "sku_code": str(row["sku_code"]).strip(),
            "sku_name": str(row["sku_name"]).strip(),
        }
        for _, row in batch.iterrows()
    ]

    return records, start, end


# ---------------------------------------------------------------------------
# MANIFEST
# ---------------------------------------------------------------------------

MANIFEST_COLUMNS = [
    "batch",
    "start_row",
    "end_row",
    "product_count",
    "sku_start",
    "sku_end",
    "image_file",
    "prompt_file",
    "status",
    "timestamp",
    "error",
]


def ensure_manifest(path: Path):
    if not path.exists():
        with path.open("w", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS).writeheader()


def read_manifest(path: Path):
    result = {}

    if not path.exists():
        return result

    with path.open("r", newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            try:
                result[int(row["batch"])] = row
            except Exception:
                pass

    return result


def append_manifest(path: Path, row: Dict):
    with path.open("a", newline="", encoding="utf-8") as f:
        csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS).writerow(row)


# ---------------------------------------------------------------------------
# GRID COMPOSITION (Pillow) - exact Excel text labels
# ---------------------------------------------------------------------------

FONT_REGULAR_CANDIDATES = [
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/segoeui.ttf",
    "/Library/Fonts/Arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
FONT_BOLD_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/segoeuib.ttf",
    "/Library/Fonts/Arial Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]

CELL_SIZE = 300          # final cell is CELL_SIZE x CELL_SIZE px
LABEL_HEIGHT = 84        # bottom strip for the 3 text lines
CELL_PADDING = 14
LINE_COLOR = (220, 220, 220)
TEXT_COLOR = (20, 20, 20)


def load_font(candidates: List[str], size: int):
    for font_path in candidates:
        if Path(font_path).exists():
            return ImageFont.truetype(font_path, size)
    return ImageFont.load_default()


def fit_font(draw, text, candidates, max_size, min_size, max_width):
    """Largest font size (max_size..min_size) that fits text in max_width."""
    for size in range(max_size, min_size - 1, -1):
        font = load_font(candidates, size)
        if draw.textlength(text, font=font) <= max_width:
            return font
    return load_font(candidates, min_size)


def find_boundaries(gray: np.ndarray, count: int, axis: int) -> List[int]:
    """Find grid divider positions along one axis.

    axis=1 -> vertical lines (x positions), axis=0 -> horizontal lines.
    Starts from an even split and snaps to a nearby divider line when one
    is detected (a line is a row/column where almost every pixel is
    non-white). Falls back to the even split.
    """
    length = gray.shape[1] if axis == 1 else gray.shape[0]
    non_white = gray < 245
    # Fraction of non-white pixels per column (axis=1) or row (axis=0).
    profile = non_white.mean(axis=0 if axis == 1 else 1)

    bounds = [0]
    window = max(4, length // (count * 8))
    for i in range(1, count):
        expected = round(i * length / count)
        lo = max(1, expected - window)
        hi = min(length - 1, expected + window)
        segment = profile[lo:hi]
        best = int(np.argmax(segment))
        bounds.append(lo + best if segment[best] >= 0.85 else expected)
    bounds.append(length)
    return bounds


def trim_to_product(cell: Image.Image) -> Image.Image:
    """Crop a cell to the bounding box of its non-white content."""
    gray = np.asarray(cell.convert("L"))
    mask = gray < 238
    # Ignore the outer 3% so divider lines are not treated as product.
    h, w = mask.shape
    edge_y, edge_x = max(2, h * 3 // 100), max(2, w * 3 // 100)
    mask[:edge_y, :] = False
    mask[-edge_y:, :] = False
    mask[:, :edge_x] = False
    mask[:, -edge_x:] = False

    ys, xs = np.where(mask)
    if len(xs) == 0:
        return cell
    pad = 4
    box = (
        max(0, xs.min() - pad),
        max(0, ys.min() - pad),
        min(w, xs.max() + pad),
        min(h, ys.max() + pad),
    )
    return cell.crop(box)


def split_raw_grid(raw: Image.Image) -> List[Image.Image]:
    """Split the raw 5x5 image into 25 product crops (row-major order)."""
    raw = raw.convert("RGB")
    gray = np.asarray(raw.convert("L"))
    xs = find_boundaries(gray, GRID_COLS, axis=1)
    ys = find_boundaries(gray, GRID_ROWS, axis=0)

    cells = []
    for r in range(GRID_ROWS):
        for c in range(GRID_COLS):
            cell = raw.crop((xs[c], ys[r], xs[c + 1], ys[r + 1]))
            cells.append(trim_to_product(cell))
    return cells


def compose_labeled_grid(
    raw_path: Path,
    records: list[dict[str, str]],
    output_path: Path,
    cells_dir: Path | None = None,
) -> None:
    """Build the final contact sheet: product crops from the AI image +
    exact SKU labels from the input file."""
    raw = Image.open(raw_path)
    products = split_raw_grid(raw)

    width = GRID_COLS * CELL_SIZE
    height = GRID_ROWS * CELL_SIZE
    sheet = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(sheet)

    product_box_w = CELL_SIZE - 2 * CELL_PADDING
    product_box_h = CELL_SIZE - LABEL_HEIGHT - CELL_PADDING
    text_max_w = CELL_SIZE - 10

    for index, record in enumerate(records[: GRID_COLS * GRID_ROWS]):
        r, c = divmod(index, GRID_COLS)
        x0, y0 = c * CELL_SIZE, r * CELL_SIZE

        product = products[index].copy()
        product.thumbnail((product_box_w, product_box_h), Image.LANCZOS)
        px = x0 + (CELL_SIZE - product.width) // 2
        py = y0 + CELL_PADDING + (product_box_h - product.height) // 2
        sheet.paste(product, (px, py))

        if cells_dir is not None:
            products[index].save(
                cells_dir / f"{safe_filename_part(record['sku_code'])}.png"
            )

        line2, line3 = split_sku_name(record["sku_name"])
        lines = [
            (record["sku_code"], FONT_BOLD_CANDIDATES, 20),
            (line2, FONT_REGULAR_CANDIDATES, 17),
            (line3, FONT_REGULAR_CANDIDATES, 17),
        ]
        ty = y0 + CELL_SIZE - LABEL_HEIGHT + 4
        for text, candidates, size in lines:
            if not text:
                continue
            font = fit_font(draw, text, candidates, size, 10, text_max_w)
            tw = draw.textlength(text, font=font)
            draw.text((x0 + (CELL_SIZE - tw) / 2, ty), text, font=font, fill=TEXT_COLOR)
            ty += 25

    for i in range(1, GRID_COLS):
        draw.line([(i * CELL_SIZE, 0), (i * CELL_SIZE, height)], fill=LINE_COLOR, width=1)
    for i in range(1, GRID_ROWS):
        draw.line([(0, i * CELL_SIZE), (width, i * CELL_SIZE)], fill=LINE_COLOR, width=1)

    sheet.save(output_path)


# ---------------------------------------------------------------------------
# CHATGPT UI AUTOMATION
# ---------------------------------------------------------------------------

def find_prompt_box(page: Page):
    selectors = [
        '#prompt-textarea',
        '[contenteditable="true"][role="textbox"]',
        'textarea[placeholder*="Message"]',
        'textarea[placeholder*="message"]',
        '[contenteditable="true"]',
        'textarea',
    ]

    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.is_visible(timeout=1500):
                return locator
        except Exception:
            continue

    return None


# ---------------------------------------------------------------------------
# HUMAN VERIFICATION / CAPTCHA HANDLING
# ---------------------------------------------------------------------------

CAPTCHA_TEXT_MARKERS = [
    "are you human",
    "verify you are human",
    "verify you're human",
    "human verification",
    "security check",
    "captcha",
    "cf-chl",
    "checking your browser",
    "verify that you are human",
]


def is_human_verification_page(page: Page) -> bool:
    """Detect a visible human-verification/CAPTCHA challenge.

    This function ONLY detects the challenge. It never attempts to solve or
    bypass it. The user must complete the verification manually in the
    visible browser window.
    """
    try:
        url = (page.url or "").lower()
        if any(marker in url for marker in ("captcha", "challenge", "verify")):
            return True
    except Exception:
        pass

    try:
        text = page.locator("body").inner_text(timeout=1500).lower()
        return any(marker in text for marker in CAPTCHA_TEXT_MARKERS)
    except Exception:
        return False


def pause_for_manual_verification(page: Page, context: str = "") -> None:
    """Pause until the user manually completes human verification."""
    logging.warning("")
    logging.warning("=" * 70)
    logging.warning("HUMAN VERIFICATION / CAPTCHA DETECTED")
    logging.warning("=" * 70)
    if context:
        logging.warning("Context: %s", context)
    logging.warning("The script will NOT attempt to solve or bypass the CAPTCHA.")
    logging.warning("Please solve the verification manually in the browser window.")
    logging.warning("When ChatGPT is accessible again, return to this terminal.")
    logging.warning("Press ENTER to continue the current batch.")
    logging.warning("=" * 70)

    input()

    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(2000)

    while is_human_verification_page(page):
        logging.warning("CAPTCHA/human verification is still present.")
        logging.warning("Complete it in the browser, then press ENTER again.")
        input()
        page.wait_for_timeout(2000)

    logging.info("Human verification cleared. Resuming current batch.")


def wait_for_composer_with_manual_verification(page: Page, timeout_seconds: int = 300):
    """Find the composer, pausing for manual CAPTCHA verification when needed."""
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        if is_human_verification_page(page):
            pause_for_manual_verification(page, "Waiting for ChatGPT composer")
            deadline = time.time() + timeout_seconds
            continue

        box = find_prompt_box(page)
        if box is not None:
            return box

        time.sleep(2)

    raise RuntimeError(
        "Could not find the ChatGPT message box. Log in manually in the "
        "opened browser window and run again."
    )


def wait_for_login(page: Page):
    """Wait until ChatGPT is usable, with manual CAPTCHA/login handoff."""
    logging.info("Checking ChatGPT session/composer...")
    wait_for_composer_with_manual_verification(page, timeout_seconds=300)
    logging.info("ChatGPT session is ready.")


def send_prompt(page: Page, prompt: str):
    """Enter and submit a prompt; pause for manual CAPTCHA if encountered."""
    while True:
        if is_human_verification_page(page):
            pause_for_manual_verification(page, "Before submitting prompt")

        box = find_prompt_box(page)

        if box is None:
            box = wait_for_composer_with_manual_verification(
                page,
                timeout_seconds=120,
            )

        logging.info("Entering prompt...")

        try:
            box.click()

            try:
                box.fill(prompt)
            except Exception:
                page.evaluate(
                    """async (text) => {
                        await navigator.clipboard.writeText(text);
                    }""",
                    prompt,
                )
                box.press("Control+V")

            time.sleep(1.5)

            send_button = page.locator('[data-testid="send-button"]').first
            try:
                if send_button.is_visible(timeout=1500) and send_button.is_enabled():
                    send_button.click()
                else:
                    box.press("Enter")
            except Exception:
                box.press("Enter")

            logging.info("Prompt submitted.")
            return

        except Exception as exc:
            if is_human_verification_page(page):
                logging.warning("Human verification appeared while submitting.")
                pause_for_manual_verification(page, "While submitting prompt")
                continue
            raise exc


ASSISTANT_SELECTORS = [
    '[data-message-author-role="assistant"]',
    'main article',
    'article',
]


def get_last_assistant(page: Page):
    """Return a locator for the newest assistant message, or None."""
    for selector in ASSISTANT_SELECTORS:
        try:
            loc = page.locator(selector)
            count = loc.count()
            if count:
                return loc.nth(count - 1)
        except Exception:
            continue
    return None


def get_last_assistant_signature(page: Page) -> str:
    """Return a compact DOM signature for the newest assistant message."""
    node = get_last_assistant(page)
    if node is None:
        return ""
    try:
        return node.inner_html(timeout=1000)[-12000:]
    except Exception:
        return ""


def find_largest_image(page: Page):
    """Largest visible generated image in the newest response.

    Generated images are sometimes rendered outside the assistant message
    container, so fall back to the last big image on the page.
    """
    scopes = []
    node = get_last_assistant(page)
    if node is not None:
        scopes.append(node)
    scopes.append(page.locator("main"))

    for scope in scopes:
        best, best_area = None, 0
        try:
            images = scope.locator("img")
            for i in range(images.count()):
                img = images.nth(i)
                try:
                    if not img.is_visible(timeout=300):
                        continue
                    box = img.bounding_box()
                    if not box or box["width"] < 250 or box["height"] < 250:
                        continue
                    area = box["width"] * box["height"]
                    # Prefer later images when equal (newest response).
                    if area >= best_area:
                        best, best_area = img, area
                except Exception:
                    continue
        except Exception:
            continue
        if best is not None:
            return best
    return None


def last_assistant_has_image(page: Page) -> bool:
    return find_largest_image(page) is not None


def is_generating(page: Page) -> bool:
    """True while ChatGPT is still streaming / creating the image."""
    selectors = [
        '[data-testid="stop-button"]',
        'button[aria-label*="Stop"]',
    ]
    for selector in selectors:
        try:
            if page.locator(selector).first.is_visible(timeout=300):
                return True
        except Exception:
            continue
    return False


RATE_LIMIT_MARKERS = [
    "reached your limit",
    "reached the limit",
    "hit the limit",
    "image generation limit",
    "image creation limit",
    "limit for image",
    "try again later",
    "try again in",
    "too many requests",
    "usage cap",
]

IN_PROGRESS_MARKERS = [
    "creating image",
    "generating image",
    "getting started",
]

RETRY_MESSAGE = (
    "Please create the image now with the image generation tool, exactly as "
    "described in my previous message. Reply with the image only, no text."
)


def get_last_assistant_text(page: Page) -> str:
    node = get_last_assistant(page)
    if node is None:
        return ""
    try:
        return node.inner_text(timeout=1000).strip()
    except Exception:
        return ""


def select_image_tool(page: Page) -> bool:
    """Best-effort: turn on the composer's 'Create image' tool so ChatGPT
    must generate an image instead of answering in text.

    ChatGPT's UI changes often, so failure here is not fatal.
    """
    try:
        plus = page.locator(
            '[data-testid="composer-plus-btn"], '
            'button[aria-label*="Add photos"], '
            'button[aria-label*="Add files"]'
        ).first
        if not plus.is_visible(timeout=1500):
            return False
        plus.click()
        page.wait_for_timeout(800)

        item = page.get_by_role(
            "menuitem", name=re.compile(r"create image", re.I)
        ).first
        if not item.is_visible(timeout=1500):
            item = page.get_by_text(re.compile(r"^create image$", re.I)).first
        if item.is_visible(timeout=1000):
            item.click()
            page.wait_for_timeout(500)
            logging.info("'Create image' tool selected.")
            return True

        page.keyboard.press("Escape")
    except Exception as exc:
        logging.debug("Could not select image tool: %s", exc)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
    logging.info("'Create image' tool not found; relying on the prompt text.")
    return False


def wait_for_new_response_and_image(
    page: Page,
    previous_assistant_signature: str,
    timeout_seconds: int,
) -> str:
    """Wait for a NEW assistant response.

    Returns:
        "image"        - a finished image is present
        "text_only"    - ChatGPT finished answering with text and no image
        "rate_limited" - ChatGPT reports an image/usage limit
        "timeout"      - nothing usable before the deadline
    """
    deadline = time.time() + timeout_seconds
    logging.info(
        "Waiting for NEW ChatGPT response and generated image (up to %d seconds)...",
        timeout_seconds,
    )

    response_changed = False
    last_src = None
    stable_polls = 0
    last_text = None
    text_stable_polls = 0

    while time.time() < deadline:
        if is_human_verification_page(page):
            pause_for_manual_verification(
                page,
                "While waiting for the current batch response",
            )
            deadline = time.time() + timeout_seconds
            continue

        current_signature = get_last_assistant_signature(page)

        if current_signature and current_signature != previous_assistant_signature:
            if not response_changed:
                logging.info("New assistant response detected.")
            response_changed = True

        if response_changed and not is_generating(page):
            img = find_largest_image(page)
            if img is not None:
                try:
                    src = img.get_attribute("src") or ""
                except Exception:
                    src = ""
                # The image src changes while the preview is progressively
                # rendered; wait until it stays the same for two polls.
                if src and src == last_src:
                    stable_polls += 1
                else:
                    stable_polls = 0
                last_src = src
                if stable_polls >= 2:
                    logging.info("Generated image for the current batch detected.")
                    return "image"
            else:
                text = get_last_assistant_text(page)
                lowered = text.lower()
                if any(marker in lowered for marker in RATE_LIMIT_MARKERS):
                    logging.warning("ChatGPT reports a limit: %s", text[:300])
                    return "rate_limited"

                # A text answer that stays unchanged for ~15 s after the
                # stop button disappeared means no image is coming.
                if text and not any(m in lowered for m in IN_PROGRESS_MARKERS):
                    if text == last_text:
                        text_stable_polls += 1
                    else:
                        text_stable_polls = 0
                    last_text = text
                    if text_stable_polls * DEFAULT_POLL_SECONDS >= 15:
                        logging.warning(
                            "ChatGPT answered with text only: %s", text[:300]
                        )
                        return "text_only"

        time.sleep(DEFAULT_POLL_SECONDS)

    return "timeout"


def save_bytes_as_png(data: bytes, output_path: Path) -> None:
    image = Image.open(io.BytesIO(data)).convert("RGB")
    image.save(output_path, format="PNG")
    logging.info("Saved image %s (%dx%d)", output_path, image.width, image.height)


def click_download_button(page: Page, output_path: Path) -> bool:
    """Hover the newest image and click ChatGPT's Download button."""
    img = find_largest_image(page)
    if img is None:
        return False

    download_name = re.compile(r"download", re.I)

    def try_click(scope) -> bool:
        button = scope.get_by_role("button", name=download_name).last
        try:
            if not button.is_visible(timeout=1500):
                return False
        except Exception:
            return False
        with page.expect_download(timeout=60000) as download_info:
            button.click()
        tmp_path = output_path.with_suffix(".download")
        download_info.value.save_as(str(tmp_path))
        save_bytes_as_png(tmp_path.read_bytes(), output_path)
        tmp_path.unlink(missing_ok=True)
        return True

    try:
        img.hover()
        page.wait_for_timeout(1000)
        node = get_last_assistant(page)
        if node is not None and try_click(node):
            return True

        # Open the image viewer; it has its own Download button.
        img.click()
        page.wait_for_timeout(1500)
        ok = try_click(page)
        page.keyboard.press("Escape")
        return ok
    except Exception as exc:
        logging.warning("Download button approach failed: %s", exc)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return False


def save_image_from_visible_image(page: Page, output_path: Path) -> bool:
    """Fallback: read the image bytes from its src, or screenshot it."""
    img = find_largest_image(page)
    if img is None:
        return False

    try:
        src = img.get_attribute("src") or ""

        if src.startswith("http"):
            # context.request carries the logged-in cookies.
            response = page.context.request.get(src, timeout=60000)
            if response.ok:
                save_bytes_as_png(response.body(), output_path)
                return True

        if src.startswith(("blob:", "data:")):
            b64 = page.evaluate(
                """async (src) => {
                    const blob = await (await fetch(src)).blob();
                    return await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result.split(',')[1]);
                        reader.readAsDataURL(blob);
                    });
                }""",
                src,
            )
            save_bytes_as_png(base64.b64decode(b64), output_path)
            return True
    except Exception as exc:
        logging.warning("Reading image src failed: %s", exc)

    try:
        # Last resort: element screenshot (lower resolution than original).
        img.screenshot(path=str(output_path))
        logging.warning("Saved image via element screenshot (reduced quality).")
        return True
    except Exception as exc:
        logging.warning("Element screenshot failed: %s", exc)
        return False


def save_debug_screenshot(page: Page, output_dir: Path, batch_number: int) -> None:
    debug_dir = output_dir / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    path = debug_dir / f"batch_{batch_number:05d}.png"
    try:
        page.screenshot(path=str(path))
        logging.info("Debug screenshot saved: %s", path)
    except Exception as exc:
        logging.warning("Could not save debug screenshot: %s", exc)


def generate_and_save_image(
    page: Page,
    prompt: str,
    output_path: Path,
    timeout_seconds: int,
    retries: int,
) -> str:
    """Send the prompt, re-ask if ChatGPT replies with text only, and save
    the image.

    Returns "success", "failed_download", "text_only", "rate_limited" or
    "timeout".
    """
    select_image_tool(page)
    previous_signature = get_last_assistant_signature(page)
    send_prompt(page, prompt)

    for attempt in range(retries + 1):
        result = wait_for_new_response_and_image(
            page,
            previous_signature,
            timeout_seconds,
        )

        if result == "image":
            # Give the UI a moment to finish rendering image controls.
            time.sleep(3)
            if click_download_button(page, output_path):
                return "success"
            if save_image_from_visible_image(page, output_path):
                return "success"
            logging.error("Image was visible but could not be downloaded automatically.")
            return "failed_download"

        if result != "text_only" or attempt == retries:
            return result

        logging.info(
            "Asking ChatGPT again to generate the image (retry %d/%d)...",
            attempt + 1,
            retries,
        )
        time.sleep(5)
        select_image_tool(page)
        previous_signature = get_last_assistant_signature(page)
        send_prompt(page, RETRY_MESSAGE)

    return "text_only"


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def batch_paths(records, image_dir: Path, raw_dir: Path, prompt_dir: Path):
    stem = (
        f"{safe_filename_part(records[0]['sku_code'])}_to_"
        f"{safe_filename_part(records[-1]['sku_code'])}"
    )
    return (
        image_dir / f"{stem}.png",
        raw_dir / f"{stem}.png",
        prompt_dir / f"{stem}.txt",
    )


def manifest_row(batch_number, start_index, end_index, records,
                 image_path, prompt_path, status, error=""):
    return {
        "batch": batch_number,
        "start_row": start_index + 1,
        "end_row": end_index,
        "product_count": len(records),
        "sku_start": records[0]["sku_code"],
        "sku_end": records[-1]["sku_code"],
        "image_file": str(image_path),
        "prompt_file": str(prompt_path),
        "status": status,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "error": error,
    }


def process(
    input_path: Path,
    sheet: str | int,
    chat_url: str | None,
    output_dir: Path,
    profile_dir: Path,
    batch_size: int,
    start_batch: int,
    end_batch: int | None,
    wait_seconds: int,
    force: bool,
    slow_mode: float,
    connect_existing: bool,
    cdp_url: str,
    label_mode: str,
    save_cells: bool,
    compose_only: bool,
    retries: int,
):
    df = load_input(input_path, sheet)

    total = get_total_batches(len(df), batch_size)
    end_batch = min(end_batch or total, total)

    image_dir = output_dir / "images"
    raw_dir = output_dir / "raw"
    prompt_dir = output_dir / "prompts"
    cells_dir = output_dir / "cells" if save_cells else None
    for folder in (image_dir, raw_dir, prompt_dir, cells_dir):
        if folder is not None:
            folder.mkdir(parents=True, exist_ok=True)

    manifest_path = output_dir / "manifest.csv"
    ensure_manifest(manifest_path)
    manifest = read_manifest(manifest_path)

    new_chat_per_batch = not chat_url
    open_url = chat_url or NEW_CHAT_URL

    logging.info("Input: %s", input_path)
    logging.info("Records: %d", len(df))
    logging.info("Batch size: %d", batch_size)
    logging.info("Total batches: %d", total)
    logging.info("Processing batches %d-%d", start_batch, end_batch)
    logging.info("Label mode: %s", label_mode)
    logging.info(
        "Session: %s",
        "NEW chat per batch" if new_chat_per_batch else f"reuse {chat_url}",
    )

    def finalize(records, raw_path, image_path):
        if label_mode == "overlay":
            compose_labeled_grid(raw_path, records, image_path, cells_dir)
            logging.info("Final labeled grid saved: %s", image_path)
        else:
            image_path.write_bytes(raw_path.read_bytes())

    if compose_only:
        for batch_number in range(start_batch, end_batch + 1):
            records, _, _ = get_batch(df, batch_number, batch_size)
            if not records:
                continue
            image_path, raw_path, _ = batch_paths(records, image_dir, raw_dir, prompt_dir)
            if raw_path.exists():
                finalize(records, raw_path, image_path)
            else:
                logging.warning("No raw image for batch %d: %s", batch_number, raw_path)
        return

    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        if connect_existing:
            logging.info("Connecting to existing Chrome via CDP: %s", cdp_url)
            try:
                browser = p.chromium.connect_over_cdp(cdp_url)
            except Exception as exc:
                raise RuntimeError(
                    f"Could not connect to existing Chrome at {cdp_url}. "
                    "Start Chrome with remote debugging enabled, keep it open, "
                    "and make sure ChatGPT is logged in. Original error: "
                    f"{exc}"
                ) from exc

            if not browser.contexts:
                raise RuntimeError("Connected to Chrome, but no browser context was found.")
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else context.new_page()
        else:
            logging.info("Opening persistent Chromium profile: %s", profile_dir)
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir),
                headless=False,
                viewport={"width": 1440, "height": 1000},
                accept_downloads=True,
            )
            page = context.pages[0] if context.pages else context.new_page()

        logging.info("Opening ChatGPT...")
        page.goto(open_url, wait_until="domcontentloaded", timeout=120000)
        page.wait_for_timeout(5000)
        wait_for_login(page)

        for batch_number in range(start_batch, end_batch + 1):
            records, start_index, end_index = get_batch(df, batch_number, batch_size)
            if not records:
                continue

            image_path, raw_path, prompt_path = batch_paths(
                records, image_dir, raw_dir, prompt_dir
            )

            previous = manifest.get(batch_number)
            if (
                not force
                and image_path.exists()
                and previous
                and previous.get("status") == "success"
            ):
                logging.info(
                    "[%d/%d] SKIP already completed: %s",
                    batch_number, total, image_path,
                )
                continue

            logging.info(
                "\n==================================================\n"
                "BATCH %d/%d\n"
                "Rows: %d-%d\n"
                "SKU: %s -> %s\n"
                "Output: %s\n"
                "==================================================",
                batch_number, total, start_index + 1, end_index,
                records[0]["sku_code"], records[-1]["sku_code"], image_path.name,
            )

            # Each batch gets its own tab. With no --chat-url this is a brand
            # new chat, so earlier batches can never leak into this image.
            batch_page = context.new_page()
            batch_page.goto(open_url, wait_until="domcontentloaded", timeout=120000)
            batch_page.wait_for_timeout(3000)

            prompt = build_prompt(records, batch_number, label_mode)
            prompt_path.write_text(prompt, encoding="utf-8")

            stop_run = False
            try:
                wait_for_login(batch_page)

                status = generate_and_save_image(
                    batch_page,
                    prompt,
                    raw_path,
                    wait_seconds,
                    retries,
                )

                if new_chat_per_batch:
                    logging.info("Chat for this batch: %s", batch_page.url)

                if status == "success":
                    finalize(records, raw_path, image_path)
                    append_manifest(manifest_path, manifest_row(
                        batch_number, start_index, end_index, records,
                        image_path, prompt_path, "success",
                    ))
                    logging.info("SUCCESS: batch %d", batch_number)
                else:
                    errors = {
                        "failed_download": "Generated image was not automatically downloaded.",
                        "text_only": "ChatGPT replied with text only (no image) after retries.",
                        "rate_limited": "ChatGPT image/usage limit reached.",
                        "timeout": "No image before --wait-seconds.",
                    }
                    save_debug_screenshot(batch_page, output_dir, batch_number)
                    append_manifest(manifest_path, manifest_row(
                        batch_number, start_index, end_index, records,
                        image_path, prompt_path, status, errors.get(status, ""),
                    ))
                    logging.error("FAILED (%s): batch %d", status, batch_number)

                    if status == "rate_limited":
                        logging.error(
                            "Stopping the run: ChatGPT limit reached. Re-run "
                            "later; completed batches are skipped automatically."
                        )
                        stop_run = True

            except Exception as exc:
                logging.exception("Batch %d failed: %s", batch_number, exc)
                save_debug_screenshot(batch_page, output_dir, batch_number)
                append_manifest(manifest_path, manifest_row(
                    batch_number, start_index, end_index, records,
                    image_path, prompt_path, "failed", str(exc),
                ))

            finally:
                try:
                    batch_page.close()
                except Exception:
                    pass

            if stop_run:
                break

            if slow_mode > 0:
                logging.info("Waiting %.1f seconds before next batch...", slow_mode)
                time.sleep(slow_mode)

        logging.info("All requested batches processed.")
        time.sleep(3)

        if not connect_existing:
            context.close()
        else:
            logging.info("Leaving attached Chrome session open.")


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Generate 5x5 SKU product grids using the ChatGPT UI "
            "without an API key."
        )
    )

    parser.add_argument("--input", required=True, help="CSV/XLSX/XLS input file.")
    parser.add_argument(
        "--sheet",
        default="0",
        help="Excel sheet name or index. Default: first sheet.",
    )
    parser.add_argument(
        "--chat-url",
        default=None,
        help=(
            "Optional existing ChatGPT conversation URL to reuse for every "
            "batch. If omitted (recommended), a NEW chat is opened per batch."
        ),
    )
    parser.add_argument("--output", default="./product_grids", help="Output directory.")
    parser.add_argument(
        "--profile",
        default="./chatgpt_browser_profile",
        help="Persistent browser profile directory (used in normal launch mode).",
    )
    parser.add_argument(
        "--connect-existing",
        action="store_true",
        help="Attach to an already-running Chrome via CDP instead of launching one.",
    )
    parser.add_argument(
        "--cdp-url",
        default="http://127.0.0.1:9222",
        help="Chrome DevTools Protocol endpoint. Default: http://127.0.0.1:9222",
    )
    parser.add_argument(
        "--label-mode",
        choices=["overlay", "ai"],
        default="overlay",
        help=(
            "overlay (default): AI draws products only, labels are drawn "
            "from the input file with exact text. ai: ask ChatGPT to render "
            "the labels itself (may misspell)."
        ),
    )
    parser.add_argument(
        "--save-cells",
        action="store_true",
        help="Also save one cropped product image per SKU in <output>/cells.",
    )
    parser.add_argument(
        "--compose-only",
        action="store_true",
        help="Do not open the browser; rebuild final grids from <output>/raw.",
    )
    parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE,
                        help="Products per image. Default: 25.")
    parser.add_argument("--start-batch", type=int, default=1, help="First batch. Default: 1.")
    parser.add_argument("--end-batch", type=int, default=None,
                        help="Last batch. Default: all batches.")
    parser.add_argument("--wait-seconds", type=int, default=DEFAULT_WAIT_SECONDS,
                        help="Maximum wait for image generation. Default: 300.")
    parser.add_argument("--slow-mode", type=float, default=10,
                        help="Delay between batches. Default: 10 seconds.")
    parser.add_argument("--retries", type=int, default=2,
                        help="Re-ask this many times if ChatGPT replies with text only. Default: 2.")
    parser.add_argument("--force", action="store_true",
                        help="Regenerate batches already marked successful.")

    return parser.parse_args()


def main():
    args = parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    if args.batch_size != GRID_COLS * GRID_ROWS:
        logging.error("Batch size must be 25 for a 5x5 grid.")
        sys.exit(2)

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    try:
        process(
            input_path=Path(args.input),
            sheet=sheet,
            chat_url=args.chat_url,
            output_dir=Path(args.output),
            profile_dir=Path(args.profile),
            batch_size=args.batch_size,
            start_batch=args.start_batch,
            end_batch=args.end_batch,
            wait_seconds=args.wait_seconds,
            force=args.force,
            slow_mode=args.slow_mode,
            connect_existing=args.connect_existing,
            cdp_url=args.cdp_url,
            label_mode=args.label_mode,
            save_cells=args.save_cells,
            compose_only=args.compose_only,
            retries=args.retries,
        )

    except KeyboardInterrupt:
        logging.warning("Stopped by user.")
        sys.exit(130)

    except Exception as exc:
        logging.exception("Fatal error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
