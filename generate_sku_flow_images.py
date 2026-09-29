#!/usr/bin/env python3
"""
generate_sku_flow_images.py

Automated workflow for generating product catalog images using Google Flow:
1. Reads product title and SKU ID from a CSV or Excel file.
2. Groups products into batches of 10 (arranged in a 2x5 grid).
3. Automates Google Flow (https://labs.google/fx/tools/flow) via Playwright:
   - Uses a persistent browser profile (or attaches to existing Chrome via CDP).
   - Prompts Google Flow to generate a clean 2x5 catalog grid containing all 10 products.
   - Downloads the high-resolution composite grid image.
4. Splits the composite image into 10 individual product images.
5. Saves each product image named by its SKU ID (e.g., SKU123.png or SKU123.jpg).
6. Preserves strict 1-to-1 mapping between product titles, SKU IDs, and output images.
7. Logs detailed batch execution history in manifest.csv.

LOGIN & CREDENTIALS:
-------------------
You do NOT need to hardcode Google credentials in this script.
Google blocks automated form-filling on login pages for security.
Instead, this script uses a persistent Chromium profile (./google_flow_profile)
or connects to your already-running Google Chrome via CDP (--connect-existing).
You log in manually once in the browser window, and your session is automatically
saved and reused for all future runs.

INSTALLATION:
------------
    pip install playwright pillow pandas openpyxl
    playwright install chromium

USAGE EXAMPLES:
--------------
1. Standard run (first 2 batches of 10, PNG format):
    python generate_sku_flow_images.py \\
        --input "sku_master_with_description.csv" \\
        --output-dir "./output_images" \\
        --start-batch 1 \\
        --end-batch 2

2. Output as JPG format:
    python generate_sku_flow_images.py \\
        --input "sku_master_with_description.csv" \\
        --format jpg

3. Custom SKU and Title columns:
    python generate_sku_flow_images.py \\
        --input "products.csv" \\
        --sku-col "SKU" \\
        --title-col "Title"

4. Connect to an already-open Chrome browser (where you are already logged in):
    First start Chrome in terminal:
        google-chrome --remote-debugging-port=9222
    Then run:
        python generate_sku_flow_images.py --connect-existing

5. Re-split existing downloaded grid images without opening the browser:
    python generate_sku_flow_images.py --split-only

6. Dry-run (validate CSV, view batch mappings & prompts without generating):
    python generate_sku_flow_images.py --dry-run
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
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageOps


# ---------------------------------------------------------------------------
# CONSTANTS & DEFAULTS
# ---------------------------------------------------------------------------

DEFAULT_BATCH_SIZE = 10
DEFAULT_GRID_ROWS = 2
DEFAULT_GRID_COLS = 5
DEFAULT_FLOW_URL = "https://labs.google/fx/tools/flow"
DEFAULT_WAIT_SECONDS = 180
DEFAULT_SLOW_MODE = 5.0
DEFAULT_POLL_INTERVAL = 3.0

DEFAULT_SKU_COL = "sku_code"
DEFAULT_TITLE_COL = "sku_name"

MANIFEST_FIELDNAMES = [
    "batch_number",
    "position_in_batch",
    "grid_row",
    "grid_col",
    "sku_id",
    "product_title",
    "product_image_file",
    "raw_grid_file",
    "status",
    "timestamp",
    "error",
]


# ---------------------------------------------------------------------------
# UTILITIES & STRING HELPERS
# ---------------------------------------------------------------------------

def safe_filename(value: str) -> str:
    """Sanitize SKU ID for safe filesystem naming across OS platforms."""
    cleaned = re.sub(r'[\\/*?:"<>|]+', "_", value.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned.strip("._") or "UNKNOWN_SKU"


# ---------------------------------------------------------------------------
# DATA INGESTION & BATCHING
# ---------------------------------------------------------------------------

def load_products(
    path: Path,
    sku_col: str = DEFAULT_SKU_COL,
    title_col: str = DEFAULT_TITLE_COL,
    sheet: str | int = 0,
) -> List[Dict[str, str]]:
    """Load product records from CSV or Excel file."""
    if not path.exists():
        raise FileNotFoundError(f"Input catalog file not found: {path}")

    suffix = path.suffix.lower()
    records: List[Dict[str, str]] = []

    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames is None:
                raise ValueError(f"CSV file '{path}' is empty or invalid.")

            # Case-insensitive column matching
            field_map = {col.strip().lower(): col.strip() for col in reader.fieldnames if col}
            actual_sku = field_map.get(sku_col.lower())
            actual_title = field_map.get(title_col.lower())

            if not actual_sku or not actual_title:
                available = list(reader.fieldnames)
                raise ValueError(
                    f"Required columns not found! Looking for SKU: '{sku_col}', Title: '{title_col}'.\n"
                    f"Available columns in file: {available}"
                )

            for idx, row in enumerate(reader, start=1):
                sku = str(row.get(actual_sku, "")).strip()
                title = str(row.get(actual_title, "")).strip()
                if sku and title:
                    records.append({
                        "sku": sku,
                        "title": title,
                        "row_num": str(idx),
                    })

    elif suffix in {".xlsx", ".xls"}:
        try:
            import pandas as pd
        except ImportError:
            raise ImportError("Pandas is required to read Excel files. Run: pip install pandas openpyxl")

        df = pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")
        field_map = {str(col).strip().lower(): str(col).strip() for col in df.columns}
        actual_sku = field_map.get(sku_col.lower())
        actual_title = field_map.get(title_col.lower())

        if not actual_sku or not actual_title:
            raise ValueError(
                f"Required columns not found in Excel! Looking for SKU: '{sku_col}', Title: '{title_col}'.\n"
                f"Available columns: {list(df.columns)}"
            )

        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            sku = str(row[actual_sku]).strip()
            title = str(row[actual_title]).strip()
            if sku and title:
                records.append({
                    "sku": sku,
                    "title": title,
                    "row_num": str(idx),
                })
    else:
        raise ValueError(f"Unsupported catalog file format '{suffix}'. Use .csv, .xlsx, or .xls.")

    if not records:
        raise ValueError(f"No valid rows with both '{sku_col}' and '{title_col}' found in '{path}'.")

    logging.info("Loaded %d valid products from %s", len(records), path)
    return records


def get_batch(
    records: List[Dict[str, str]],
    batch_number: int,
    batch_size: int = DEFAULT_BATCH_SIZE,
) -> Tuple[List[Dict[str, str]], int, int]:
    """Retrieve items for a specific 1-indexed batch number."""
    start_idx = (batch_number - 1) * batch_size
    end_idx = min(start_idx + batch_size, len(records))

    if start_idx >= len(records):
        return [], start_idx, end_idx

    return records[start_idx:end_idx], start_idx, end_idx


# ---------------------------------------------------------------------------
# PROMPT CONSTRUCTION
# ---------------------------------------------------------------------------

def build_grid_prompt(
    batch_records: List[Dict[str, str]],
    rows: int = DEFAULT_GRID_ROWS,
    cols: int = DEFAULT_GRID_COLS,
) -> str:
    """Build a detailed prompt instructing Google Flow to create a grid of products."""
    total_slots = rows * cols
    prompt_lines = [
        f"Create a single high-resolution photorealistic ecommerce product catalog contact sheet.",
        f"The image must be composed of a strict {rows} rows by {cols} columns grid ({total_slots} equal rectangular cells total).",
        "",
        "STRICT VISUAL STYLE & GUIDELINES:",
        "- Photorealistic professional commercial studio product photography.",
        "- Each cell must have a clean, seamless, pure white background with soft realistic contact shadow underneath.",
        "- Exactly ONE product per cell, centered, perfectly in focus, full product visible.",
        "- Consistent 3/4 front studio angle and uniform scale across all cells.",
        "- Thin, subtle light-grey divider lines between all cells so each product boundary is distinct.",
        "- DO NOT include people, models, hands, lifestyle scenery, props, boxes, price tags, text overlays, or watermarks.",
        "",
        "PRODUCTS IN EXACT CELL POSITIONS (Row-by-Row, Left-to-Right):",
    ]

    for index, record in enumerate(batch_records, start=1):
        r = (index - 1) // cols + 1
        c = (index - 1) % cols + 1
        prompt_lines.append(f"Cell {index:02d} (Row {r}, Col {c}): {record['title']}")

    # If the batch has fewer items than grid capacity
    if len(batch_records) < total_slots:
        for index in range(len(batch_records) + 1, total_slots + 1):
            r = (index - 1) // cols + 1
            c = (index - 1) % cols + 1
            prompt_lines.append(f"Cell {index:02d} (Row {r}, Col {c}): [Leave completely empty pure white]")

    prompt_lines.extend([
        "",
        f"Generate the image now showing all {len(batch_records)} items arranged cleanly in their exact cell positions."
    ])

    return "\n".join(prompt_lines)


# ---------------------------------------------------------------------------
# IMAGE SPLITTING & PROCESSING (Pillow)
# ---------------------------------------------------------------------------

def split_composite_grid(
    raw_image_path: Path,
    batch_records: List[Dict[str, str]],
    output_products_dir: Path,
    rows: int = DEFAULT_GRID_ROWS,
    cols: int = DEFAULT_GRID_COLS,
    image_format: str = "png",
    trim_border_pct: float = 0.02,
) -> List[Dict[str, Any]]:
    """Split the composite grid image into individual product images.

    Preserves exact 1-to-1 mapping:
        Cell index -> Record -> SKU ID -> Filename.
    """
    image_format = image_format.lower()
    ext = ".jpg" if image_format in {"jpg", "jpeg"} else ".png"

    try:
        raw_img = Image.open(raw_image_path)
    except Exception as exc:
        raise RuntimeError(f"Failed to open downloaded raw image '{raw_image_path}': {exc}") from exc

    img_w, img_h = raw_img.size
    cell_w = img_w / cols
    cell_h = img_h / rows

    results = []

    for index, record in enumerate(batch_records):
        sku = record["sku"]
        title = record["title"]
        r = index // cols
        c = index % cols

        # Calculate bounding box for this cell
        x0 = int(c * cell_w)
        y0 = int(r * cell_h)
        x1 = int(min(img_w, (c + 1) * cell_w))
        y1 = int(min(img_h, (r + 1) * cell_h))

        # Trim margin slightly (2%) to remove cell divider lines
        trim_x = int(cell_w * trim_border_pct)
        trim_y = int(cell_h * trim_border_pct)
        crop_box = (
            x0 + trim_x,
            y0 + trim_y,
            max(x0 + trim_x + 1, x1 - trim_x),
            max(y0 + trim_y + 1, y1 - trim_y),
        )

        cell_crop = raw_img.crop(crop_box)

        # Handle format & color mode
        sku_filename = f"{safe_filename(sku)}{ext}"
        product_file_path = output_products_dir / sku_filename

        try:
            if image_format in {"jpg", "jpeg"}:
                # Ensure clean white background for transparency in JPEG
                if cell_crop.mode in ("RGBA", "LA") or (cell_crop.mode == "P" and "transparency" in cell_crop.info):
                    bg = Image.new("RGB", cell_crop.size, (255, 255, 255))
                    alpha = cell_crop.convert("RGBA").split()[-1]
                    bg.paste(cell_crop.convert("RGB"), mask=alpha)
                    bg.save(product_file_path, "JPEG", quality=95, optimize=True)
                else:
                    cell_crop.convert("RGB").save(product_file_path, "JPEG", quality=95, optimize=True)
            else:
                cell_crop.save(product_file_path, "PNG")

            logging.info(
                "  [Cell %02d -> Row %d, Col %d] SKU: %s -> %s (%dx%d)",
                index + 1, r + 1, c + 1, sku, product_file_path.name,
                cell_crop.width, cell_crop.height
            )

            results.append({
                "index": index + 1,
                "row": r + 1,
                "col": c + 1,
                "sku": sku,
                "title": title,
                "file_path": product_file_path,
                "status": "success",
                "error": "",
            })
        except Exception as exc:
            logging.error("Failed saving product image for SKU '%s': %s", sku, exc)
            results.append({
                "index": index + 1,
                "row": r + 1,
                "col": c + 1,
                "sku": sku,
                "title": title,
                "file_path": product_file_path,
                "status": "failed",
                "error": str(exc),
            })

    return results


# ---------------------------------------------------------------------------
# MANIFEST MANAGEMENT
# ---------------------------------------------------------------------------

def ensure_manifest(manifest_path: Path) -> None:
    if not manifest_path.exists():
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        with manifest_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=MANIFEST_FIELDNAMES)
            writer.writeheader()


def read_completed_batches(manifest_path: Path) -> set[int]:
    """Read manifest and return batch numbers that completed successfully."""
    if not manifest_path.exists():
        return set()

    completed: set[int] = set()
    failed_batches: set[int] = set()

    with manifest_path.open("r", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                batch_num = int(row["batch_number"])
                if row.get("status") == "success":
                    completed.add(batch_num)
                else:
                    failed_batches.add(batch_num)
            except Exception:
                continue

    # A batch is only considered completed if no failed items exist
    return completed - failed_batches


def append_manifest_records(manifest_path: Path, rows: List[Dict[str, Any]]) -> None:
    with manifest_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_FIELDNAMES)
        for row in rows:
            writer.writerow(row)


# ---------------------------------------------------------------------------
# GOOGLE FLOW BROWSER AUTOMATION (Playwright)
# ---------------------------------------------------------------------------

def find_flow_prompt_box(page) -> Optional[Any]:
    """Locate the prompt input area in Google Flow across various DOM designs."""
    selectors = [
        'textarea[placeholder*="prompt" i]',
        'textarea[placeholder*="describe" i]',
        'textarea[placeholder*="create" i]',
        'textarea',
        'div[contenteditable="true"][role="textbox"]',
        'div[contenteditable="true"]',
        'input[type="text"][placeholder*="prompt" i]',
        '[data-testid*="prompt-input" i]',
        '[aria-label*="prompt" i]',
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=1000):
                return loc
        except Exception:
            continue
    return None


def find_flow_generate_button(page) -> Optional[Any]:
    """Locate the Generate/Create button in Google Flow."""
    selectors = [
        'button:has-text("Generate")',
        'button:has-text("Create")',
        'button:has-text("Submit")',
        'button[aria-label*="Generate" i]',
        'button[aria-label*="Create" i]',
        'button[data-testid*="generate" i]',
        'button[type="submit"]',
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=1000):
                return loc
        except Exception:
            continue
    return None


def check_and_handle_login(page, timeout_seconds: int = 180) -> None:
    """Detect if the user is on Google Sign In, and pause for manual login."""
    url = (page.url or "").lower()
    is_login = any(k in url for k in ["accounts.google.com", "signin", "login"])

    # Also check page content for sign in prompt
    try:
        sign_in_button = page.locator('button:has-text("Sign in"), a:has-text("Sign in")').first
        if sign_in_button.is_visible(timeout=1500):
            is_login = True
    except Exception:
        pass

    if is_login:
        logging.warning("=" * 70)
        logging.warning("GOOGLE LOGIN REQUIRED")
        logging.warning("=" * 70)
        logging.warning("Please sign into your Google account in the opened browser window.")
        logging.warning("Once logged in and Google Flow is loaded, return here and press ENTER.")
        logging.warning("=" * 70)

        # Allow user to press ENTER, or poll until prompt box appears
        start_wait = time.time()
        while time.time() - start_wait < timeout_seconds:
            prompt_box = find_flow_prompt_box(page)
            if prompt_box:
                logging.info("Google Flow prompt box detected! Continuing...")
                return
            time.sleep(2)

        input("Press ENTER when you have completed login and Google Flow is visible: ")


def enter_and_submit_prompt(page, prompt: str) -> None:
    """Enter the prompt into Google Flow and click generate."""
    prompt_box = find_flow_prompt_box(page)
    if not prompt_box:
        check_and_handle_login(page)
        prompt_box = find_flow_prompt_box(page)

    if not prompt_box:
        raise RuntimeError(
            "Could not locate the prompt input box on Google Flow. "
            "Ensure you are logged into Google Flow in the browser."
        )

    logging.info("Entering batch prompt into Google Flow...")
    prompt_box.click()

    try:
        prompt_box.fill(prompt)
    except Exception:
        # Fallback via clipboard
        page.evaluate(
            """async (text) => {
                await navigator.clipboard.writeText(text);
            }""",
            prompt,
        )
        prompt_box.press("Control+V")

    time.sleep(1.0)

    # Click generate button
    gen_btn = find_flow_generate_button(page)
    if gen_btn and gen_btn.is_enabled():
        gen_btn.click()
        logging.info("Clicked 'Generate' button.")
    else:
        prompt_box.press("Enter")
        logging.info("Pressed Enter to submit prompt.")


def wait_for_generated_image(page, timeout_seconds: int = DEFAULT_WAIT_SECONDS) -> Any:
    """Wait for Google Flow to complete generation and return the image element."""
    deadline = time.time() + timeout_seconds
    logging.info("Waiting for Google Flow image generation (timeout %ds)...", timeout_seconds)

    time.sleep(5)  # initial delay for generation to trigger

    while time.time() < deadline:
        # Check for error banners
        try:
            body_text = page.locator("body").inner_text(timeout=1000).lower()
            if any(err in body_text for err in ["something went wrong", "rate limit", "quota exceeded"]):
                logging.warning("Google Flow reported an issue in the interface: %s", body_text[:200])
        except Exception:
            pass

        # Check for image elements in the canvas or output viewer
        try:
            images = page.locator("main img, [role='main'] img, img")
            count = images.count()
            for i in range(count - 1, -1, -1):
                img = images.nth(i)
                if not img.is_visible(timeout=500):
                    continue
                box = img.bounding_box()
                if box and box["width"] >= 350 and box["height"] >= 250:
                    src = img.get_attribute("src") or ""
                    # Ensure it's not a tiny avatar or icon
                    if not any(icon in src.lower() for icon in ["avatar", "profile", "logo", "icon"]):
                        logging.info("Generated image detected (%dx%d px).", int(box["width"]), int(box["height"]))
                        return img
        except Exception:
            pass

        time.sleep(DEFAULT_POLL_INTERVAL)

    raise TimeoutError(f"Google Flow did not produce a generated image within {timeout_seconds} seconds.")


def download_flow_image(page, image_element, output_path: Path) -> None:
    """Download or capture the high-resolution generated image from Google Flow."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Strategy 1: Check for explicit Download button
    try:
        download_btn = page.locator('button[aria-label*="download" i], button:has-text("Download")').first
        if download_btn.is_visible(timeout=1500):
            with page.expect_download(timeout=15000) as download_info:
                download_btn.click()
            download = download_info.value
            download.save_as(str(output_path))
            logging.info("Image downloaded via Download button: %s", output_path)
            return
    except Exception as exc:
        logging.debug("Download button strategy bypassed: %s", exc)

    # Strategy 2: Fetch via src URL or Blob
    try:
        src = image_element.get_attribute("src") or ""
        if src.startswith("http"):
            response = page.context.request.get(src, timeout=30000)
            if response.ok:
                output_path.write_bytes(response.body())
                logging.info("Image downloaded via direct HTTP fetch: %s", output_path)
                return

        if src.startswith(("blob:", "data:")):
            b64_data = page.evaluate(
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
            data = base64.b64decode(b64_data)
            output_path.write_bytes(data)
            logging.info("Image downloaded via Blob decoding: %s", output_path)
            return
    except Exception as exc:
        logging.warning("Fetch/Blob strategy failed: %s", exc)

    # Strategy 3: High-resolution element screenshot fallback
    logging.info("Using element screenshot fallback...")
    image_element.screenshot(path=str(output_path))
    logging.info("Saved image via screenshot fallback: %s", output_path)


# ---------------------------------------------------------------------------
# MAIN WORKFLOW CONTROLLER
# ---------------------------------------------------------------------------

def run_workflow(
    input_file: Path,
    sku_col: str,
    title_col: str,
    output_dir: Path,
    batch_size: int,
    grid_rows: int,
    grid_cols: int,
    image_format: str,
    flow_url: str,
    profile_dir: Path,
    connect_existing: bool,
    cdp_url: str,
    start_batch: int,
    end_batch: Optional[int],
    wait_seconds: int,
    slow_mode: float,
    force: bool,
    split_only: bool,
    dry_run: bool,
    retries: int,
) -> None:
    """Execute the full end-to-end product image generation pipeline."""

    # 1. Load catalog
    products = load_products(input_file, sku_col=sku_col, title_col=title_col)
    total_records = len(products)
    total_batches = (total_records + batch_size - 1) // batch_size
    actual_end_batch = min(end_batch or total_batches, total_batches)

    # 2. Setup directories
    products_output_dir = output_dir / "products"
    batch_grids_dir = output_dir / "batch_grids"
    prompts_dir = output_dir / "prompts"
    debug_dir = output_dir / "debug"

    for d in [products_output_dir, batch_grids_dir, prompts_dir, debug_dir]:
        d.mkdir(parents=True, exist_ok=True)

    manifest_path = output_dir / "manifest.csv"
    ensure_manifest(manifest_path)
    completed_batches = read_completed_batches(manifest_path)

    logging.info("=" * 60)
    logging.info("GOOGLE FLOW SKU IMAGE GENERATOR")
    logging.info("=" * 60)
    logging.info("Input File:        %s", input_file)
    logging.info("Total Products:    %d", total_records)
    logging.info("Batch Size:        %d (Grid: %d rows x %d cols)", batch_size, grid_rows, grid_cols)
    logging.info("Total Batches:     %d", total_batches)
    logging.info("Processing Range:  Batches %d to %d", start_batch, actual_end_batch)
    logging.info("Output Directory:  %s", output_dir)
    logging.info("Output Format:     %s", image_format.upper())
    logging.info("Mode:              %s", "SPLIT-ONLY" if split_only else ("DRY-RUN" if dry_run else "BROWSER AUTOMATION"))
    logging.info("=" * 60)

    # Dry-run inspection mode
    if dry_run:
        logging.info("DRY-RUN: Inspecting batches and prompt structures...")
        for b_num in range(start_batch, actual_end_batch + 1):
            batch, s_idx, e_idx = get_batch(products, b_num, batch_size)
            prompt = build_grid_prompt(batch, rows=grid_rows, cols=grid_cols)
            logging.info("--- Batch %d (%d items: rows %d-%d) ---", b_num, len(batch), s_idx + 1, e_idx)
            for idx, item in enumerate(batch, 1):
                logging.info("  Item %02d | SKU: %-15s | Title: %s", idx, item["sku"], item["title"])
            sample_prompt_file = prompts_dir / f"batch_{b_num:05d}_sample_prompt.txt"
            sample_prompt_file.write_text(prompt, encoding="utf-8")
        logging.info("Dry run complete. Sample prompts saved in: %s", prompts_dir)
        return

    # Split-only mode: process existing raw grids without browser
    if split_only:
        logging.info("SPLIT-ONLY MODE: Slicing previously downloaded batch images...")
        for b_num in range(start_batch, actual_end_batch + 1):
            batch, _, _ = get_batch(products, b_num, batch_size)
            if not batch:
                continue
            raw_path = batch_grids_dir / f"batch_{b_num:05d}_grid.png"
            if not raw_path.exists():
                logging.warning("Batch %d: raw image not found at %s. Skipping.", b_num, raw_path)
                continue

            logging.info("Splitting Batch %d (%s)...", b_num, raw_path.name)
            split_results = split_composite_grid(
                raw_image_path=raw_path,
                batch_records=batch,
                output_products_dir=products_output_dir,
                rows=grid_rows,
                cols=grid_cols,
                image_format=image_format,
            )
            manifest_rows = []
            for item in split_results:
                manifest_rows.append({
                    "batch_number": b_num,
                    "position_in_batch": item["index"],
                    "grid_row": item["row"],
                    "grid_col": item["col"],
                    "sku_id": item["sku"],
                    "product_title": item["title"],
                    "product_image_file": str(item["file_path"]),
                    "raw_grid_file": str(raw_path),
                    "status": item["status"],
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "error": item["error"],
                })
            append_manifest_records(manifest_path, manifest_rows)
        logging.info("Split-only processing finished.")
        return

    # Browser automation mode
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ImportError(
            "Playwright is required for browser automation. "
            "Please install it via:\n  pip install playwright\n  playwright install chromium"
        )

    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        if connect_existing:
            logging.info("Connecting to existing Chrome instance at %s ...", cdp_url)
            browser = p.chromium.connect_over_cdp(cdp_url)
            try:
                browser = p.chromium.connect_over_cdp(cdp_url)
            except Exception as exc:
                raise RuntimeError(
                    f"\n{'='*70}\n"
                    f"COULD NOT CONNECT TO CHROME AT {cdp_url}\n"
                    f"{'='*70}\n"
                    f"To use --connect-existing, Chrome must be running with remote debugging enabled.\n\n"
                    f"HOW TO FIX:\n"
                    f"Option A (Recommended): Simply run the script WITHOUT --connect-existing:\n"
                    f"    python generate_sku_flow_images.py --start-batch 1 --end-batch 2\n\n"
                    f"Option B: First launch Chrome with debugging port 9222 in terminal:\n"
                    f"    google-chrome --remote-debugging-port=9222 &\n"
                    f"Then run:\n"
                    f"    python generate_sku_flow_images.py --connect-existing\n"
                    f"{'='*70}"
                ) from None
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.pages[0] if context.pages else context.new_page()
        else:
            logging.info("Launching Chromium with persistent profile: %s", profile_dir)
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir),
                headless=False,
                viewport={"width": 1440, "height": 960},
                accept_downloads=True,
            )
            page = context.pages[0] if context.pages else context.new_page()

        logging.info("Navigating to Google Flow: %s", flow_url)
        page.goto(flow_url, wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3000)

        check_and_handle_login(page)

        # Batch Processing Loop
        for b_num in range(start_batch, actual_end_batch + 1):
            batch, start_idx, end_idx = get_batch(products, b_num, batch_size)
            if not batch:
                continue

            if not force and b_num in completed_batches:
                logging.info("[%d/%d] Skipping Batch %d (already completed in manifest).", b_num, total_batches, b_num)
                continue

            logging.info(
                "\n" + ("=" * 50) +
                f"\nPROCESSING BATCH {b_num}/{total_batches} (Rows {start_idx + 1}-{end_idx})\n"
                f"SKU Range: {batch[0]['sku']} -> {batch[-1]['sku']}\n" +
                ("=" * 50)
            )

            prompt = build_grid_prompt(batch, rows=grid_rows, cols=grid_cols)
            prompt_file = prompts_dir / f"batch_{b_num:05d}_prompt.txt"
            prompt_file.write_text(prompt, encoding="utf-8")

            raw_grid_path = batch_grids_dir / f"batch_{b_num:05d}_grid.png"

            batch_succeeded = False
            for attempt in range(1, retries + 2):
                try:
                    logging.info("Batch %d: Prompt attempt %d/%d", b_num, attempt, retries + 1)
                    enter_and_submit_prompt(page, prompt)

                    img_elem = wait_for_generated_image(page, timeout_seconds=wait_seconds)
                    download_flow_image(page, img_elem, raw_grid_path)

                    if not raw_grid_path.exists() or raw_grid_path.stat().st_size == 0:
                        raise RuntimeError("Raw composite image file was not successfully created on disk.")

                    # Split the downloaded grid
                    logging.info("Batch %d: Splitting grid into %d product images...", b_num, len(batch))
                    split_results = split_composite_grid(
                        raw_image_path=raw_grid_path,
                        batch_records=batch,
                        output_products_dir=products_output_dir,
                        rows=grid_rows,
                        cols=grid_cols,
                        image_format=image_format,
                    )

                    # Update manifest
                    manifest_rows = []
                    for item in split_results:
                        manifest_rows.append({
                            "batch_number": b_num,
                            "position_in_batch": item["index"],
                            "grid_row": item["row"],
                            "grid_col": item["col"],
                            "sku_id": item["sku"],
                            "product_title": item["title"],
                            "product_image_file": str(item["file_path"]),
                            "raw_grid_file": str(raw_grid_path),
                            "status": item["status"],
                            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "error": item["error"],
                        })
                    append_manifest_records(manifest_path, manifest_rows)
                    batch_succeeded = True
                    logging.info("Batch %d COMPLETED successfully.", b_num)
                    break

                except Exception as exc:
                    logging.error("Batch %d Attempt %d failed: %s", b_num, attempt, exc)
                    # Capture debug screenshot
                    debug_screenshot = debug_dir / f"batch_{b_num:05d}_attempt_{attempt}.png"
                    try:
                        page.screenshot(path=str(debug_screenshot))
                        logging.info("Debug screenshot captured: %s", debug_screenshot)
                    except Exception:
                        pass

                    if attempt <= retries:
                        logging.info("Retrying batch %d in 5 seconds...", b_num)
                        time.sleep(5)

            if not batch_succeeded:
                logging.error("Batch %d failed after %d attempts. Logging failure to manifest.", b_num, retries + 1)
                manifest_rows = []
                for idx, record in enumerate(batch, 1):
                    manifest_rows.append({
                        "batch_number": b_num,
                        "position_in_batch": idx,
                        "grid_row": (idx - 1) // grid_cols + 1,
                        "grid_col": (idx - 1) % grid_cols + 1,
                        "sku_id": record["sku"],
                        "product_title": record["title"],
                        "product_image_file": "",
                        "raw_grid_file": str(raw_grid_path),
                        "status": "failed",
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "error": "Failed generation after retries",
                    })
                append_manifest_records(manifest_path, manifest_rows)

            if slow_mode > 0:
                logging.info("Waiting %.1f seconds before next batch...", slow_mode)
                time.sleep(slow_mode)

        logging.info("\nAll specified batches processed.")
        if not connect_existing:
            context.close()


# ---------------------------------------------------------------------------
# CLI ARGUMENT PARSER
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate and split SKU product images in batches using Google Flow.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    # Input & Output
    parser.add_argument("--input", "-i", default="sku_master_with_description.csv",
                        help="Input CSV or Excel catalog file.")
    parser.add_argument("--sku-col", default=DEFAULT_SKU_COL,
                        help="Column name for SKU ID in input file.")
    parser.add_argument("--title-col", default=DEFAULT_TITLE_COL,
                        help="Column name for product title in input file.")
    parser.add_argument("--output-dir", "-o", default="./output_images",
                        help="Root directory where images and manifest are saved.")
    parser.add_argument("--format", "-f", choices=["png", "jpg", "jpeg"], default="png",
                        help="Output image format for split product files.")

    # Batch & Grid Layout
    parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE,
                        help="Number of products processed per batch / composite image.")
    parser.add_argument("--grid-rows", type=int, default=DEFAULT_GRID_ROWS,
                        help="Number of rows in the composite contact sheet.")
    parser.add_argument("--grid-cols", type=int, default=DEFAULT_GRID_COLS,
                        help="Number of columns in the composite contact sheet.")
    parser.add_argument("--start-batch", type=int, default=1,
                        help="Initial batch index to process (1-based).")
    parser.add_argument("--end-batch", type=int, default=None,
                        help="Final batch index to process (defaults to all batches).")

    # Google Flow & Browser Settings
    parser.add_argument("--flow-url", default=DEFAULT_FLOW_URL,
                        help="Web URL for Google Flow.")
    parser.add_argument("--profile-dir", default="./google_flow_profile",
                        help="Persistent local Chromium user data directory for saved logins.")
    parser.add_argument("--connect-existing", action="store_true",
                        help="Connect to an already open Chrome browser via Chrome DevTools Protocol.")
    parser.add_argument("--cdp-url", default="http://127.0.0.1:9222",
                        help="Chrome DevTools Protocol URL (when using --connect-existing).")
    parser.add_argument("--wait-seconds", type=int, default=DEFAULT_WAIT_SECONDS,
                        help="Timeout in seconds to wait for Google Flow image generation.")
    parser.add_argument("--slow-mode", type=float, default=DEFAULT_SLOW_MODE,
                        help="Delay in seconds between successive batches.")
    parser.add_argument("--retries", type=int, default=1,
                        help="Number of retries per batch if generation or download fails.")

    # Execution Modes
    parser.add_argument("--force", action="store_true",
                        help="Force re-generation of batches already marked successful in manifest.")
    parser.add_argument("--split-only", action="store_true",
                        help="Skip browser generation; re-slice previously downloaded batch images.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Validate input data, print batch mappings and prompts without launching browser.")

    return parser.parse_args()


# ---------------------------------------------------------------------------
# SCRIPT ENTRYPOINT
# ---------------------------------------------------------------------------

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    args = parse_args()

    if args.grid_rows * args.grid_cols != args.batch_size:
        logging.warning(
            "Batch size (%d) does not match grid dimensions (%d rows x %d cols = %d slots). "
            "Grid will be adjusted to accommodate batch size.",
            args.batch_size, args.grid_rows, args.grid_cols, args.grid_rows * args.grid_cols
        )

    try:
        run_workflow(
            input_file=Path(args.input),
            sku_col=args.sku_col,
            title_col=args.title_col,
            output_dir=Path(args.output_dir),
            batch_size=args.batch_size,
            grid_rows=args.grid_rows,
            grid_cols=args.grid_cols,
            image_format=args.format,
            flow_url=args.flow_url,
            profile_dir=Path(args.profile_dir),
            connect_existing=args.connect_existing,
            cdp_url=args.cdp_url,
            start_batch=args.start_batch,
            end_batch=args.end_batch,
            wait_seconds=args.wait_seconds,
            slow_mode=args.slow_mode,
            force=args.force,
            split_only=args.split_only,
            dry_run=args.dry_run,
            retries=args.retries,
        )
    except KeyboardInterrupt:
        logging.warning("\nExecution halted by user (Ctrl+C).")
        sys.exit(130)
    except Exception as exc:
        logging.exception("Fatal error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
