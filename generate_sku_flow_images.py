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
1. Standard run (first 2 batches of 15, PNG format):
    python generate_sku_flow_images.py \\
        --input "sku_master_with_description.csv" \\
        --output-dir "./output_images" \\
        --start-batch 1 \\
        --end-batch 2

2. Specific SKU range run (e.g. SKU-AP10401 to SKU-AP15286, batches of 15):
    python generate_sku_flow_images.py \\
        --start-sku "SKU-AP10401" \\
        --end-sku "SKU-AP15286" \\
        --batch-size 15

3. Dry-run a specific SKU range (inspect batches and sample prompts without launching browser):
    python generate_sku_flow_images.py \\
        --start-sku "SKU-AP10401" \\
        --end-sku "SKU-AP15286" \\
        --batch-size 15 \\
        --dry-run \\
        --start-batch 1 \\
        --end-batch 2

4. Output as JPG format:
    python generate_sku_flow_images.py \\
        --input "sku_master_with_description.csv" \\
        --format jpg

5. Custom SKU and Title columns:
    python generate_sku_flow_images.py \\
        --input "products.csv" \\
        --sku-col "SKU" \\
        --title-col "Title"

6. Connect to an already-open Chrome browser (where you are already logged in):
    First start Chrome in terminal:
        google-chrome --remote-debugging-port=9222
    Then run:
        python generate_sku_flow_images.py --connect-existing

7. Re-split existing downloaded grid images without opening the browser:
    python generate_sku_flow_images.py --split-only
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import logging
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from PIL import Image, ImageOps


# ---------------------------------------------------------------------------
# CONSTANTS & DEFAULTS
# ---------------------------------------------------------------------------

# Default SKU Range Configuration (None = process all records)
DEFAULT_START_SKU: Optional[str] = None
DEFAULT_END_SKU: Optional[str] = None

DEFAULT_BATCH_SIZE = 15
DEFAULT_GRID_ROWS = 3
DEFAULT_GRID_COLS = 5
DEFAULT_FLOW_URL = "https://flow.google.com/project/74cfbe79-73e5-4edb-98e7-3e0cccbca7e0"
DEFAULT_WAIT_SECONDS = 420
DEFAULT_SLOW_MODE = 8.0
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
    if not path.exists():
        script_dir_file = Path(__file__).resolve().parent / path.name
        if script_dir_file.exists():
            path = script_dir_file
        else:
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


def filter_products_by_sku_range(
    products: List[Dict[str, str]],
    start_sku: Optional[str] = None,
    end_sku: Optional[str] = None,
) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Filter product records to only those between start_sku and end_sku (inclusive).

    Args:
        products: Full list of catalog records.
        start_sku: Optional starting SKU ID (case-insensitive, whitespace-trimmed).
        end_sku: Optional ending SKU ID (case-insensitive, whitespace-trimmed).

    Returns:
        A tuple of (filtered_products_list, metadata_dict).

    Raises:
        ValueError: If start_sku or end_sku is not found, or if end_sku appears before start_sku.
    """
    clean_start = str(start_sku).strip() if start_sku and str(start_sku).strip() else None
    clean_end = str(end_sku).strip() if end_sku and str(end_sku).strip() else None

    if not clean_start and not clean_end:
        return products, {
            "is_filtered": False,
            "start_sku": products[0]["sku"] if products else "",
            "end_sku": products[-1]["sku"] if products else "",
            "start_index": 0,
            "end_index": len(products) - 1,
            "total_selected": len(products),
        }

    # Find start index
    start_idx = 0
    if clean_start:
        target_start = clean_start.lower()
        for idx, rec in enumerate(products):
            if rec["sku"].strip().lower() == target_start:
                start_idx = idx
                break
        else:
            raise ValueError(
                f"Start SKU '{clean_start}' was not found in catalog ({len(products)} products available). "
                f"Please verify the SKU ID."
            )

    # Find end index
    end_idx = len(products) - 1
    if clean_end:
        target_end = clean_end.lower()
        found = False
        for idx in range(start_idx, len(products)):
            if products[idx]["sku"].strip().lower() == target_end:
                end_idx = idx
                found = True
                break

        if not found:
            # Check if it was located before start_idx to provide an actionable error
            for idx in range(0, start_idx):
                if products[idx]["sku"].strip().lower() == target_end:
                    raise ValueError(
                        f"Start SKU '{clean_start}' (record #{start_idx + 1}, SKU '{products[start_idx]['sku']}') "
                        f"appears AFTER End SKU '{clean_end}' (record #{idx + 1}, SKU '{products[idx]['sku']}') "
                        f"in the catalog. Please invert or correct the specified range."
                    )
            raise ValueError(
                f"End SKU '{clean_end}' was not found in catalog ({len(products)} products available). "
                f"Please verify the SKU ID."
            )

    subset = products[start_idx : end_idx + 1]

    info = {
        "is_filtered": True,
        "start_sku": subset[0]["sku"],
        "end_sku": subset[-1]["sku"],
        "start_index": start_idx,
        "end_index": end_idx,
        "total_selected": len(subset),
    }

    logging.info(
        "Catalog filtered by SKU range: '%s' (row %s) to '%s' (row %s) -> %d products selected (out of %d total).",
        subset[0]["sku"],
        subset[0].get("row_num", str(start_idx + 1)),
        subset[-1]["sku"],
        subset[-1].get("row_num", str(end_idx + 1)),
        len(subset),
        len(products),
    )

    return subset, info


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
    effective_rows = rows

    # Detection: Google Flow / Imagen 3 generates in 16:9 widescreen canvas (e.g. 1376x768).
    # When requesting 5 columns in 16:9, the model naturally generates 3 rows of 5 items (15 cells)
    # to maintain standard square 1:1 cell proportions. If 2 rows were configured for 10 items,
    # the 10 products reside cleanly in Row 1 and Row 2, while Row 3 contains duplicated filler cells.
    if cols == 5 and rows == 2 and (img_w / img_h) > 1.5:
        logging.info(
            "Detected 16:9 widescreen composite grid (%dx%d). Using 3-row layout (cell height: ~%dpx) "
            "to cleanly extract the 10 products from the top 2 rows without bleeding.",
            img_w, img_h, int(img_h / 3)
        )
        effective_rows = 3

    cell_w = img_w / cols
    cell_h = img_h / effective_rows

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


def read_completed_skus(manifest_path: Path, products_dir: Optional[Path] = None) -> set[str]:
    """Read manifest and return set of SKU IDs that completed successfully with verified files on disk."""
    if not manifest_path.exists():
        return set()

    # Track latest status and image path per SKU to support retried batches
    sku_latest: Dict[str, Tuple[str, Optional[str]]] = {}

    with manifest_path.open("r", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sku = (row.get("sku_id") or "").strip()
            if not sku:
                continue
            status = (row.get("status") or "").lower()
            img_path_str = row.get("product_image_file")
            sku_latest[sku] = (status, img_path_str)

    completed: set[str] = set()
    for sku, (status, img_path_str) in sku_latest.items():
        if status == "success":
            image_exists = False
            if img_path_str:
                img_path = Path(img_path_str)
                if img_path.is_file() and img_path.stat().st_size > 0:
                    image_exists = True
                elif products_dir and (products_dir / img_path.name).is_file() and (products_dir / img_path.name).stat().st_size > 0:
                    image_exists = True
            if image_exists or not img_path_str:
                completed.add(sku)

    return completed


def is_batch_completed(batch_records: List[Dict[str, str]], completed_skus: set[str]) -> bool:
    """Return True if all SKUs in the batch have been successfully generated and exist on disk."""
    if not batch_records:
        return False
    return all(rec["sku"] in completed_skus for rec in batch_records)


def get_batch_grid_path(
    batch_grids_dir: Path,
    b_num: int,
    batch_records: List[Dict[str, str]],
    is_custom_range: bool = False,
) -> Path:
    """Determine the raw composite grid image path for a batch."""
    first_sku = safe_filename(batch_records[0]["sku"]) if batch_records else "empty"

    if is_custom_range:
        return batch_grids_dir / f"batch_{b_num:05d}_{first_sku}_grid.png"

    # Default run: prefer standard naming unless custom-tagged exists
    standard = batch_grids_dir / f"batch_{b_num:05d}_grid.png"
    tagged = batch_grids_dir / f"batch_{b_num:05d}_{first_sku}_grid.png"
    if tagged.exists() and not standard.exists():
        return tagged
    return standard


def get_batch_prompt_path(
    prompts_dir: Path,
    b_num: int,
    batch_records: List[Dict[str, str]],
    is_custom_range: bool = False,
    is_sample: bool = False,
) -> Path:
    """Determine the prompt text file path for a batch."""
    prefix = f"batch_{b_num:05d}"
    first_sku = safe_filename(batch_records[0]["sku"]) if batch_records else "empty"
    tag = f"_{first_sku}" if is_custom_range else ""
    suffix = "_sample_prompt.txt" if is_sample else "_prompt.txt"
    return prompts_dir / f"{prefix}{tag}{suffix}"


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
        'div.ProseMirror[contenteditable="true"]',
        'div[contenteditable="true"]',
        'textarea[placeholder*="prompt" i]',
        'textarea[placeholder*="describe" i]',
        'textarea[placeholder*="create" i]',
        'textarea',
        'div[contenteditable="true"][role="textbox"]',
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
        'button[aria-label="Start generation"]',
        'button:has-text("arrow_forward")',
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
    """Detect if the user is on Google Sign In or landing page, and pause for manual login."""
    # Check if we are on landing splash page and click 'Create with Google Flow'
    try:
        landing_btn = page.locator('button:has-text("Create with Google Flow"), a:has-text("Create with Google Flow"), button:has-text("Get started"), a:has-text("Get started")').first
        if landing_btn.is_visible(timeout=2000):
            logging.info("Found Google Flow landing page. Clicking 'Create with Google Flow'...")
            landing_btn.click()
            page.wait_for_timeout(3000)
    except Exception:
        pass

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
        logging.warning("Please sign into your Google account in the opened Google Chrome window.")
        logging.warning("Once logged in and Google Flow workspace is loaded, return here and press ENTER.")
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


def get_flow_status(page) -> Dict[str, Any]:
    """Inspect current Google Flow state: generating vs idle, progress %, error tiles count, etc."""
    try:
        return page.evaluate(r'''() => {
            const flowGen = document.querySelector('flow-generate-icon-button');
            const btn = flowGen ? flowGen.querySelector('button') : document.querySelector('button.generate-icon-button');
            const iconText = (btn?.querySelector('mat-icon')?.innerText || '').trim().toLowerCase();
            const ariaLabel = (btn?.getAttribute('aria-label') || '').toLowerCase();

            const isIdle = (ariaLabel === 'start generation' || iconText === 'arrow_forward');
            const hasStop = ariaLabel.includes('stop') || iconText === 'stop' || (!isIdle && btn !== null);

            // Check progress percentage text on tiles (e.g. 50%, 76%)
            const tiles = Array.from(document.querySelectorAll('flow-grid-tile-container, flow-image-tile, flow-node'));
            let maxPercent = null;
            for (const t of tiles) {
                const m = (t.innerText || '').match(/\b(\d{1,2})%\b/);
                if (m) {
                    const val = parseInt(m[1], 10);
                    if (maxPercent === null || val > maxPercent) {
                        maxPercent = val;
                    }
                }
            }

            // Count error tiles
            let errorCount = 0;
            const errorTexts = [];
            for (const t of tiles) {
                if (t.tagName === 'FLOW-IMAGE-TILE') {
                    const text = (t.innerText || '');
                    if (text.includes('Sorry, this image failed to generate') || text.includes('Failed to generate')) {
                        errorCount++;
                        if (errorTexts.length < 3) {
                            errorTexts.push(text.slice(0, 100).replace(/\n/g, ' '));
                        }
                    }
                }
            }

            // Rate limit / quota
            const bodyText = (document.body.innerText || '').toLowerCase();
            const isQuotaExceeded = bodyText.includes('rate limit') || bodyText.includes('quota exceeded') || bodyText.includes('resource exhausted');

            return {
                isIdle: !!isIdle,
                hasStop: !!hasStop,
                maxPercent,
                errorCount,
                errorTexts,
                isGenerating: !isIdle || (maxPercent !== null),
                isQuotaExceeded
            };
        }''')
    except Exception:
        return {
            "isIdle": True,
            "hasStop": False,
            "maxPercent": None,
            "errorCount": 0,
            "errorTexts": [],
            "isGenerating": False,
            "isQuotaExceeded": False,
        }


def get_flow_error_count(page) -> int:
    """Count failed generation tiles currently on the Google Flow canvas."""
    status = get_flow_status(page)
    return status.get("errorCount", 0)


def wait_for_flow_idle(page, timeout_seconds: int = 240) -> bool:
    """Wait until Google Flow completes any in-progress generation and returns to idle."""
    start = time.time()
    last_log = 0.0
    while time.time() - start < timeout_seconds:
        status = get_flow_status(page)
        if not status.get("isGenerating") and not status.get("hasStop"):
            prompt_box = find_flow_prompt_box(page)
            if prompt_box:
                return True
        now = time.time()
        if now - last_log >= 10:
            pct_info = f" ({status['maxPercent']}%)" if status.get("maxPercent") is not None else ""
            logging.info("Google Flow is busy generating%s. Waiting for it to become idle...", pct_info)
            last_log = now
        time.sleep(3.0)
    logging.warning("Timed out waiting for Google Flow to become idle.")
    return False


def enter_and_submit_prompt(page, prompt: str) -> None:
    """Enter the prompt into Google Flow and click generate."""
    # Ensure Flow is not busy generating from a previous prompt before typing
    wait_for_flow_idle(page)

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
    time.sleep(0.3)

    try:
        select_all = "Meta+A" if sys.platform == "darwin" else "Control+A"
        page.keyboard.press(select_all)
        page.keyboard.press("Backspace")
        time.sleep(0.2)
        page.keyboard.insert_text(prompt)
    except Exception:
        try:
            prompt_box.fill(prompt)
        except Exception:
            page.evaluate(
                """async (text) => {
                    await navigator.clipboard.writeText(text);
                }""",
                prompt,
            )
            paste_key = "Meta+V" if sys.platform == "darwin" else "Control+V"
            prompt_box.press(paste_key)

    time.sleep(1.0)

    # Click generate button
    gen_btn = find_flow_generate_button(page)
    if gen_btn and gen_btn.is_enabled():
        gen_btn.click()
        logging.info("Clicked 'Generate' button.")
    else:
        prompt_box.press("Enter")
        logging.info("Pressed Enter to submit prompt.")

    # Brief delay for Flow to transition to generating
    time.sleep(3.0)


def get_existing_flow_images(page) -> Set[str]:
    """Retrieve all high-resolution generated image URLs currently in Google Flow."""
    try:
        urls = page.evaluate("""() => {
            const imgs = Array.from(document.querySelectorAll("flow-image-tile img, img.image, [class*='image'] img, img"));
            return imgs
                .filter(img => (img.naturalWidth > 300 || img.clientWidth > 300) && img.src)
                .map(img => img.src);
        }""")
        return set(urls)
    except Exception:
        return set()


def wait_for_generated_image(
    page,
    existing_srcs: Optional[Set[str]] = None,
    initial_error_count: int = 0,
    timeout_seconds: int = DEFAULT_WAIT_SECONDS,
) -> Dict[str, Any]:
    """Wait for Google Flow to complete generation of a NEW image and return image metadata.

    Handles 3-4 minute generation times, active rendering progress (e.g. 50%, 76%),
    and avoids false-positive triggers from pre-existing canvas errors.
    """
    start_time = time.time()
    deadline = start_time + timeout_seconds
    known_srcs = set(existing_srcs or [])
    logging.info(
        "Waiting for NEW Google Flow image generation (timeout %ds, %d pre-existing canvas images, %d pre-existing canvas errors)...",
        timeout_seconds,
        len(known_srcs),
        initial_error_count,
    )

    last_logged_pct = None
    last_log_time = 0.0

    # Initial delay for Google Flow backend to acknowledge generation
    time.sleep(4)

    while time.time() < deadline:
        elapsed = int(time.time() - start_time)

        # 1. PRIORITY 1: Check for NEW high-resolution completed image elements
        try:
            new_img = page.evaluate("""(knownList) => {
                const known = new Set(knownList);

                // Priority 1: Top tile in virtual scroll feed
                const tiles = Array.from(document.querySelectorAll("flow-grid-tile-container flow-image-tile, flow-image-tile"));
                for (const tile of tiles) {
                    const img = tile.querySelector("img");
                    if (img && img.src && !known.has(img.src)) {
                        if (img.complete && (img.naturalWidth > 350 || img.clientWidth > 350)) {
                            return {
                                src: img.src,
                                width: img.naturalWidth || img.clientWidth,
                                height: img.naturalHeight || img.clientHeight,
                            };
                        }
                    }
                }

                // Priority 2: General canvas images
                const imgs = Array.from(document.querySelectorAll("main img, [role='main'] img, img"));
                for (const img of imgs) {
                    if (img.src && !known.has(img.src)) {
                        const s = img.src.toLowerCase();
                        if (s.includes("avatar") || s.includes("profile") || s.includes("logo") || s.includes("icon")) {
                            continue;
                        }
                        if (img.complete && (img.naturalWidth > 350 || img.clientWidth > 350)) {
                            return {
                                src: img.src,
                                width: img.naturalWidth || img.clientWidth,
                                height: img.naturalHeight || img.clientHeight,
                            };
                        }
                    }
                }
                return null;
            }""", list(known_srcs))

            if new_img and new_img.get("src"):
                logging.info(
                    "New generated image detected (%dx%d px) after %ds: %s",
                    new_img.get("width", 0),
                    new_img.get("height", 0),
                    elapsed,
                    new_img["src"][:80],
                )
                return new_img
        except Exception as exc:
            logging.debug("Error checking for new image: %s", exc)

        # 2. PRIORITY 2: Check current Flow UI status (generating, percentage, stop button)
        status = get_flow_status(page)

        if status.get("isQuotaExceeded"):
            raise RuntimeError("Rate limit / quota exceeded reported by Google Flow.")

        is_generating = status.get("isGenerating", False) or status.get("hasStop", False)
        current_pct = status.get("maxPercent")

        # If actively generating, log progress periodically and keep waiting!
        if is_generating or current_pct is not None:
            now = time.time()
            if current_pct != last_logged_pct or (now - last_log_time >= 15):
                pct_str = f" ({current_pct}%)" if current_pct is not None else ""
                logging.info(
                    "Google Flow is generating%s... elapsed: %ds / %ds",
                    pct_str,
                    elapsed,
                    timeout_seconds,
                )
                last_logged_pct = current_pct
                last_log_time = now

            time.sleep(DEFAULT_POLL_INTERVAL)
            continue

        # 3. PRIORITY 3: Generation has stopped (idle state) and NO new image was detected.
        # Check if a genuine NEW error occurred for THIS generation.
        current_error_count = status.get("errorCount", 0)
        if current_error_count > initial_error_count:
            error_msg = status.get("errorTexts", ["Image generation failed in Google Flow."])[0] if status.get("errorTexts") else "Image generation failed in Google Flow."
            raise RuntimeError(f"Image generation failed in Google Flow: {error_msg}")

        # If Flow is idle and we've been waiting at least 15 seconds after submission,
        # perform a final check for new images before concluding failure
        if elapsed > 15:
            time.sleep(3)
            try:
                final_check = page.evaluate("""(knownList) => {
                    const known = new Set(knownList);
                    const imgs = Array.from(document.querySelectorAll("flow-image-tile img, img"));
                    for (const img of imgs) {
                        if (img.src && !known.has(img.src) && img.complete && (img.naturalWidth > 350 || img.clientWidth > 350)) {
                            return {
                                src: img.src,
                                width: img.naturalWidth || img.clientWidth,
                                height: img.naturalHeight || img.clientHeight,
                            };
                        }
                    }
                    return null;
                }""", list(known_srcs))
                if final_check and final_check.get("src"):
                    logging.info(
                        "New generated image detected (%dx%d px) after final check: %s",
                        final_check.get("width", 0),
                        final_check.get("height", 0),
                        final_check["src"][:80],
                    )
                    return final_check
            except Exception:
                pass

            if elapsed > 30:
                raise RuntimeError("Google Flow generation returned to idle without producing a new image.")

        time.sleep(DEFAULT_POLL_INTERVAL)

    raise TimeoutError(f"Google Flow did not produce a new generated image within {timeout_seconds} seconds.")


def download_flow_image(page, image_target: Union[Dict[str, Any], Any], output_path: Path) -> None:
    """Download or capture the high-resolution generated image from Google Flow."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    src = None
    if isinstance(image_target, dict):
        src = image_target.get("src")
    elif hasattr(image_target, "get_attribute"):
        src = image_target.get_attribute("src")

    # Strategy 1: Fetch via Playwright context request (direct HTTP download)
    if src and src.startswith("http"):
        try:
            response = page.context.request.get(src, timeout=30000)
            if response.ok:
                output_path.write_bytes(response.body())
                logging.info("Image downloaded via direct HTTP fetch (%d bytes): %s", len(response.body()), output_path)
                return
        except Exception as exc:
            logging.warning("Direct HTTP fetch failed: %s", exc)

    # Strategy 2: Check for explicit Download button in the UI
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

    # Strategy 3: Fetch via Blob / Data URL decoding
    if src and src.startswith(("blob:", "data:")):
        try:
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
            logging.warning("Blob decode failed: %s", exc)

    # Strategy 4: High-resolution element screenshot fallback
    try:
        if hasattr(image_target, "screenshot"):
            image_target.screenshot(path=str(output_path))
            logging.info("Saved image via element screenshot fallback: %s", output_path)
            return
        elif src:
            elem = page.locator(f"img[src='{src}']").first
            if elem.is_visible(timeout=2000):
                elem.screenshot(path=str(output_path))
                logging.info("Saved image via locator screenshot fallback: %s", output_path)
                return
    except Exception as exc:
        logging.warning("Element screenshot fallback failed: %s", exc)

    raise RuntimeError(f"Could not download or capture image to {output_path}")


def ensure_cdp_chrome_running(cdp_url: str = "http://127.0.0.1:9222", profile_dir: Optional[Path] = None) -> bool:
    """Check if Chrome is running on CDP port; if not, attempt to launch native Google Chrome on macOS/Linux/Windows."""
    try:
        req = urllib.request.Request(f"{cdp_url}/json/version")
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            if resp.status == 200:
                logging.info("Existing Chrome CDP instance detected at %s.", cdp_url)
                return True
    except Exception:
        pass

    logging.info("No active Chrome detected at %s. Attempting to launch native Google Chrome with remote debugging...", cdp_url)
    chrome_profile = (profile_dir or (Path.cwd() / "chrome_flow_profile")).resolve()
    chrome_profile.mkdir(parents=True, exist_ok=True)

    port = 9222
    m = re.search(r":(\d+)", cdp_url)
    if m:
        port = int(m.group(1))

    launched = False
    if sys.platform == "darwin":
        chrome_app = Path("/Applications/Google Chrome.app")
        if chrome_app.exists():
            subprocess.Popen([
                "open", "-na", "Google Chrome",
                "--args",
                f"--remote-debugging-port={port}",
                f"--user-data-dir={chrome_profile}",
                "--no-first-run",
                "--no-default-browser-check",
            ])
            launched = True
        else:
            logging.warning("Google Chrome not found at /Applications/Google Chrome.app")
    elif sys.platform.startswith("linux"):
        for binary in ["google-chrome", "google-chrome-stable", "chromium-browser", "chromium"]:
            if shutil.which(binary):
                subprocess.Popen([
                    binary,
                    f"--remote-debugging-port={port}",
                    f"--user-data-dir={chrome_profile}",
                    "--no-first-run",
                    "--no-default-browser-check",
                ])
                launched = True
                break
    elif sys.platform == "win32":
        for path_str in [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]:
            if Path(path_str).exists():
                subprocess.Popen([
                    path_str,
                    f"--remote-debugging-port={port}",
                    f"--user-data-dir={chrome_profile}",
                    "--no-first-run",
                    "--no-default-browser-check",
                ])
                launched = True
                break

    if launched:
        logging.info("Waiting for Google Chrome remote debugging port %d to initialize...", port)
        for _ in range(15):
            time.sleep(1.0)
            try:
                req = urllib.request.Request(f"{cdp_url}/json/version")
                with urllib.request.urlopen(req, timeout=1.0) as resp:
                    if resp.status == 200:
                        logging.info("Native Google Chrome successfully initialized on port %d!", port)
                        return True
            except Exception:
                continue
        logging.warning("Timed out waiting for Chrome remote debugging port %d to respond.", port)

    return False


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
    start_sku: Optional[str] = None,
    end_sku: Optional[str] = None,
) -> None:
    """Execute the full end-to-end product image generation pipeline."""

    # 1. Load catalog & apply SKU range filter
    all_products = load_products(input_file, sku_col=sku_col, title_col=title_col)
    products, sku_range_info = filter_products_by_sku_range(
        all_products,
        start_sku=start_sku,
        end_sku=end_sku,
    )
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
    completed_skus = read_completed_skus(manifest_path, products_output_dir)

    logging.info("=" * 60)
    logging.info("GOOGLE FLOW SKU IMAGE GENERATOR")
    logging.info("=" * 60)
    logging.info("Input File:        %s", input_file)
    logging.info("Total in Catalog:  %d", len(all_products))
    if sku_range_info["is_filtered"]:
        logging.info("SKU Range Filter:  '%s' -> '%s'", sku_range_info["start_sku"], sku_range_info["end_sku"])
        logging.info("Selected Products: %d (records #%d to #%d)", total_records, sku_range_info["start_index"] + 1, sku_range_info["end_index"] + 1)
    else:
        logging.info("Selected Products: %d (entire catalog)", total_records)
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
            sample_prompt_file = get_batch_prompt_path(
                prompts_dir, b_num, batch, is_custom_range=sku_range_info["is_filtered"], is_sample=True
            )
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
            raw_path = get_batch_grid_path(batch_grids_dir, b_num, batch, is_custom_range=sku_range_info["is_filtered"])
            if not raw_path.exists():
                fallback_tagged = batch_grids_dir / f"batch_{b_num:05d}_{safe_filename(batch[0]['sku'])}_grid.png"
                fallback_std = batch_grids_dir / f"batch_{b_num:05d}_grid.png"
                if fallback_tagged.exists():
                    raw_path = fallback_tagged
                elif fallback_std.exists():
                    raw_path = fallback_std
                else:
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

    if connect_existing:
        logging.info("Checking connection to Chrome instance at %s ...", cdp_url)
        ensure_cdp_chrome_running(cdp_url, profile_dir=Path("./chrome_flow_profile"))

    with sync_playwright() as p:
        is_cdp_connected = False
        if connect_existing:
            try:
                browser = p.chromium.connect_over_cdp(cdp_url)
                context = browser.contexts[0] if browser.contexts else browser.new_context()
                page = None
                if context.pages:
                    for p_candidate in context.pages:
                        u = (p_candidate.url or "").lower()
                        if "flow.google.com/project" in u or "498b28ec-6d53-467c-a557-7cd400c259a5" in u:
                            page = p_candidate
                            logging.info("Reusing existing Google Flow tab: %s", page.url)
                            break
                        elif "flow.google.com" in u:
                            page = p_candidate
                if not page:
                    page = context.pages[0] if context.pages else context.new_page()

                is_cdp_connected = True
                logging.info("Successfully connected to Google Chrome via CDP.")
            except Exception as exc:
                logging.warning(
                    "Could not connect to Chrome at %s (%s).\n"
                    "Note: If needed, start Chrome manually with:\n"
                    "  open -na \"Google Chrome\" --args --remote-debugging-port=9222 --user-data-dir=\"$PWD/chrome_flow_profile\"\n"
                    "Falling back to Playwright persistent profile...",
                    cdp_url, exc,
                )

        if not is_cdp_connected:
            logging.info("Launching browser with persistent profile: %s", profile_dir)
            launch_args: Dict[str, Any] = {
                "user_data_dir": str(profile_dir),
                "headless": False,
                "viewport": {"width": 1440, "height": 960},
                "accept_downloads": True,
                "ignore_default_args": ["--enable-automation"],
                "args": ["--disable-blink-features=AutomationControlled"],
            }
            if sys.platform == "darwin" and Path("/Applications/Google Chrome.app").exists():
                launch_args["channel"] = "chrome"
            elif shutil.which("google-chrome"):
                launch_args["channel"] = "chrome"

            context = p.chromium.launch_persistent_context(**launch_args)
            page = context.pages[0] if context.pages else context.new_page()

        current_url = (page.url or "").rstrip("/")
        target_url = flow_url.rstrip("/")
        if target_url not in current_url:
            logging.info("Navigating to Google Flow project: %s", flow_url)
            page.goto(flow_url, wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(3000)

        try:
            page.bring_to_front()
        except Exception:
            pass

        check_and_handle_login(page)

        # Batch Processing Loop
        for b_num in range(start_batch, actual_end_batch + 1):
            batch, start_idx, end_idx = get_batch(products, b_num, batch_size)
            if not batch:
                continue

            if not force and is_batch_completed(batch, completed_skus):
                logging.info(
                    "[%d/%d] Skipping Batch %d (%s -> %s): all %d products already completed in manifest.",
                    b_num, total_batches, b_num, batch[0]["sku"], batch[-1]["sku"], len(batch)
                )
                continue

            logging.info(
                "\n" + ("=" * 50) +
                f"\nPROCESSING BATCH {b_num}/{total_batches} (Items {start_idx + 1}-{end_idx} of selection)\n"
                f"SKU Range: {batch[0]['sku']} -> {batch[-1]['sku']}\n" +
                ("=" * 50)
            )

            prompt = build_grid_prompt(batch, rows=grid_rows, cols=grid_cols)
            prompt_file = get_batch_prompt_path(
                prompts_dir, b_num, batch, is_custom_range=sku_range_info["is_filtered"]
            )
            prompt_file.write_text(prompt, encoding="utf-8")

            raw_grid_path = get_batch_grid_path(
                batch_grids_dir, b_num, batch, is_custom_range=sku_range_info["is_filtered"]
            )

            batch_succeeded = False
            for attempt in range(1, retries + 2):
                try:
                    logging.info("Batch %d: Prompt attempt %d/%d", b_num, attempt, retries + 1)
                    # Ensure Flow is idle before snapshotting and submitting
                    wait_for_flow_idle(page)

                    # Snapshot existing images and error count before submitting prompt so we strictly wait for the new one
                    existing_srcs = get_existing_flow_images(page)
                    initial_errors = get_flow_error_count(page)
                    logging.info(
                        "Batch %d: Detected %d pre-existing image(s) and %d error card(s) on Flow canvas.",
                        b_num, len(existing_srcs), initial_errors
                    )

                    enter_and_submit_prompt(page, prompt)

                    img_target = wait_for_generated_image(
                        page,
                        existing_srcs=existing_srcs,
                        initial_error_count=initial_errors,
                        timeout_seconds=wait_seconds,
                    )
                    download_flow_image(page, img_target, raw_grid_path)

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
                        if item["status"] == "success":
                            completed_skus.add(item["sku"])

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
                        logging.info("Retrying batch %d in 10 seconds...", b_num)
                        time.sleep(10)

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
        if not is_cdp_connected:
            context.close()


# ---------------------------------------------------------------------------
# GRID & DIMENSION UTILITIES
# ---------------------------------------------------------------------------

def calculate_grid_dimensions(
    batch_size: int,
    user_rows: Optional[int] = None,
    user_cols: Optional[int] = None,
) -> Tuple[int, int]:
    """Calculate or adjust grid rows and cols to cleanly accommodate batch_size."""
    if user_rows is not None and user_cols is not None:
        if user_rows * user_cols == batch_size:
            return user_rows, user_cols

    presets: Dict[int, Tuple[int, int]] = {
        15: (3, 5),
        10: (2, 5),
        20: (4, 5),
        25: (5, 5),
        30: (6, 5),
        12: (3, 4),
        16: (4, 4),
        9: (3, 3),
        8: (2, 4),
        6: (2, 3),
        4: (2, 2),
    }

    if batch_size in presets:
        return presets[batch_size]

    cols = user_cols or DEFAULT_GRID_COLS
    rows = (batch_size + cols - 1) // cols
    return rows, cols


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

    # SKU Range Selection (Optional subsetting)
    parser.add_argument("--start-sku", "-s", default=DEFAULT_START_SKU,
                        help="Start SKU ID to filter catalog range (inclusive, e.g. 'SKU-AP10401').")
    parser.add_argument("--end-sku", "-e", default=DEFAULT_END_SKU,
                        help="End SKU ID to filter catalog range (inclusive, e.g. 'SKU-AP15286').")

    # Batch & Grid Layout
    parser.add_argument("--batch-size", "-b", type=int, default=DEFAULT_BATCH_SIZE,
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

    grid_rows, grid_cols = calculate_grid_dimensions(
        batch_size=args.batch_size,
        user_rows=args.grid_rows if args.grid_rows != DEFAULT_GRID_ROWS else None,
        user_cols=args.grid_cols if args.grid_cols != DEFAULT_GRID_COLS else None,
    )

    if grid_rows * grid_cols != args.batch_size:
        logging.info(
            "Batch size (%d) fits into grid layout (%d rows x %d cols = %d slots).",
            args.batch_size, grid_rows, grid_cols, grid_rows * grid_cols
        )

    try:
        run_workflow(
            input_file=Path(args.input),
            sku_col=args.sku_col,
            title_col=args.title_col,
            output_dir=Path(args.output_dir),
            batch_size=args.batch_size,
            grid_rows=grid_rows,
            grid_cols=grid_cols,
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
            start_sku=args.start_sku,
            end_sku=args.end_sku,
        )
    except KeyboardInterrupt:
        logging.warning("\nExecution halted by user (Ctrl+C).")
        sys.exit(130)
    except Exception as exc:
        logging.exception("Fatal error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
