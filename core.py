# core.py — Pure logic, no UI, no tkinter
# Mirrors the original CMD script logic for maximum speed

import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
import requests
from PIL import Image as PILImage
import io
import concurrent.futures
import os
import time
import gc

# ── Constants ──────────────────────────────────────────────────────────────────
BATCH_SIZE      = 100                            # larger batches = fewer overhead cycles
MAX_WORKERS     = min(32, (os.cpu_count() or 1) * 5)  # same as original script
REQUEST_TIMEOUT = 20
MAX_RETRIES     = 3
BACKOFF         = 0.75

# Same headers as your original script — critical for servers that block plain requests
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; ExcelImageFetcher/1.0; +https://pattern.com)",
    "Accept": "image/*,application/octet-stream;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}


def fetch_bytes_with_retries(url):
    """Download bytes with retries + proper headers. Same as original script."""
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            resp.raise_for_status()
            return resp.content
        except Exception as e:
            last_exc = e
            if attempt < MAX_RETRIES:
                time.sleep(BACKOFF * attempt)
    raise last_exc


def download_and_process_image(task):
    """
    Worker: download one image, return raw bytes + metadata.
    Returns bytes instead of openpyxl Image so it's thread-safe.
    Same logic as original script but decoupled from openpyxl.
    """
    url = task['url']
    try:
        raw = fetch_bytes_with_retries(url)

        # Fully decode image (same as original — avoids verify() trap)
        pil_img = PILImage.open(io.BytesIO(raw))
        pil_img.load()

        # Handle all image modes (WEBP, HEIC, CMYK, etc.) — same as original
        if pil_img.mode in ("P", "LA"):
            pil_img = pil_img.convert("RGBA")
        elif pil_img.mode not in ("RGB", "RGBA"):
            pil_img = pil_img.convert("RGB")

        # Re-encode as PNG — robust for all Excel/openpyxl versions
        out_buf = io.BytesIO()
        pil_img.save(out_buf, format="PNG", optimize=True)
        out_buf.seek(0)
        img_bytes = out_buf.getvalue()

        # Sizing: 150px wide, preserve aspect ratio — same as original
        aspect_ratio = pil_img.height / pil_img.width if pil_img.width else 1.0
        img_w = 150
        img_h = img_w * aspect_ratio

        return {
            'status':     'success',
            'task':       task,
            'img_bytes':  img_bytes,
            'img_width':  img_w,
            'img_height': img_h,
            'row_height': img_h * 0.75,
            'col_width':  img_w / 7,
        }

    except Exception:
        return {
            'status': 'error',
            'task':   task,
        }


def _embed_batch(workbook, batch_results, red_font, stats):
    """Embed one batch into workbook. Single-threaded — openpyxl is not thread-safe."""
    for result in batch_results:
        task_info = result['task']
        sheet = workbook[task_info['sheet_name']]
        cell  = sheet[task_info['coordinate']]

        if result['status'] == 'success':
            img_obj        = Image(io.BytesIO(result['img_bytes']))
            img_obj.width  = result['img_width']
            img_obj.height = result['img_height']

            sheet.row_dimensions[cell.row].height             = result['row_height']
            sheet.column_dimensions[cell.column_letter].width = result['col_width']
            cell.value = None
            sheet.add_image(img_obj, cell.coordinate)
            stats['success'] += 1
        else:
            cell.value = "Image is corrupt or failed to download"
            cell.font  = red_font
            stats['error'] += 1


def process_excel(input_bytes, progress_callback=None):
    """
    Main processing function.

    Args:
        input_bytes       : Raw bytes of the uploaded Excel file.
        progress_callback : Optional callable(current, total) for live progress.

    Returns:
        output_bytes : BytesIO containing the processed workbook.
        stats        : dict with keys 'success', 'error', 'total'.
    """
    workbook = openpyxl.load_workbook(io.BytesIO(input_bytes))
    red_font = Font(color="FF0000", bold=True)

    # ── Collect all URL tasks ──────────────────────────────────────────────────
    tasks_to_process = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith(('http://', 'https://')):
                    tasks_to_process.append({
                        'sheet_name': sheet.title,
                        'coordinate': cell.coordinate,
                        'url':        cell.value,
                    })

    total = len(tasks_to_process)
    stats = {'success': 0, 'error': 0, 'total': total}

    if total == 0:
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output, stats

    # ── Process in batches — high concurrency like original script ─────────────
    completed = 0

    for batch_start in range(0, total, BATCH_SIZE):
        batch = tasks_to_process[batch_start : batch_start + BATCH_SIZE]

        batch_results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(download_and_process_image, task): task
                       for task in batch}
            for future in concurrent.futures.as_completed(futures):
                batch_results.append(future.result())

        _embed_batch(workbook, batch_results, red_font, stats)

        completed += len(batch)
        if progress_callback:
            progress_callback(completed, total)

        del batch_results
        gc.collect()

    # ── Save ──────────────────────────────────────────────────────────────────
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, stats
