# core.py — Pure logic, no UI, no tkinter
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
import requests
from PIL import Image as PILImage
import io
import concurrent.futures
import os
import gc


# ── Constants ─────────────────────────────────────────────────────────────────
BATCH_SIZE      = 50       # Process N images at a time — keeps RAM flat
MAX_WORKERS     = 8        # Conservative thread count; plenty for I/O-bound work
CELL_WIDTH_PX   = 150      # Cell display width in pixels (does NOT affect image resolution)
REQUEST_TIMEOUT = 20       # Seconds per image download


def download_and_process_image(task):
    """
    Worker: download one image URL and return compressed bytes + dimensions.
    Returns bytes (not an openpyxl Image) so the object is picklable and
    does NOT hold an open file handle across batch boundaries.
    """
    url = task['url']
    try:
        response = requests.get(url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()

        # ── FIX: open twice — once to verify/resize, once clean for openpyxl ──
        raw = response.content

        # Verify + get dimensions using a fresh buffer
        with PILImage.open(io.BytesIO(raw)) as pil_img:
            pil_img.verify()                        # raises if corrupt
            # Re-open (verify() closes the image internally)

        with PILImage.open(io.BytesIO(raw)) as pil_img:
            orig_w, orig_h = pil_img.size
            aspect = orig_h / orig_w if orig_w else 1

            # ── Keep ORIGINAL resolution — no downscaling ──────────────────
            # Only convert to RGBA-safe PNG for lossless embedding
            buf = io.BytesIO()
            pil_img.convert("RGBA").save(buf, format="PNG")
            buf.seek(0)
            img_bytes = buf.getvalue()          # plain bytes — no open handles

        # Cell display size stays at CELL_WIDTH_PX but image data is full-res
        display_h = int(CELL_WIDTH_PX * aspect)

        return {
            'status':      'success',
            'task':        task,
            'img_bytes':   img_bytes,
            'img_width':   CELL_WIDTH_PX,       # how wide the cell shows it
            'img_height':  display_h,           # how tall the cell shows it
            'row_height':  display_h * 0.75,
            'col_width':   CELL_WIDTH_PX / 7,
        }

    except Exception as exc:
        return {
            'status': 'error',
            'task':   task,
            'reason': str(exc),
        }


def _embed_batch(workbook, batch_results, red_font, stats):
    """Embed one batch of download results into the workbook (single-threaded)."""
    for result in batch_results:
        task_info = result['task']
        sheet = workbook[task_info['sheet_name']]
        cell  = sheet[task_info['coordinate']]

        if result['status'] == 'success':
            # Build the openpyxl Image object HERE (main thread, after download)
            img_obj        = Image(io.BytesIO(result['img_bytes']))
            img_obj.width  = result['img_width']
            img_obj.height = result['img_height']

            sheet.row_dimensions[cell.row].height          = result['row_height']
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

    # ── Process in batches ────────────────────────────────────────────────────
    completed = 0

    for batch_start in range(0, total, BATCH_SIZE):
        batch = tasks_to_process[batch_start : batch_start + BATCH_SIZE]

        # Download this batch concurrently
        batch_results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(download_and_process_image, task): task
                       for task in batch}
            for future in concurrent.futures.as_completed(futures):
                batch_results.append(future.result())

        # Embed into workbook (single-threaded — openpyxl is not thread-safe)
        _embed_batch(workbook, batch_results, red_font, stats)

        completed += len(batch)
        if progress_callback:
            progress_callback(completed, total)

        # Free memory before next batch
        del batch_results
        gc.collect()

    # ── Save ──────────────────────────────────────────────────────────────────
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, stats
