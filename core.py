# core.py — Pure logic, no UI, no tkinter
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
import requests
from PIL import Image as PILImage
import io
import concurrent.futures
import os


def download_and_process_image(task):
    """
    Worker function to download and process a single image.
    Runs in a separate thread.
    """
    url = task['url']
    try:
        response = requests.get(url, timeout=20)
        response.raise_for_status()

        img_data = io.BytesIO(response.content)
        pil_img = PILImage.open(img_data)
        pil_img.verify()          # Verify image integrity
        img_data.seek(0)          # Re-open stream after verify

        img = Image(img_data)
        aspect_ratio = pil_img.height / pil_img.width
        img.width = 150
        img.height = img.width * aspect_ratio

        return {
            'status': 'success',
            'task': task,
            'image': img,
            'row_height': img.height * 0.75,
            'col_width': img.width / 7
        }
    except Exception:
        return {
            'status': 'error',
            'task': task
        }


def process_excel(input_bytes, progress_callback=None):
    """
    Main processing function.

    Args:
        input_bytes  : Raw bytes of the uploaded Excel file.
        progress_callback : Optional callable(current, total) for live progress.

    Returns:
        output_bytes : BytesIO containing the processed workbook.
        stats        : dict with keys 'success', 'error', 'total'.
    """
    workbook = openpyxl.load_workbook(io.BytesIO(input_bytes))
    red_font = Font(color="FF0000", bold=True)

    # --- Find all URLs ---
    tasks_to_process = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith(('http://', 'https://')):
                    tasks_to_process.append({
                        'sheet_name': sheet.title,
                        'coordinate': cell.coordinate,
                        'url': cell.value
                    })

    total = len(tasks_to_process)
    stats = {'success': 0, 'error': 0, 'total': total}

    if total == 0:
        # Nothing to process — return the file as-is with a note
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output, stats

    # --- Concurrent download & embed ---
    max_workers = min(32, (os.cpu_count() or 1) * 5)
    completed = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {
            executor.submit(download_and_process_image, task): task
            for task in tasks_to_process
        }

        for future in concurrent.futures.as_completed(future_to_task):
            result = future.result()
            task_info = result['task']
            sheet = workbook[task_info['sheet_name']]
            cell = sheet[task_info['coordinate']]

            if result['status'] == 'success':
                sheet.row_dimensions[cell.row].height = result['row_height']
                sheet.column_dimensions[cell.column_letter].width = result['col_width']
                cell.value = None
                sheet.add_image(result['image'], cell.coordinate)
                stats['success'] += 1
            else:
                cell.value = "Image is corrupt or failed to download"
                cell.font = red_font
                stats['error'] += 1

            completed += 1
            if progress_callback:
                progress_callback(completed, total)

    # --- Save to BytesIO (in memory — no disk write needed) ---
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, stats
