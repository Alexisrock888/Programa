from fastapi import APIRouter, UploadFile, File
from fastapi.responses import JSONResponse
import shutil
import os
import time

from ..services.excel_service import load_excel, save_excel, set_cells_green, set_cells_red, generate_output_filename, get_rows_data
from ..services.file_service import copy_file
from ..config import UPLOADS_DIR, OUTPUTS_DIR

router = APIRouter()


@router.post("/mover-archivo")
async def mover_archivo(file: UploadFile = File(...)):
    upload_path = str(UPLOADS_DIR / file.filename)
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    wb = load_excel(upload_path)
    ws = wb.active

    results = []
    success_count = 0
    error_count = 0

    rows = get_rows_data(ws)

    for row_num, row_data in rows:
        source = row_data[0] if len(row_data) > 0 else None
        destination = row_data[1] if len(row_data) > 1 else None

        if not source or not destination:
            set_cells_red(ws, row_num, [1, 2])
            error_count += 1
            results.append({
                "row": row_num,
                "source": source,
                "destination": destination,
                "success": False,
                "error": "Datos incompletos"
            })
            continue

        ok, error = copy_file(str(source), str(destination))

        if ok:
            set_cells_green(ws, row_num, [1, 2])
            success_count += 1
            results.append({
                "row": row_num,
                "source": source,
                "destination": destination,
                "success": True,
                "error": None
            })
        else:
            set_cells_red(ws, row_num, [1, 2])
            error_count += 1
            results.append({
                "row": row_num,
                "source": source,
                "destination": destination,
                "success": False,
                "error": error
            })

    output_name = generate_output_filename(file.filename, "mover")
    output_path = str(OUTPUTS_DIR / output_name)
    save_excel(wb, output_path)

    os.remove(upload_path)

    return JSONResponse({
        "total_rows": len(results),
        "success_count": success_count,
        "error_count": error_count,
        "rows": results,
        "download_url": f"/api/download/{output_name}"
    })
