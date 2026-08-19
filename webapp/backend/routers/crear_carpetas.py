from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import JSONResponse
import shutil
import os

from ..services.excel_service import load_excel, save_excel, set_cells_green, set_cells_red, generate_output_filename, get_rows_data
from ..services.file_service import create_folder
from ..config import UPLOADS_DIR, OUTPUTS_DIR

router = APIRouter()


@router.post("/crear-carpetas")
async def crear_carpetas(
    file: UploadFile = File(...),
    base_path: str = Form(...),
    columna: int = Form(1)
):
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
        folder_name = row_data[columna - 1] if len(row_data) >= columna else None

        if not folder_name:
            set_cells_red(ws, row_num, [columna])
            error_count += 1
            results.append({
                "row": row_num,
                "folder_name": None,
                "full_path": None,
                "success": False,
                "error": "Nombre vacío"
            })
            continue

        ok, full_path, error = create_folder(base_path, str(folder_name))

        if ok:
            set_cells_green(ws, row_num, [columna])
            success_count += 1
            results.append({
                "row": row_num,
                "folder_name": str(folder_name),
                "full_path": full_path,
                "success": True,
                "error": None
            })
        else:
            set_cells_red(ws, row_num, [columna])
            error_count += 1
            results.append({
                "row": row_num,
                "folder_name": str(folder_name),
                "full_path": None,
                "success": False,
                "error": error
            })

    output_name = generate_output_filename(file.filename, "carpetas")
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
