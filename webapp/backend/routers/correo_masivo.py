from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import JSONResponse
import shutil
import os
import time

from ..services.excel_service import load_excel, save_excel, set_cells_green, set_cells_red, generate_output_filename, get_rows_data
from ..services.outlook_service import send_email
from ..utils.html_template import EMAIL_HTML_TEMPLATE, EMAIL_SUBJECT
from ..config import UPLOADS_DIR, OUTPUTS_DIR

router = APIRouter()


@router.post("/correo-masivo")
async def correo_masivo(
    file: UploadFile = File(...),
    pdf_folder: str = Form(...),
    columna_nombre: int = Form(1),
    columna_correo: int = Form(2)
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
        nombre_archivo = row_data[columna_nombre - 1] if len(row_data) >= columna_nombre else None
        correo = row_data[columna_correo - 1] if len(row_data) >= columna_correo else None

        if not nombre_archivo or not correo:
            set_cells_red(ws, row_num, [columna_nombre, columna_correo])
            error_count += 1
            results.append({
                "row": row_num,
                "correo": correo,
                "archivo": nombre_archivo,
                "success": False,
                "error": "Datos incompletos"
            })
            continue

        archivos = str(nombre_archivo).replace(",", ";").split(";")
        adjuntos = []
        archivos_faltantes = []

        for archivo in archivos:
            archivo = archivo.strip()
            if archivo:
                ruta_adjunto = os.path.join(pdf_folder, archivo)
                if os.path.exists(ruta_adjunto):
                    adjuntos.append(ruta_adjunto)
                else:
                    archivos_faltantes.append(archivo)

        if archivos_faltantes:
            set_cells_red(ws, row_num, [columna_nombre, columna_correo])
            error_count += 1
            results.append({
                "row": row_num,
                "correo": correo,
                "archivo": nombre_archivo,
                "success": False,
                "error": f"Archivos no encontrados: {', '.join(archivos_faltantes)}"
            })
            continue

        ok, error = send_email(str(correo), EMAIL_SUBJECT, EMAIL_HTML_TEMPLATE, adjuntos)

        if ok:
            set_cells_green(ws, row_num, [columna_nombre, columna_correo])
            success_count += 1
            results.append({
                "row": row_num,
                "correo": correo,
                "archivo": nombre_archivo,
                "success": True,
                "error": None
            })
        else:
            set_cells_red(ws, row_num, [columna_nombre, columna_correo])
            error_count += 1
            results.append({
                "row": row_num,
                "correo": correo,
                "archivo": nombre_archivo,
                "success": False,
                "error": error
            })

        time.sleep(6)

    output_name = generate_output_filename(file.filename, "correo")
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
