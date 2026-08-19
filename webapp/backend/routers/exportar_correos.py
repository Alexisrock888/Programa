from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import JSONResponse
import shutil
import os

from ..services.excel_service import load_excel, save_excel, set_row_green, set_row_red, generate_output_filename, get_rows_data
from ..services.outlook_service import get_sent_items_folder, find_email, save_email_as_pdf
from ..config import UPLOADS_DIR, OUTPUTS_DIR

router = APIRouter()


@router.post("/exportar-correos")
async def exportar_correos(
    file: UploadFile = File(...),
    output_folder: str = Form(...)
):
    upload_path = str(UPLOADS_DIR / file.filename)
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    os.makedirs(output_folder, exist_ok=True)

    wb = load_excel(upload_path)
    ws = wb.active

    results = []
    success_count = 0
    error_count = 0

    try:
        sent_folder, outlook_app = get_sent_items_folder()
    except Exception as e:
        os.remove(upload_path)
        return JSONResponse({
            "error": f"No se pudo conectar a Outlook: {str(e)}"
        }, status_code=503)

    rows = get_rows_data(ws)

    for row_num, row_data in rows:
        destinatario = str(row_data[0]).strip().replace("'", "") if len(row_data) > 0 and row_data[0] else None
        asunto = str(row_data[1]).strip() if len(row_data) > 1 and row_data[1] else None
        hora = row_data[2] if len(row_data) > 2 else None
        fecha = row_data[4] if len(row_data) > 4 else None
        nombre_archivo = str(row_data[6]).strip() if len(row_data) > 6 and row_data[6] else None

        if not all([destinatario, asunto, hora, fecha, nombre_archivo]):
            set_row_red(ws, row_num)
            error_count += 1
            results.append({
                "row": row_num,
                "destinatario": destinatario,
                "asunto": asunto,
                "success": False,
                "error": "Datos incompletos"
            })
            continue

        mail = find_email(sent_folder, destinatario, asunto, fecha, hora)

        if mail:
            ok, pdf_path, error = save_email_as_pdf(mail, output_folder, nombre_archivo)
            if ok:
                set_row_green(ws, row_num)
                success_count += 1
                results.append({
                    "row": row_num,
                    "destinatario": destinatario,
                    "asunto": asunto,
                    "pdf_path": pdf_path,
                    "success": True,
                    "error": None
                })
            else:
                set_row_red(ws, row_num)
                error_count += 1
                results.append({
                    "row": row_num,
                    "destinatario": destinatario,
                    "asunto": asunto,
                    "success": False,
                    "error": error
                })
        else:
            set_row_red(ws, row_num)
            error_count += 1
            results.append({
                "row": row_num,
                "destinatario": destinatario,
                "asunto": asunto,
                "success": False,
                "error": "Correo no encontrado"
            })

    output_name = generate_output_filename(file.filename, "exportar")
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
