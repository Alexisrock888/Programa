from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import JSONResponse
import shutil
import os

from ..services.excel_service import load_excel, save_excel, generate_output_filename
from ..services.word_service import split_word_to_pdfs
from ..config import UPLOADS_DIR, OUTPUTS_DIR

router = APIRouter()


@router.post("/dividir-resoluciones")
async def dividir_resoluciones(
    file: UploadFile = File(...),
    word_document: UploadFile = File(...)
):
    excel_path = str(UPLOADS_DIR / file.filename)
    with open(excel_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    word_path = str(UPLOADS_DIR / word_document.filename)
    with open(word_path, "wb") as f:
        shutil.copyfileobj(word_document.file, f)

    wb = load_excel(excel_path)
    ws = wb.active

    filenames = []
    max_row = ws.max_row
    for row in range(2, max_row + 1):
        val = ws.cell(row=row, column=3).value
        filenames.append(val)

    word_dir = os.path.dirname(word_path)
    pdf_folder = os.path.join(word_dir, "PDF")

    pdf_results = split_word_to_pdfs(word_path, pdf_folder, filenames)

    for i, result in enumerate(pdf_results):
        row_num = i + 2
        ws.cell(row=row_num, column=4).value = "Sí" if result["success"] else "No"

    success_count = sum(1 for r in pdf_results if r["success"])
    error_count = sum(1 for r in pdf_results if not r["success"])

    output_name = generate_output_filename(file.filename, "resoluciones")
    output_path = str(OUTPUTS_DIR / output_name)
    save_excel(wb, output_path)

    os.remove(excel_path)
    os.remove(word_path)

    return JSONResponse({
        "total_rows": len(pdf_results),
        "success_count": success_count,
        "error_count": error_count,
        "rows": pdf_results,
        "pdf_folder": pdf_folder,
        "download_url": f"/api/download/{output_name}"
    })
