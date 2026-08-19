from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pathlib import Path
import os

from .routers import mover_archivo, exportar_correos, correo_masivo, dividir_resoluciones, crear_carpetas

app = FastAPI(title="Sistema de Correspondencia", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

frontend_path = Path(__file__).parent.parent / "frontend"
if frontend_path.exists():
    app.mount("/static", StaticFiles(directory=str(frontend_path)), name="static")

app.include_router(mover_archivo.router, prefix="/api", tags=["Mover Archivo"])
app.include_router(exportar_correos.router, prefix="/api", tags=["Exportar Correos"])
app.include_router(correo_masivo.router, prefix="/api", tags=["Correo Masivo"])
app.include_router(dividir_resoluciones.router, prefix="/api", tags=["Dividir Resoluciones"])
app.include_router(crear_carpetas.router, prefix="/api", tags=["Crear Carpetas"])


@app.get("/")
async def root():
    index_path = frontend_path / "index.html"
    if index_path.exists():
        return FileResponse(str(index_path))
    return {"message": "Sistema de Correspondencia API"}


@app.get("/api/health")
async def health():
    result = {"status": "ok", "outlook": False, "word": False}

    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        result["outlook"] = True
    except Exception:
        pass

    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        result["word"] = True
    except Exception:
        pass

    return result


uploads_dir = Path(__file__).parent / "uploads"
outputs_dir = Path(__file__).parent / "outputs"
uploads_dir.mkdir(exist_ok=True)
outputs_dir.mkdir(exist_ok=True)


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = outputs_dir / filename
    if file_path.exists():
        return FileResponse(
            str(file_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename
        )
    return {"error": "Archivo no encontrado"}
