"""
FastAPI — backend do gerador de Relatório de Peças V3A.
"""
import os
import sys
import uuid
import tempfile
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# Garante que o diretório pai está no path
sys.path.insert(0, str(Path(__file__).parent.parent))

app = FastAPI(title="Relatório de Peças V3A")

# Serve arquivos estáticos (frontend)
STATIC_DIR = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Diretório de arquivos temporários gerados
OUTPUT_DIR = Path(tempfile.gettempdir()) / "v3a_reports"
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/")
async def index():
    """Serve o frontend."""
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/api/generate")
async def generate_report(
    subtitle: str = Form(...),
    files: List[UploadFile] = File(...),
):
    """
    Recebe subtítulo + imagens, retorna o PPTX gerado.

    Os agentes são acionados nesta ordem:
      Orquestrador → Builder → Diretor de Arte
    """
    from app.agents import OrchestratorAgent

    if not files:
        raise HTTPException(status_code=400, detail="Nenhuma imagem enviada.")

    # Salva uploads em arquivos temporários
    pieces = []
    tmp_files = []
    try:
        for upload in files:
            suffix = Path(upload.filename).suffix or ".png"
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            content = await upload.read()
            tmp.write(content)
            tmp.close()
            tmp_files.append(tmp.name)
            pieces.append({
                "filename": upload.filename,
                "image_path": tmp.name,
            })

        # Aciona o pipeline de agentes
        orchestrator = OrchestratorAgent()
        result = orchestrator.process(subtitle=subtitle, pieces=pieces)

        if not result.success:
            raise HTTPException(status_code=422, detail=result.error)

        # Salva PPTX em disco para download
        output_filename = f"relatorio_{uuid.uuid4().hex[:8]}.pptx"
        output_path = OUTPUT_DIR / output_filename
        output_path.write_bytes(result.data["pptx_bytes"])

        return JSONResponse({
            "success": True,
            "download_url": f"/download/{output_filename}",
            "art_director_message": result.data.get("art_director_message", ""),
            "issues": result.data.get("issues", []),
        })

    finally:
        # Limpa temporários
        for f in tmp_files:
            try:
                os.unlink(f)
            except Exception:
                pass


@app.get("/download/{filename}")
async def download(filename: str):
    """Serve o PPTX gerado para download."""
    path = OUTPUT_DIR / filename
    if not path.exists() or not path.name.endswith(".pptx"):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")
    return FileResponse(
        path=str(path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="Relatorio_Pecas_V3A.pptx",
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app.main:app", host="0.0.0.0", port=port, reload=False)
