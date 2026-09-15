"""
FastAPI Web Server for OmniConverter Call Log Dashboard.
"""

from datetime import datetime
import io
import json
import os
from pathlib import Path
import tempfile
import zipfile
from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from converter import (
    carregar_dados,
    converter_todos,
    exportar_csv,
    exportar_excel,
    exportar_ics,
    exportar_json,
)

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "web"
STATIC_DIR = WEB_DIR if WEB_DIR.exists() else BASE_DIR
DADOS_DEFAULT = BASE_DIR / "dados.json"

app = FastAPI(
    title="OmniConverter - Call Log Converter & Analytics",
    description="Multi-format converter (Excel, CSV, JSON, ICS) and analytics dashboard for call history exports",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/data")
async def get_default_data():
    """Returns normalized dataset from local dados.json."""
    if not DADOS_DEFAULT.exists():
        sample_path = BASE_DIR / "sample_data.json"
        if sample_path.exists():
            return JSONResponse(content=carregar_dados(str(sample_path)))
        raise HTTPException(status_code=404, detail="Data file not found on server.")
    try:
        data = carregar_dados(str(DADOS_DEFAULT))
        return JSONResponse(content=data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/upload")
async def upload_file(request: Request):
    """Accepts uploaded file or raw JSON payload and returns normalized data."""
    content_type = request.headers.get("content-type", "")
    try:
        if "application/json" in content_type:
            raw_json = await request.json()
            data = carregar_dados(raw_json)
            return JSONResponse(content={"filename": "uploaded.json", "data": data})

        body = await request.body()
        raw_text = body.decode("utf-8", errors="ignore")
        raw_json = json.loads(raw_text)
        data = carregar_dados(raw_json)
        return JSONResponse(content={"filename": "uploaded.json", "data": data})
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to process file: {e}")


@app.post("/api/export/{fmt}")
async def export_format(
    fmt: str,
    payload: dict,
    title: str = Query(default="Call with {contact}"),
    filename: str = Query(default="call_logs")
):
    """
    Exports provided dataset in requested format:
    xlsx, csv, json, ics, or zip.
    """
    fmt = fmt.lower()
    if fmt not in ("xlsx", "csv", "json", "ics", "zip"):
        raise HTTPException(status_code=400, detail="Invalid format. Supported: xlsx, csv, json, ics, zip")

    try:
        data = carregar_dados(payload)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error parsing data: {e}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_base = os.path.join(tmpdir, filename)

        if fmt == "xlsx":
            out_file = f"{tmp_base}.xlsx"
            exportar_excel(data, out_file, event_title=title)
            with open(out_file, "rb") as f:
                content = f.read()
            return Response(
                content=content,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f'attachment; filename="{filename}.xlsx"'}
            )

        elif fmt == "csv":
            out_file = f"{tmp_base}.csv"
            exportar_csv(data, out_file)
            with open(out_file, "rb") as f:
                content = f.read()
            return Response(
                content=content,
                media_type="text/csv; charset=utf-8",
                headers={"Content-Disposition": f'attachment; filename="{filename}.csv"'}
            )

        elif fmt == "json":
            out_file = f"{tmp_base}.json"
            exportar_json(data, out_file)
            with open(out_file, "rb") as f:
                content = f.read()
            return Response(
                content=content,
                media_type="application/json; charset=utf-8",
                headers={"Content-Disposition": f'attachment; filename="{filename}.json"'}
            )

        elif fmt == "ics":
            out_file = f"{tmp_base}.ics"
            exportar_ics(data, out_file, event_title=title)
            with open(out_file, "rb") as f:
                content = f.read()
            return Response(
                content=content,
                media_type="text/calendar; charset=utf-8",
                headers={"Content-Disposition": f'attachment; filename="{filename}.ics"'}
            )

        elif fmt == "zip":
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                f_xlsx = f"{tmp_base}.xlsx"
                f_csv = f"{tmp_base}.csv"
                f_json = f"{tmp_base}.json"
                f_ics = f"{tmp_base}.ics"

                exportar_excel(data, f_xlsx, event_title=title)
                exportar_csv(data, f_csv)
                exportar_json(data, f_json)
                exportar_ics(data, f_ics, event_title=title)

                zf.write(f_xlsx, arcname=f"{filename}.xlsx")
                zf.write(f_csv, arcname=f"{filename}.csv")
                zf.write(f_json, arcname=f"{filename}.json")
                zf.write(f_ics, arcname=f"{filename}.ics")

            zip_buffer.seek(0)
            return Response(
                content=zip_buffer.getvalue(),
                media_type="application/zip",
                headers={"Content-Disposition": f'attachment; filename="{filename}_bundle.zip"'}
            )


# Mount static assets
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
else:
    @app.get("/")
    async def index_root():
        return HTMLResponse("<h1>OmniConverter</h1><p>Web frontend folder not found.</p>")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
