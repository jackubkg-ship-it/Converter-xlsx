import io
import os
import tempfile

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from converter import list_sheets, convert_sheet, guess_period_label

app = FastAPI(title="Format Konvertoru")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html", context={})


@app.post("/api/inspect")
async def inspect(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Zəhmət olmasa .xlsx fayl seçin")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        sheets = list_sheets(tmp_path)
    except Exception as exc:
        raise HTTPException(400, f"Fayl oxuna bilmədi: {exc}") from exc
    finally:
        os.unlink(tmp_path)

    if not sheets:
        raise HTTPException(400, "Bu faylda köhnə format başlıqları (S/S, Tarix, DQ №-si...) tapılmadı")

    period_guess = guess_period_label(file.filename)
    suggested_filename = period_guess.replace(" ", "_") if period_guess else "converted"

    return {
        "sheets": sheets,
        "suggested_contractor": sheets[0]["name"],
        "suggested_filename": suggested_filename,
    }


@app.post("/api/convert")
async def convert(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    contractor_name: str = Form(...),
    client_name: str = Form(default="Salyan Oil"),
):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        wb_out = convert_sheet(tmp_path, sheet_name, contractor_name.strip(), client_name.strip())
    except Exception as exc:
        raise HTTPException(400, str(exc)) from exc
    finally:
        os.unlink(tmp_path)

    buffer = io.BytesIO()
    wb_out.save(buffer)
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="converted.xlsx"'},
    )
