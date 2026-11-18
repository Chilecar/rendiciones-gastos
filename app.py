import os, json, base64, tempfile, subprocess, shutil
from datetime import date, datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", APP_DIR / "templates" / "Rendicion Gastos.xlsx"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", APP_DIR / "outputs"))
OUTPUT_DIR.mkdir(exist_ok=True)
ALLOWED = {"png", "jpg", "jpeg", "webp"}
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # límite de carga por petición
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-secret")

def extract_with_openai(image_bytes, mime_type):
    """Optional vision extraction. Requires OPENAI_API_KEY."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta configurar OPENAI_API_KEY para activar el reconocimiento IA.")
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    data_url = f"data:{mime_type};base64,{base64.b64encode(image_bytes).decode()}"
    prompt = """
Lee el comprobante chileno de gasto. Devuelve SOLO JSON válido con estas claves:
fecha (YYYY-MM-DD o null), numero_documento (string o null),
descripcion (string breve que identifique los productos o servicios),
monto (número entero en pesos chilenos, o null), confianza (0 a 1).
No inventes datos. Si un dato no aparece o no es legible, usa null.
No confundas subtotal ni neto con el total: monto debe ser el TOTAL FINAL a pagar.
No devuelvas clasificación ni información contable.
"""
    response = client.responses.create(
        model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
        input=[{"role":"user","content":[
            {"type":"input_text","text":prompt},
            {"type":"input_image","image_url":data_url}
        ]}],
    )
    raw = response.output_text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    return json.loads(raw)

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/api/extract")
def extract():
    f = request.files.get("image")
    if not f or not f.filename:
        return jsonify(error="Selecciona una fotografía."), 400
    ext = Path(secure_filename(f.filename)).suffix.lower().lstrip(".")
    if ext not in ALLOWED:
        return jsonify(error="Formato no admitido. Usa JPG, PNG o WEBP."), 400
    try:
        result = extract_with_openai(f.read(), f.mimetype or "image/jpeg")
        result["fecha"] = result.get("fecha") or ""
        result["numero_documento"] = result.get("numero_documento") or ""
        result["descripcion"] = result.get("descripcion") or ""
        result["monto"] = result.get("monto") if result.get("monto") is not None else ""
        return jsonify(result)
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.post("/api/pdf")
def create_pdf():
    payload = request.get_json(force=True)
    nombre = str(payload.get("nombre", "")).strip()
    rows = payload.get("gastos", [])
    if not nombre:
        return jsonify(error="Ingresa el nombre de quien rinde."), 400
    if not rows:
        return jsonify(error="Agrega al menos un comprobante."), 400
    if len(rows) > 39:
        return jsonify(error="La rendición admite un máximo de 39 comprobantes."), 400
    if not TEMPLATE_PATH.exists():
        return jsonify(error="No se encontró la plantilla Excel oficial."), 500

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        xlsx = td / "Rendicion_completada.xlsx"
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb["r. d .gastos"] if "r. d .gastos" in wb.sheetnames else wb[wb.sheetnames[0]]
        # El PDF de rendición contiene únicamente la hoja oficial de gastos.
        # Las hojas auxiliares del libro original se conservan en el archivo fuente,
        # pero no se imprimen en el PDF de rendición.
        for sheet in wb.worksheets:
            sheet.sheet_state = "visible" if sheet.title == ws.title else "hidden"
        wb.active = wb.sheetnames.index(ws.title)

        # Mantener diseño, bordes, encabezados, distribución y firmas de la plantilla.
        ws["D3"] = nombre
        ws["F3"] = date.today()
        for r in range(5, 46):
            for col in range(2, 7):
                ws.cell(r, col).value = None

        from copy import copy
        from openpyxl.styles import Alignment
        for idx, item in enumerate(rows[:39], start=5):
            ws.cell(idx, 2).value = parse_date(item.get("fecha"))
            ws.cell(idx, 3).value = str(item.get("numero_documento", "") or "")
            ws.cell(idx, 4).value = str(item.get("descripcion", "") or "")
            amount = item.get("monto", "")
            try:
                amount = float(str(amount).replace(".", "").replace(",", "."))
            except Exception:
                amount = None
            ws.cell(idx, 5).value = amount
            # La columna de contabilidad se mantiene deliberadamente vacía.
            ws.cell(idx, 6).value = None
            # Ajuste de texto y reducción automática para no desbordar las celdas.
            for col in range(2, 7):
                cell = ws.cell(idx, col)
                old = copy(cell.alignment)
                cell.alignment = Alignment(horizontal=old.horizontal, vertical="center",
                    text_rotation=old.textRotation, wrap_text=True, shrink_to_fit=True,
                    indent=old.indent, relativeIndent=old.relativeIndent,
                    justifyLastLine=old.justifyLastLine, readingOrder=old.readingOrder)
            ws.row_dimensions[idx].height = 23 if len(str(item.get("descripcion", ""))) > 52 else 18

        ws["E47"] = "=SUM(E5:E43)"
        ws.print_area = "B1:F48"
        # Hoja carta, una sola página: ajustar escala y márgenes para impresión.
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.22
        ws.page_margins.right = 0.22
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25
        ws.page_margins.header = 0.1
        ws.page_margins.footer = 0.1
        wb.save(xlsx)

        # LibreOffice converts the filled original workbook to PDF.
        outdir = td / "pdf"
        outdir.mkdir()
        try:
            subprocess.run(
                ["libreoffice", f"-env:UserInstallation=file://{td / 'lo-profile'}", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(xlsx)],
                check=True, capture_output=True, text=True, timeout=60
            )
        except FileNotFoundError:
            return jsonify(error="El servidor necesita LibreOffice instalado para exportar PDF."), 500
        except subprocess.CalledProcessError as e:
            return jsonify(error=f"No se pudo generar el PDF: {e.stderr}"), 500

        pdf_path = outdir / "Rendicion_completada.pdf"
        if not pdf_path.exists():
            return jsonify(error="LibreOffice no generó el PDF esperado."), 500
        final_path = OUTPUT_DIR / f"Rendicion_{datetime.now():%Y%m%d_%H%M%S}.pdf"
        shutil.copy2(pdf_path, final_path)
    preview = request.args.get("preview") == "1"
    return send_file(final_path, as_attachment=not preview,
                     download_name=final_path.name, mimetype="application/pdf")

def parse_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except Exception:
        return str(value)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False)
