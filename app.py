
import os
import re
import json
import base64
import secrets
import tempfile
import subprocess

from io import BytesIO
from pathlib import Path
from datetime import datetime, date, timezone
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, jsonify, send_file, flash
)
from flask_sqlalchemy import SQLAlchemy
from flask_wtf.csrf import CSRFProtect
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# --------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "templates" / "Rendicion Gastos.xlsx"

app = Flask(__name__)

secret_key = os.getenv("FLASK_SECRET_KEY")

if not secret_key and os.getenv("RENDER"):
    raise RuntimeError("Falta configurar FLASK_SECRET_KEY en Render.")

app.secret_key = secret_key or secrets.token_hex(32)

database_url = os.getenv("DATABASE_URL", "sqlite:///rendiciones_local.db")

# Render puede entregar la URL con el prefijo postgres://
if database_url.startswith("postgres://"):
    database_url = database_url.replace(
        "postgres://", "postgresql+psycopg://", 1
    )
elif database_url.startswith("postgresql://"):
    database_url = database_url.replace(
        "postgresql://", "postgresql+psycopg://", 1
    )

app.config.update(
    SQLALCHEMY_DATABASE_URI=database_url,
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    MAX_CONTENT_LENGTH=20 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=bool(os.getenv("RENDER")),
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=8 * 60 * 60,
)

db = SQLAlchemy(app)
csrf = CSRFProtect(app)


# --------------------------------------------------
# BASE DE DATOS: SOLO CUENTAS DE USUARIO
# --------------------------------------------------

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(254), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        nullable=False
    )

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


with app.app_context():
    db.create_all()


# --------------------------------------------------
# SESIÓN Y AUTENTICACIÓN
# --------------------------------------------------

def current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    return db.session.get(User, user_id)


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Debes iniciar sesión."}), 401
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


@app.route("/registro", methods=["GET", "POST"])
def register():
    if session.get("user_id"):
        return redirect(url_for("index"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
            flash("Ingresa un correo electrónico válido.")
            return render_template("auth.html", mode="register")

        if len(password) < 10:
            flash("La contraseña debe tener al menos 10 caracteres.")
            return render_template("auth.html", mode="register")

        if User.query.filter_by(email=email).first():
            flash("No fue posible crear la cuenta con esos datos.")
            return render_template("auth.html", mode="register")

        user = User(email=email)
        user.set_password(password)

        try:
            db.session.add(user)
            db.session.commit()
        except Exception:
            db.session.rollback()
            flash("No fue posible crear la cuenta. Inténtalo nuevamente.")
            return render_template("auth.html", mode="register")

        session.clear()
        session["user_id"] = user.id
        session.permanent = True

        return redirect(url_for("index"))

    return render_template("auth.html", mode="register")


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("index"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        user = User.query.filter_by(email=email).first()

        if not user or not user.check_password(password):
            flash("Correo o contraseña incorrectos.")
            return render_template("auth.html", mode="login")

        session.clear()
        session["user_id"] = user.id
        session.permanent = True

        return redirect(url_for("index"))

    return render_template("auth.html", mode="login")


@app.route("/logout", methods=["POST"])
@login_required
def logout():
    session.clear()
    return redirect(url_for("login"))


# --------------------------------------------------
# CABECERAS DE SEGURIDAD Y PRIVACIDAD
# --------------------------------------------------

@app.after_request
def security_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


# --------------------------------------------------
# PÁGINA PRINCIPAL
# --------------------------------------------------

@app.route("/")
@login_required
def index():
    return render_template("index.html", user=current_user())


# --------------------------------------------------
# EXTRACCIÓN DE DATOS CON OPENAI
# --------------------------------------------------

ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}

MIME_TYPES = {
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "png": "image/png",
    "webp": "image/webp",
}


def extract_with_openai(image_bytes, extension):
    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key:
        raise ValueError("Falta configurar OPENAI_API_KEY.")

    client = OpenAI(api_key=api_key)

    image_base64 = base64.b64encode(image_bytes).decode("utf-8")
    mime_type = MIME_TYPES[extension]

    prompt = """
    Analiza la imagen de un comprobante de gasto.

    Extrae exclusivamente estos datos:

    - fecha: fecha del documento en formato YYYY-MM-DD.
    - numero_documento: número de boleta, factura o comprobante.
    - descripcion: descripción breve del gasto.
    - monto: total FINAL pagado, incluyendo impuestos.
    - confianza: nivel de confianza de la lectura entre 0 y 1.

    Reglas:
    - No confundas el subtotal o neto con el total final.
    - No inventes información que no sea visible.
    - Si un dato no se puede leer, utiliza una cadena vacía.
    - El monto debe ser un número, sin símbolos de moneda.
    - Devuelve solamente un objeto JSON válido.
    """

    response = client.responses.create(
        model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {
                        "type": "input_image",
                        "image_url": (
                            f"data:{mime_type};base64,{image_base64}"
                        )
                    }
                ]
            }
        ]
    )

    content = response.output_text.strip()

    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)

    data = json.loads(content)

    return {
        "fecha": str(data.get("fecha", "")),
        "numero_documento": str(data.get("numero_documento", "")),
        "descripcion": str(data.get("descripcion", "")),
        "monto": data.get("monto", ""),
        "confianza": data.get("confianza", 0)
    }


@app.route("/api/extract", methods=["POST"])
@login_required
def extract():
    if "file" not in request.files:
        return jsonify({"error": "No se recibió ninguna fotografía."}), 400

    file = request.files["file"]

    if not file.filename:
        return jsonify({"error": "El archivo no tiene nombre."}), 400

    extension = secure_filename(file.filename).rsplit(".", 1)[-1].lower()

    if extension not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Formato de imagen no permitido."}), 400

    image_bytes = file.read()

    if not image_bytes:
        return jsonify({"error": "La imagen está vacía."}), 400

    try:
        result = extract_with_openai(image_bytes, extension)
        return jsonify(result)
    except Exception:
        app.logger.exception("Error durante la extracción de comprobante.")
        return jsonify({
            "error": "No se pudo procesar la imagen. Inténtalo nuevamente."
        }), 500


# --------------------------------------------------
# GENERACIÓN DE PDF
# --------------------------------------------------

def parse_date(value):
    if not value:
        return None

    value = str(value).strip()

    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None


def parse_amount(value):
    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        amount = float(value)
    else:
        cleaned = str(value).strip()
        cleaned = cleaned.replace("$", "").replace(" ", "")

        # Formato habitual chileno: 125.000 o 125.000,50
        if "," in cleaned and "." in cleaned:
            if cleaned.rfind(",") > cleaned.rfind("."):
                cleaned = cleaned.replace(".", "").replace(",", ".")
            else:
                cleaned = cleaned.replace(",", "")
        elif "," in cleaned:
            cleaned = cleaned.replace(",", ".")
        elif cleaned.count(".") > 1:
            cleaned = cleaned.replace(".", "")

        amount = float(cleaned)

    if amount < 0:
        raise ValueError("El monto no puede ser negativo.")

    return amount


def get_expense_sheet(workbook):
    expected = "r. d .gastos"

    for sheet in workbook.worksheets:
        if sheet.title.strip().lower() == expected:
            return sheet

    raise ValueError(
        "No se encontró la hoja 'r. d .gastos' en la plantilla."
    )


@app.route("/api/pdf", methods=["POST"])
@login_required
def generate_pdf():
    data = request.get_json(silent=True) or {}

    nombre = str(data.get("nombre", "")).strip()
    fecha_rendicion = parse_date(data.get("fecha"))
    gastos = data.get("gastos", [])

    if not nombre:
        return jsonify({"error": "Debes ingresar el nombre."}), 400

    if not fecha_rendicion:
        return jsonify({"error": "La fecha de rendición no es válida."}), 400

    if not isinstance(gastos, list) or not 1 <= len(gastos) <= 39:
        return jsonify({
            "error": "La rendición debe tener entre 1 y 39 comprobantes."
        }), 400

    if not TEMPLATE_PATH.exists():
        return jsonify({
            "error": "No se encontró la plantilla Excel."
        }), 500

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            workbook = load_workbook(TEMPLATE_PATH)
            ws = get_expense_sheet(workbook)

            # Encabezado
            ws["D3"] = nombre
            ws["F3"] = fecha_rendicion

            # Limpiar las líneas de comprobantes
            for row in range(5, 44):
                for column in range(2, 7):
                    ws.cell(row=row, column=column).value = None

            # Completar datos de cada comprobante
            for index, gasto in enumerate(gastos):
                row = 5 + index

                fecha = parse_date(gasto.get("fecha", ""))
                numero = str(
                    gasto.get("numero_documento", "")
                ).strip()
                descripcion = str(
                    gasto.get("descripcion", "")
                ).strip()
                monto = parse_amount(gasto.get("monto", 0))

                ws.cell(row=row, column=2).value = fecha
                ws.cell(row=row, column=3).value = numero
                ws.cell(row=row, column=4).value = descripcion
                ws.cell(row=row, column=5).value = monto

                # Columna Contabilidad: siempre vacía
                ws.cell(row=row, column=6).value = None

                for column in range(2, 6):
                    cell = ws.cell(row=row, column=column)
                    cell.alignment = Alignment(
                        horizontal=cell.alignment.horizontal,
                        vertical="center",
                        wrap_text=True,
                        shrink_to_fit=True
                    )

            # Total de gastos
            ws["E47"] = "=SUM(E5:E43)"

            # Configuración de impresión
            ws.print_area = "B1:F48"
            ws.page_setup.orientation = "portrait"
            ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True

            ws.page_margins.left = 0.25
            ws.page_margins.right = 0.25
            ws.page_margins.top = 0.35
            ws.page_margins.bottom = 0.35

            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True

            excel_path = temp_path / "rendicion.xlsx"
            workbook.save(excel_path)
            workbook.close()

            # Convertir Excel a PDF usando LibreOffice
            result = subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(temp_path),
                    str(excel_path)
                ],
                capture_output=True,
                text=True,
                timeout=90
            )

            pdf_path = temp_path / "rendicion.pdf"

            if result.returncode != 0 or not pdf_path.exists():
                app.logger.error(
                    "Error LibreOffice: %s", result.stderr
                )
                raise RuntimeError("No fue posible convertir el documento.")

            pdf_bytes = pdf_path.read_bytes()

        # El directorio temporal ya se eliminó.
        # El PDF se entrega desde memoria, sin guardarlo permanentemente.
        return send_file(
            BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="Rendicion_Gastos.pdf",
            max_age=0
        )

    except Exception:
        app.logger.exception("Error generando PDF.")
        return jsonify({
            "error": "No fue posible generar el PDF. Revisa los datos e inténtalo nuevamente."
        }), 500


# --------------------------------------------------
# ERRORES
# --------------------------------------------------

@app.errorhandler(413)
def file_too_large(error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "El archivo supera el tamaño permitido."}), 413
    return "El archivo supera el tamaño permitido.", 413


@app.errorhandler(400)
def bad_request(error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "Solicitud no válida."}), 400
    return "Solicitud no válida.", 400


# --------------------------------------------------
# INICIO LOCAL
# --------------------------------------------------

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", 5000)),
        debug=False
    )
