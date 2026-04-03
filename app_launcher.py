import io
import os
import re
import sys
import unicodedata
import zipfile
import threading
import webbrowser
from datetime import datetime

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# ── Path resolution (works both in dev and as PyInstaller .exe) ──────────────
def resource_path(relative_path):
    """Get absolute path to resource — works for dev and PyInstaller bundle."""
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


BASE_DIR = resource_path(".")
TEMPLATE_PATH = resource_path("AC Certs for print.pdf")
NCSA_16_PATH  = resource_path("NCSA-16March.pdf")
NCSA_19_PATH  = resource_path("NCSA-19March.pdf")
FONTS_DIR = resource_path("fonts")
TEMPLATES_DIR = resource_path("templates")

app = Flask(__name__, template_folder=TEMPLATES_DIR)

# Load template PDFs into memory at startup
with open(TEMPLATE_PATH, "rb") as _f:
    TEMPLATE_BYTES = _f.read()


def _remove_name_placeholder(pdf_bytes):
    """Remove 'Name Surename' placeholder text from NCSA template content stream."""
    import re
    from pypdf.generic import DecodedStreamObject, NameObject, ArrayObject, IndirectObject

    reader = PdfReader(io.BytesIO(pdf_bytes))
    page = reader.pages[0]
    contents = page.get("/Contents")
    if hasattr(contents, "get_object"):
        contents = contents.get_object()
    if hasattr(contents, "__iter__") and not isinstance(contents, (str, bytes)):
        raw = b"".join(obj.get_object().get_data() for obj in contents)
    else:
        raw = contents.get_data()

    # Remove the BT...ET block containing the placeholder name
    cleaned = re.sub(
        rb"BT\s*/TT0\s+1\s+Tf\s+33[^\n]*\n\[\(Name Sur\)[^\n]*\]TJ\s*ET",
        b"",
        raw,
    )

    # Rebuild PDF with modified content stream
    writer = PdfWriter()
    writer.append(io.BytesIO(pdf_bytes))
    new_stream = DecodedStreamObject()
    new_stream.set_data(cleaned)
    ref = writer._add_object(new_stream)
    writer.pages[0][NameObject("/Contents")] = ref
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


with open(NCSA_16_PATH, "rb") as _f:
    NCSA_16_BYTES = _remove_name_placeholder(_f.read())
with open(NCSA_19_PATH, "rb") as _f:
    NCSA_19_BYTES = _remove_name_placeholder(_f.read())

# Register Kanit fonts
pdfmetrics.registerFont(TTFont("Kanit", os.path.join(FONTS_DIR, "Kanit-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Kanit-Bold", os.path.join(FONTS_DIR, "Kanit-Bold.ttf")))

# Page dimensions (A4 portrait, in points)
PAGE_WIDTH = 595.276
PAGE_HEIGHT = 841.89

POS_NAME_Y = PAGE_HEIGHT - 396.4 + 12
POS_COURSE_Y = PAGE_HEIGHT - 491.3
POS_DATE_Y = PAGE_HEIGHT - 578.4
POS_ACTCNO_X       = 372.7
POS_CERTNO_LABEL_Y = 156    # aligned with signature baseline
POS_ACTCNO_Y       = 126    # aligned with (Training Director, ACinfotec)

COLOR_BLUE = (0.109804, 0.458824, 0.737255)
COLOR_DARK = (0.137255, 0.121569, 0.12549)

# ── NCSA certificate constants (A4 Landscape) ────────────────────────────────
NCSA_WIDTH  = 841.89
NCSA_HEIGHT = 595.28
NCSA_NAME_X = NCSA_WIDTH / 2   # centered
NCSA_NAME_Y = 332              # baseline (from pdfplumber: 595.28 - 263.49)
NCSA_NAME_COLOR = (0.2, 0.2, 0.2)


def fit_text(c, font, max_size, min_size, text, max_width):
    size = max_size
    while size >= min_size:
        c.setFont(font, size)
        if c.stringWidth(text, font, size) <= max_width:
            return size
        size -= 0.5
    return min_size


def create_text_overlay(name, actc_no, course, training_date):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    center_x = PAGE_WIDTH / 2
    max_width = 440

    size = fit_text(c, "Kanit-Bold", 24, 12, name, max_width)
    c.setFont("Kanit-Bold", size)
    c.setFillColorRGB(*COLOR_BLUE)
    c.drawCentredString(center_x, POS_NAME_Y, name)

    size = fit_text(c, "Kanit-Bold", 21, 10, course, max_width)
    c.setFont("Kanit-Bold", size)
    c.setFillColorRGB(*COLOR_BLUE)
    c.drawCentredString(center_x, POS_COURSE_Y, course)

    date_text = training_date if training_date.startswith("On ") else f"On {training_date}"
    size = fit_text(c, "Kanit", 18, 10, date_text, max_width)
    c.setFont("Kanit", size)
    c.setFillColorRGB(*COLOR_DARK)
    c.drawCentredString(center_x, POS_DATE_Y, date_text)

    # Cover original "Certification No." baked into template, then redraw at signature level
    c.setFillColorRGB(1, 1, 1)
    c.rect(338, 142, 220, 30, fill=1, stroke=0)

    # Calculate center X based on ACTC number width for proper centering
    c.setFont("Kanit", 18)
    actc_width = c.stringWidth(str(actc_no), "Kanit", 18)
    center_x = POS_ACTCNO_X + (actc_width / 2)

    # "Certification No." centered above ACTC number
    c.setFont("Kanit-Bold", 15)
    c.setFillColorRGB(*COLOR_DARK)
    c.drawCentredString(center_x, POS_CERTNO_LABEL_Y, "Certification No.")

    # ACTC No. — aligned with (Training Director, ACinfotec) level
    c.setFont("Kanit", 18)
    c.setFillColorRGB(*COLOR_DARK)
    c.drawString(POS_ACTCNO_X, POS_ACTCNO_Y, str(actc_no))

    c.save()
    packet.seek(0)
    return packet


def generate_certificate(name, actc_no, course, training_date):
    overlay_pdf = create_text_overlay(name, actc_no, course, training_date)
    template_reader = PdfReader(io.BytesIO(TEMPLATE_BYTES))
    overlay_reader = PdfReader(overlay_pdf)

    writer = PdfWriter()
    template_page = template_reader.pages[0]
    template_page.merge_page(overlay_reader.pages[0])
    writer.add_page(template_page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output


def parse_excel(file):
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        if not name or name.lower() == "nan":
            continue
        rows.append({
            "name": name,
            "actc_no": str(row.get("ACTC No.", "")).strip(),
            "course": str(row.get("Course", "")).strip(),
            "training_date": str(row.get("Training Date", "")).strip(),
            "company": str(row.get("Company", "")).strip(),
        })
    return rows


def is_thai(text):
    return any('\u0e00' <= c <= '\u0e7f' for c in text)


def generate_ncsa_certificate(name, template_bytes):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(NCSA_WIDTH, NCSA_HEIGHT))

    # Draw actual name
    size = fit_text(c, "Kanit", 33, 16, name, 620)
    c.setFont("Kanit", size)
    c.setFillColorRGB(*NCSA_NAME_COLOR)
    c.drawCentredString(NCSA_NAME_X, NCSA_NAME_Y, name)

    c.save()
    packet.seek(0)

    template_reader = PdfReader(io.BytesIO(template_bytes))
    overlay_reader  = PdfReader(packet)
    writer = PdfWriter()
    template_page = template_reader.pages[0]
    template_page.merge_page(overlay_reader.pages[0])
    writer.add_page(template_page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output


def _get_col(row, *names):
    for name in names:
        val = str(row.get(name, "")).strip()
        if val and val.lower() != "nan":
            return val
    return ""


def parse_excel_ncsa(file):
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        title    = _get_col(row, "คำนำหน้า", "Title", "Prefix")
        fullname = _get_col(row, "ชื่อ-นามสกุล", "ชื่อ นามสกุล", "Name", "ชื่อ")
        if not fullname:
            continue
        # title "-" ถือว่าว่าง
        if title == "-":
            title = ""
        # Thai name → prepend title, English name → ใช้ชื่อตรงๆ ไม่มี title
        if is_thai(fullname):
            full_name = f"{title} {fullname}".strip() if title else fullname
        else:
            full_name = fullname
        rows.append({"name": full_name})
    return rows


@app.route("/preview_ncsa", methods=["POST"])
def preview_ncsa():
    if "file" not in request.files:
        return jsonify({"error": "ไม่พบไฟล์"}), 400
    file = request.files["file"]
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น"}), 400
    try:
        rows = parse_excel_ncsa(file)
        return jsonify({"rows": rows, "count": len(rows)})
    except Exception as e:
        return jsonify({"error": f"อ่านไฟล์ไม่ได้: {e}"}), 500


@app.route("/generate_ncsa", methods=["POST"])
def generate_ncsa():
    if "file" not in request.files:
        return jsonify({"error": "ไม่พบไฟล์"}), 400
    file = request.files["file"]
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น"}), 400
    template_id    = request.form.get("template_id", "ncsa_16")
    template_bytes = NCSA_16_BYTES if template_id == "ncsa_16" else NCSA_19_BYTES
    try:
        rows = parse_excel_ncsa(file)
        if not rows:
            return jsonify({"error": "ไม่พบรายชื่อในไฟล์ Excel"}), 400

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, row in enumerate(rows, 1):
                pdf_data = generate_ncsa_certificate(row["name"], template_bytes)
                safe_name = "".join(
                    c for c in row["name"]
                    if unicodedata.category(c)[0] in ("L", "M", "N") or c in " -_"
                ).strip().replace(" ", "_")
                zf.writestr(f"{i:03d}_{safe_name}.pdf", pdf_data.read())

        zip_buffer.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"NCSA_Certificates_{timestamp}.zip",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    if "file" not in request.files:
        return jsonify({"error": "ไม่พบไฟล์"}), 400
    file = request.files["file"]
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น"}), 400
    try:
        rows = parse_excel(file)
        return jsonify({"rows": rows, "count": len(rows)})
    except Exception as e:
        return jsonify({"error": f"อ่านไฟล์ไม่ได้: {e}"}), 500


@app.route("/generate", methods=["POST"])
def generate():
    if "file" not in request.files:
        return jsonify({"error": "ไม่พบไฟล์"}), 400
    file = request.files["file"]
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น"}), 400
    try:
        rows = parse_excel(file)
        if not rows:
            return jsonify({"error": "ไม่พบรายชื่อในไฟล์ Excel"}), 400

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for row in rows:
                pdf_data = generate_certificate(
                    row["name"], row["actc_no"], row["course"], row["training_date"]
                )
                safe_name = "".join(
                    c for c in row["name"]
                    if unicodedata.category(c)[0] in ("L", "M", "N") or c in " -_"
                ).strip().replace(" ", "_")
                filename = f"{row['actc_no']}_{safe_name}.pdf"
                zf.writestr(filename, pdf_data.read())

        zip_buffer.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"Certificates_{timestamp}.zip",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def open_browser():
    webbrowser.open("http://localhost:5001")


if __name__ == "__main__":
    # Open browser after a short delay to let Flask start
    threading.Timer(1.5, open_browser).start()
    app.run(debug=False, port=5001, threaded=True)
