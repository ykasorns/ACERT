import io
import os
import zipfile
from datetime import datetime

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "AC Certs for print.pdf")
FONTS_DIR = os.path.join(BASE_DIR, "fonts")

# Load template into memory at startup (before Flask reloader spawns child process)
with open(TEMPLATE_PATH, "rb") as _f:
    TEMPLATE_BYTES = _f.read()

# Register Kanit fonts (same font used in the original certificate)
pdfmetrics.registerFont(TTFont("Kanit", os.path.join(FONTS_DIR, "Kanit-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Kanit-Bold", os.path.join(FONTS_DIR, "Kanit-Bold.ttf")))

# Page dimensions (A4 portrait, in points)
PAGE_WIDTH = 595.276
PAGE_HEIGHT = 841.89

# Text positions from example PDF analysis
# y = PAGE_HEIGHT - pdfplumber_bottom + descender_offset (baseline correction)
POS_NAME_Y = PAGE_HEIGHT - 396.4 + 12  # 457 — Kanit Bold 24pt, centered (corrected to baseline)
POS_COURSE_Y = PAGE_HEIGHT - 491.3     # 350.59 — Kanit Bold 21pt, centered
POS_DATE_Y = PAGE_HEIGHT - 578.4       # 263.49 — Kanit Regular 18pt, centered
POS_ACTCNO_X = 372.7                   # left-aligned
POS_ACTCNO_Y = PAGE_HEIGHT - 723.8     # 118.09 — Kanit Regular 18pt

# Colors from example PDF (normalized RGB)
COLOR_BLUE = (0.109804, 0.458824, 0.737255)   # name & course
COLOR_DARK = (0.137255, 0.121569, 0.12549)    # date & cert no.


def fit_text(c, font, max_size, min_size, text, max_width):
    """Return the largest font size that fits within max_width."""
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
    max_width = 440  # safe usable width for centered text

    # Name — Kanit Bold, up to 24pt, centered, blue
    size = fit_text(c, "Kanit-Bold", 24, 12, name, max_width)
    c.setFont("Kanit-Bold", size)
    c.setFillColorRGB(*COLOR_BLUE)
    c.drawCentredString(center_x, POS_NAME_Y, name)

    # Course — Kanit Bold, up to 21pt, centered, blue
    size = fit_text(c, "Kanit-Bold", 21, 10, course, max_width)
    c.setFont("Kanit-Bold", size)
    c.setFillColorRGB(*COLOR_BLUE)
    c.drawCentredString(center_x, POS_COURSE_Y, course)

    # Date — Kanit Regular 18pt, centered, dark (prefix with "On ")
    date_text = training_date if training_date.startswith("On ") else f"On {training_date}"
    size = fit_text(c, "Kanit", 18, 10, date_text, max_width)
    c.setFont("Kanit", size)
    c.setFillColorRGB(*COLOR_DARK)
    c.drawCentredString(center_x, POS_DATE_Y, date_text)

    # ACTC No. — Kanit Regular 18pt, left-aligned, dark
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
                    c for c in row["name"] if c.isalnum() or c in " -_"
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


if __name__ == "__main__":
    app.run(debug=True, port=5001)
