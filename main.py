"""
Turbo File Server — Railway v4
Genera PPTX y XLSX. Analiza CSV enviado como texto desde n8n.
Endpoints: POST /generate, POST /analyze, GET /health
"""

from flask import Flask, request, jsonify
import os, io, base64, traceback
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

API_TOKEN = os.environ.get("TURBO_API_TOKEN", "changeme123")

def check_auth(req):
    token = req.headers.get("X-Turbo-Token") or req.args.get("token")
    return token == API_TOKEN


# ══════════════════════════════════════════════════════════════════════════
# PPTX GENERATOR
# ══════════════════════════════════════════════════════════════════════════
def generate_pptx(payload):
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    themes = {
        "dark":  {"bg": "1A1A2E", "accent": "E94560", "text": "EAEAEA", "sub": "A0A0B0"},
        "light": {"bg": "FFFFFF", "accent": "2563EB", "text": "1F2937", "sub": "6B7280"},
        "blue":  {"bg": "0F3460", "accent": "E94560", "text": "FFFFFF",  "sub": "B0C4DE"},
    }
    t = themes.get(payload.get("theme", "dark"), themes["dark"])
    def rgb(h):
        h = h.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = rgb(t["bg"])
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.5), Inches(7.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = rgb(t["accent"]); bar.line.fill.background()
    tf = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(11), Inches(1.8))
    p = tf.text_frame.paragraphs[0]; run = p.add_run()
    run.text = payload.get("title", "Presentation")
    run.font.size = Pt(44); run.font.bold = True
    run.font.color.rgb = rgb(t["text"]); p.alignment = PP_ALIGN.LEFT
    tf2 = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(11), Inches(0.8))
    p2 = tf2.text_frame.paragraphs[0]; r2 = p2.add_run()
    r2.text = payload.get("subtitle", "")
    r2.font.size = Pt(20); r2.font.color.rgb = rgb(t["sub"])
    for idx, slide_data in enumerate(payload.get("slides", [])):
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = rgb(t["bg"])
        bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(0.12))
        bar.fill.solid(); bar.fill.fore_color.rgb = rgb(t["accent"]); bar.line.fill.background()
        num_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.9), Inches(1.5), Inches(0.4))
        np2 = num_box.text_frame.paragraphs[0]; nr = np2.add_run()
        nr.text = str(idx + 2); nr.font.size = Pt(11); nr.font.color.rgb = rgb(t["sub"])
        np2.alignment = PP_ALIGN.RIGHT
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(12), Inches(0.85))
        tp = title_box.text_frame.paragraphs[0]; tr = tp.add_run()
        tr.text = slide_data.get("title", "")
        tr.font.size = Pt(28); tr.font.bold = True; tr.font.color.rgb = rgb(t["text"])
        line = slide.shapes.add_shape(1, Inches(0.6), Inches(1.2), Inches(11.5), Inches(0.04))
        line.fill.solid(); line.fill.fore_color.rgb = rgb(t["accent"]); line.line.fill.background()
        bullets = slide_data.get("bullets", [])
        if bullets:
            content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.4), Inches(11.5), Inches(5.5))
            tf3 = content_box.text_frame; tf3.word_wrap = True
            for i, bullet in enumerate(bullets):
                p3 = tf3.paragraphs[0] if i == 0 else tf3.add_paragraph()
                p3.space_before = Pt(6); run = p3.add_run()
                if bullet.startswith("  ") or bullet.startswith("- "):
                    run.text = "    › " + bullet.lstrip(" -").strip()
                    run.font.size = Pt(16); run.font.color.rgb = rgb(t["sub"])
                else:
                    run.text = "▸  " + bullet.strip()
                    run.font.size = Pt(18); run.font.color.rgb = rgb(t["text"])
        notes = slide_data.get("notes", "")
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
    slide = prs.slides.add_slide(blank)
    bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = rgb(t["accent"])
    tf4 = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11), Inches(1.5))
    p4 = tf4.text_frame.paragraphs[0]; r4 = p4.add_run()
    r4.text = "Thank You"; r4.font.size = Pt(54); r4.font.bold = True
    r4.font.color.rgb = rgb("FFFFFF"); p4.alignment = PP_ALIGN.CENTER
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════
# XLSX GENERATOR
# ══════════════════════════════════════════════════════════════════════════
def generate_xlsx(payload):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    HEADER_FILL = PatternFill("solid", fgColor="1A3A5C")
    ACCENT_FILL = PatternFill("solid", fgColor="E94560")
    ALT_FILL    = PatternFill("solid", fgColor="F0F4F8")
    TOTAL_FILL  = PatternFill("solid", fgColor="E8F0FE")
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    TITLE_FONT  = Font(name="Arial", bold=True, color="1A3A5C", size=14)
    DATA_FONT   = Font(name="Arial", size=10)
    TOTAL_FONT  = Font(name="Arial", bold=True, color="1A3A5C", size=11)
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT   = Alignment(horizontal="left",   vertical="center")
    thin   = Side(style="thin", color="D0D7DE")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    for sheet_def in payload.get("sheets", []):
        ws = wb.create_sheet(title=sheet_def.get("name", "Report")[:31])
        headers  = sheet_def.get("headers", [])
        rows     = sheet_def.get("rows", [])
        do_total = sheet_def.get("totals", False)
        widths   = sheet_def.get("col_widths", [])
        ws.row_dimensions[1].height = 30
        title_cell = ws.cell(row=1, column=1, value=payload.get("title", "Report"))
        title_cell.font = TITLE_FONT; title_cell.alignment = LEFT
        if headers:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        for col in range(1, len(headers)+1):
            ws.cell(row=2, column=col, value="").fill = ACCENT_FILL
        ws.row_dimensions[3].height = 22
        for col_idx, header in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=header)
            c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER
        for row_idx, row_data in enumerate(rows, 4):
            ws.row_dimensions[row_idx].height = 18
            for col_idx, value in enumerate(row_data, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=value)
                c.font = DATA_FONT; c.border = BORDER
                c.alignment = CENTER if isinstance(value, (int,float)) else LEFT
                if row_idx % 2 == 0: c.fill = ALT_FILL
        if do_total and rows:
            total_row = len(rows) + 4
            ws.row_dimensions[total_row].height = 22
            ws.cell(row=total_row, column=1, value="TOTAL").font = TOTAL_FONT
            ws.cell(row=total_row, column=1).fill = TOTAL_FILL
            for col_idx in range(2, len(headers)+1):
                col_letter = get_column_letter(col_idx)
                c = ws.cell(row=total_row, column=col_idx,
                            value=f"=SUM({col_letter}4:{col_letter}{total_row-1})")
                c.font = TOTAL_FONT; c.fill = TOTAL_FILL; c.alignment = CENTER; c.border = BORDER
        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = (
                widths[col_idx-1] if widths and col_idx <= len(widths)
                else max(len(str(header))+4, 14)
            )
        ws.freeze_panes = "A4"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════
# SHARED: build_analysis
# ══════════════════════════════════════════════════════════════════════════
def build_analysis(df, filename, instruction):
    df.columns = [
        str(c) if not str(c).startswith("Unnamed") else f"Col_{i+1}"
        for i, c in enumerate(df.columns)
    ]
    df = df.dropna(how="all")
    rows, columns = df.shape
    insights = []
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    for col in numeric_cols[:5]:
        insights.append(
            f"{col}: Total={df[col].sum():,.2f} | "
            f"Avg={df[col].mean():,.2f} | "
            f"Max={df[col].max():,.2f} | "
            f"Min={df[col].min():,.2f}"
        )
    nulls = df.isnull().sum()
    null_cols = nulls[nulls > 0]
    if len(null_cols) > 0:
        insights.append(f"⚠️ Valores nulos en: {', '.join(null_cols.index.tolist()[:5])}")
    else:
        insights.append("✅ Sin valores nulos")
    insights.append(f"Total de registros: {rows:,}")
    data_preview = df.head(5).to_string(index=False, max_cols=8)
    col_list = ", ".join(df.columns.tolist()[:10])
    summary = f"Columnas: {col_list}" + (
        f" ... (+{len(df.columns)-10} más)" if len(df.columns) > 10 else ""
    )
    return {
        "success":      True,
        "filename":     filename,
        "rows":         rows,
        "columns":      columns,
        "col_names":    df.columns.tolist(),
        "insights":     insights,
        "summary":      summary,
        "data_preview": data_preview,
        "instruction":  instruction
    }


# ══════════════════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════════════════
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "Turbo File Server v4"})


@app.route("/generate", methods=["POST"])
def generate():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401
    try:
        body      = request.get_json(force=True)
        file_type = body.get("type", "").lower()
        payload   = body.get("payload", {})
        filename  = body.get("filename", "output")
        if file_type == "pptx":
            file_bytes = generate_pptx(payload)
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ext  = "pptx"
        elif file_type == "xlsx":
            file_bytes = generate_xlsx(payload)
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ext  = "xlsx"
        else:
            return jsonify({"error": f"Unknown type: {file_type}"}), 400
        encoded = base64.b64encode(file_bytes).decode("utf-8")
        return jsonify({
            "success": True, "type": ext, "filename": f"{filename}.{ext}",
            "mime": mime, "size_bytes": len(file_bytes), "data_base64": encoded
        })
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/analyze", methods=["POST"])
def analyze():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401
    try:
        body = request.get_json(force=True, silent=True)
        if not body:
            return jsonify({"error": "No se recibió payload JSON"}), 400

        # ── Recibir CSV como texto desde n8n ─────────────────────────────
        csv_text    = body.get("csv_text", "")
        instruction = body.get("instruction", "Analiza este archivo")
        filename    = body.get("filename", "archivo.csv")

        if not csv_text:
            return jsonify({"error": "No se recibió csv_text"}), 400

        # ── Leer CSV desde el string de texto ────────────────────────────
        df = None
        errors = []

        for enc in ["utf-8", "latin-1", "cp1252"]:
            try:
                buf = io.StringIO(csv_text)
                df = pd.read_csv(buf)
                if len(df.columns) > 0 and len(df) > 0:
                    break
            except Exception as e:
                errors.append(f"csv/{enc}: {e}")

        if df is None or len(df.columns) == 0:
            return jsonify({
                "error": "No se pudo leer el CSV.",
                "attempts": errors[:3]
            }), 500

        return jsonify(build_analysis(df, filename, instruction))

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
