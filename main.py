"""
Turbo File Server — Railway
Genera PPTX y XLSX profesionales desde instrucciones de Claude.
Endpoint: POST /generate
"""

from flask import Flask, request, jsonify, send_file
import json, os, io, base64, traceback
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ── Security: simple token check ─────────────────────────────────────────
API_TOKEN = os.environ.get("TURBO_API_TOKEN", "changeme123")

def check_auth(req):
    token = req.headers.get("X-Turbo-Token") or req.args.get("token")
    return token == API_TOKEN

# ══════════════════════════════════════════════════════════════════════════
# PPTX GENERATOR
# ══════════════════════════════════════════════════════════════════════════
def generate_pptx(payload: dict) -> bytes:
    """
    payload = {
      "title": "SEC CAP Training",
      "subtitle": "Chris Guzman · April 2026",
      "theme": "dark",           # dark | light | blue
      "slides": [
        {
          "title": "Slide Title",
          "bullets": ["point 1", "point 2"],
          "notes": "Speaker notes here"
        }
      ]
    }
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # ── Theme colors ──────────────────────────────────────────────────────
    themes = {
        "dark":  {"bg": "1A1A2E", "accent": "E94560", "text": "EAEAEA", "sub": "A0A0B0"},
        "light": {"bg": "FFFFFF", "accent": "2563EB", "text": "1F2937", "sub": "6B7280"},
        "blue":  {"bg": "0F3460", "accent": "E94560", "text": "FFFFFF", "sub": "B0C4DE"},
    }
    t = themes.get(payload.get("theme", "dark"), themes["dark"])

    def rgb(hex_str):
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    blank_layout = prs.slide_layouts[6]  # blank

    # ── TITLE SLIDE ───────────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = rgb(t["bg"])

    # Accent bar left
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.5), Inches(7.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = rgb(t["accent"])
    bar.line.fill.background()

    # Title
    tf = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(11), Inches(1.8))
    p = tf.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = payload.get("title", "Presentation")
    run.font.size = Pt(44); run.font.bold = True
    run.font.color.rgb = rgb(t["text"]); p.alignment = PP_ALIGN.LEFT

    # Subtitle
    tf2 = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(11), Inches(0.8))
    p2 = tf2.text_frame.paragraphs[0]
    r2 = p2.add_run()
    r2.text = payload.get("subtitle", "")
    r2.font.size = Pt(20); r2.font.color.rgb = rgb(t["sub"])

    # ── CONTENT SLIDES ────────────────────────────────────────────────────
    for slide_data in payload.get("slides", []):
        slide = prs.slides.add_slide(blank_layout)
        bg = slide.background.fill; bg.solid()
        bg.fore_color.rgb = rgb(t["bg"])

        # Top accent bar
        bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(0.12))
        bar.fill.solid(); bar.fill.fore_color.rgb = rgb(t["accent"])
        bar.line.fill.background()

        # Slide number bottom right
        num_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.9), Inches(1.5), Inches(0.4))
        np = num_box.text_frame.paragraphs[0]
        nr = np.add_run()
        nr.text = str(payload["slides"].index(slide_data) + 2)
        nr.font.size = Pt(11); nr.font.color.rgb = rgb(t["sub"])
        np.alignment = PP_ALIGN.RIGHT

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(12), Inches(0.85))
        tp = title_box.text_frame.paragraphs[0]
        tr = tp.add_run()
        tr.text = slide_data.get("title", "")
        tr.font.size = Pt(28); tr.font.bold = True
        tr.font.color.rgb = rgb(t["text"])

        # Divider line
        line = slide.shapes.add_shape(1, Inches(0.6), Inches(1.2), Inches(11.5), Inches(0.04))
        line.fill.solid(); line.fill.fore_color.rgb = rgb(t["accent"])
        line.line.fill.background()

        # Bullets
        bullets = slide_data.get("bullets", [])
        if bullets:
            content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.4), Inches(11.5), Inches(5.5))
            tf = content_box.text_frame
            tf.word_wrap = True
            for i, bullet in enumerate(bullets):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.space_before = Pt(6)
                run = p.add_run()
                # Detect sub-bullets (start with spaces or -)
                if bullet.startswith("  ") or bullet.startswith("- "):
                    run.text = "    › " + bullet.lstrip(" -").strip()
                    run.font.size = Pt(16)
                    run.font.color.rgb = rgb(t["sub"])
                else:
                    run.text = "▸  " + bullet.strip()
                    run.font.size = Pt(18)
                    run.font.color.rgb = rgb(t["text"])
                    run.font.bold = False

        # Speaker notes
        notes = slide_data.get("notes", "")
        if notes:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = notes

    # ── CLOSING SLIDE ─────────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.background.fill; bg.solid()
    bg.fore_color.rgb = rgb(t["accent"])

    tf = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11), Inches(1.5))
    p = tf.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "Thank You"
    r.font.size = Pt(54); r.font.bold = True
    r.font.color.rgb = rgb("FFFFFF"); p.alignment = PP_ALIGN.CENTER

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════
# XLSX GENERATOR
# ══════════════════════════════════════════════════════════════════════════
def generate_xlsx(payload: dict) -> bytes:
    """
    payload = {
      "title": "Loan Portfolio Report",
      "sheets": [
        {
          "name": "Sheet1",
          "headers": ["Col A", "Col B", "Col C"],
          "rows": [["val1", "val2", "val3"]],
          "totals": true,          # add SUM row at bottom
          "col_widths": [20,15,15] # optional
        }
      ],
      "summary": "Optional summary text"
    }
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    # ── Style helpers ─────────────────────────────────────────────────────
    HEADER_FILL  = PatternFill("solid", fgColor="1A3A5C")
    ACCENT_FILL  = PatternFill("solid", fgColor="E94560")
    ALT_FILL     = PatternFill("solid", fgColor="F0F4F8")
    TOTAL_FILL   = PatternFill("solid", fgColor="E8F0FE")
    HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    TITLE_FONT   = Font(name="Arial", bold=True, color="1A3A5C", size=14)
    DATA_FONT    = Font(name="Arial", size=10)
    TOTAL_FONT   = Font(name="Arial", bold=True, color="1A3A5C", size=11)
    CENTER       = Alignment(horizontal="center", vertical="center")
    LEFT         = Alignment(horizontal="left",   vertical="center")
    thin         = Side(style="thin", color="D0D7DE")
    BORDER       = Border(left=thin, right=thin, top=thin, bottom=thin)

    for sheet_def in payload.get("sheets", []):
        ws = wb.create_sheet(title=sheet_def.get("name", "Report")[:31])
        headers  = sheet_def.get("headers", [])
        rows     = sheet_def.get("rows", [])
        do_total = sheet_def.get("totals", False)
        widths   = sheet_def.get("col_widths", [])

        # Title row
        ws.row_dimensions[1].height = 30
        title_cell = ws.cell(row=1, column=1, value=payload.get("title", "Report"))
        title_cell.font = TITLE_FONT
        title_cell.alignment = LEFT
        if headers:
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1,   end_column=len(headers))

        # Accent bar row 2
        for col in range(1, len(headers)+1):
            c = ws.cell(row=2, column=col, value="")
            c.fill = ACCENT_FILL

        # Headers row 3
        ws.row_dimensions[3].height = 22
        for col_idx, header in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=header)
            c.font = HEADER_FONT; c.fill = HEADER_FILL
            c.alignment = CENTER; c.border = BORDER

        # Data rows starting row 4
        for row_idx, row_data in enumerate(rows, 4):
            ws.row_dimensions[row_idx].height = 18
            for col_idx, value in enumerate(row_data, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=value)
                c.font = DATA_FONT; c.border = BORDER
                c.alignment = CENTER if isinstance(value, (int,float)) else LEFT
                if row_idx % 2 == 0:
                    c.fill = ALT_FILL

        # Totals row
        if do_total and rows:
            total_row = len(rows) + 4
            ws.row_dimensions[total_row].height = 22
            ws.cell(row=total_row, column=1, value="TOTAL").font = TOTAL_FONT
            ws.cell(row=total_row, column=1).fill = TOTAL_FILL
            for col_idx in range(2, len(headers)+1):
                col_letter = get_column_letter(col_idx)
                formula = f"=SUM({col_letter}4:{col_letter}{total_row-1})"
                c = ws.cell(row=total_row, column=col_idx, value=formula)
                c.font = TOTAL_FONT; c.fill = TOTAL_FILL
                c.alignment = CENTER; c.border = BORDER

        # Column widths
        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            if widths and col_idx <= len(widths):
                ws.column_dimensions[col_letter].width = widths[col_idx-1]
            else:
                ws.column_dimensions[col_letter].width = max(len(str(header))+4, 14)

        # Freeze header
        ws.freeze_panes = "A4"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════════════════
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "Turbo File Server"})


@app.route("/generate", methods=["POST"])
def generate():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401

    try:
        body = request.get_json(force=True)
        file_type = body.get("type", "").lower()  # "pptx" or "xlsx"
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
            return jsonify({"error": f"Unknown type: {file_type}. Use 'pptx' or 'xlsx'"}), 400

        # Return as base64 JSON (easier for n8n to handle)
        encoded = base64.b64encode(file_bytes).decode("utf-8")
        return jsonify({
            "success": True,
            "type": ext,
            "filename": f"{filename}.{ext}",
            "mime": mime,
            "size_bytes": len(file_bytes),
            "data_base64": encoded
        })

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
