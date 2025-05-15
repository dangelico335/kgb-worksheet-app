
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import tempfile

app = Flask(__name__)

# Supported chord image filenames
supported_files = [
    "C_Guitar.png", "C_Piano.png", "C_Bass.png", "Am_Guitar.png", "Am_Piano.png",
    "G_Guitar.png", "G_Piano.png", "G_Bass.png", "D_Guitar.png", "D_Piano.png", "D_Bass.png",
    "Em_Guitar.png", "Em_Piano.png", "E_Guitar.png", "E_Piano.png", "E_Bass.png"
    # add more as needed
]

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        title = request.form["title"]
        composer = request.form["composer"]
        key = request.form["key"]
        instruments = request.form.getlist("instruments")

        # Collect sections
        sections = []
        for i in range(1, 4):
            sec_name = request.form.get(f"section{i}_name")
            sec_chords = request.form.get(f"section{i}_chords")
            if sec_name and sec_chords:
                chords = [c.strip() for c in sec_chords.split(",")]
                sections.append((sec_name, chords))

        doc = Document()
        sec = doc.sections[0]
        sec.orientation = 1
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
        sec.top_margin = Inches(0.4)
        sec.bottom_margin = Inches(0.4)
        sec.left_margin = Inches(0.4)
        sec.right_margin = Inches(0.4)

        # Title and composer
        doc.add_paragraph().add_run(title).font.name = "Bodoni MT Black"
        composer_p = doc.add_paragraph()
        composer_run = composer_p.add_run(composer)
        composer_run.font.name = "Bodoni MT Black"
        composer_run.font.size = Pt(9)
        composer_p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        for i, (section_name, chords) in enumerate(sections):
            if i > 0:
                doc.add_page_break()

            doc.add_paragraph().add_run(section_name).bold = True

            chord_groups = [[c] if "/" not in c else c.split("/") for c in chords]
            grouped_rows = []
            row, count = [], 0
            for g in chord_groups:
                row.append(g)
                count += 1
                if count == 4:
                    grouped_rows.append(row)
                    row, count = [], 0
            if row:
                grouped_rows.append(row)

            for group in grouped_rows:
                flat_chords = [c for g in group for c in g]
                table = doc.add_table(rows=len(instruments)+1, cols=len(flat_chords)+1)
                table.cell(0, 0).text = ""
                col = 1
                for g in group:
                    if len(g) == 2:
                        cell = table.cell(0, col)
                        cell.merge(table.cell(0, col + 1))
                        para = cell.paragraphs[0]
                        para.add_run(f"{g[0]} / {g[1]}").bold = True
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        col += 2
                    else:
                        para = table.cell(0, col).paragraphs[0]
                        para.add_run(g[0]).bold = True
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        col += 1

                for idx, instr in enumerate(instruments):
                    row = table.rows[idx + 1]
                    row.cells[0].text = instr
                    row.cells[0].paragraphs[0].runs[0].bold = True
                    col = 1
                    for g in group:
                        for chord in g:
                            chord = chord.replace("add9", "")
                            base = chord.rstrip("m") if instr == "Bass" else chord
                            match = next((f for f in supported_files if f.startswith(base) and instr in f), None)
                            img_path = f"static/images/{match}" if match else None
                            p = row.cells[col].paragraphs[0]
                            if img_path and os.path.exists(img_path):
                                p.add_run().add_picture(img_path, width=Inches(1.0))
                            else:
                                p.add_run("[Missing]")
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            col += 1

        doc.add_paragraph().add_run(f"Key: {key}").bold = True
        for _ in range(4):
            doc.add_paragraph("_________________________")

        temp_path = tempfile.mktemp(suffix=".docx")
        doc.save(temp_path)
        return send_file(temp_path, as_attachment=True)

    return render_template("index.html")

# ✅ Required for Render
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
