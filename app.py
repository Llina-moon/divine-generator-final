from flask import Flask, render_template, request, send_from_directory, redirect, url_for
from docx import Document
from docx.shared import Pt, RGBColor
import os
import re
from datetime import date

app = Flask(__name__)

TEMPLATES = {
    "contract": "contract_template.docx",
    "appendix": "appendix_template.docx",
}
OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

PLACEHOLDER_RE = re.compile(r"\{[^{}]+\}")  # всё внутри фигурных скобок

def extract_placeholders(paths):
    """Собираем уникальные плейсхолдеры из всех шаблонов."""
    found = set()
    for p in paths:
        doc = Document(p)
        # параграфы
        for par in doc.paragraphs:
            found.update(PLACEHOLDER_RE.findall(par.text))
        # таблицы
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for par in cell.paragraphs:
                        found.update(PLACEHOLDER_RE.findall(par.text))
    return sorted(found)

def replace_in_paragraph(paragraph, mapping):
    """Надёжная замена по параграфу: собираем весь текст, меняем, заново заливаем в один run."""
    full_text = "".join(run.text for run in paragraph.runs) or paragraph.text
    if not full_text:
        return
    new_text = full_text
    for k, v in mapping.items():
        new_text = new_text.replace(k, v)
    if new_text != full_text:
        # чистим рансы и записываем заново
        for _ in range(len(paragraph.runs)):
            paragraph.runs[0].clear() ; paragraph.runs[0].text = ""
            paragraph.runs[0].font.name = "Times New Roman"
            paragraph.runs[0].font.size = Pt(11)
            paragraph.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            # удаляем оставшиеся runs
            if len(paragraph.runs) > 1:
                paragraph.runs[1]._element.getparent().remove(paragraph.runs[1]._element)
        paragraph.clear()
        run = paragraph.add_run(new_text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)

def replace_in_doc(doc, mapping):
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, mapping)
    return doc

@app.route("/", methods=["GET", "POST"])
def index():
    # Собираем все плейсхолдеры из шаблонов, чтобы показать их в форме
    placeholders = extract_placeholders(TEMPLATES.values())

    if request.method == "POST":
        # значения из формы; добавим «умолчалки»
        values = {}
        for ph in placeholders:
            key = ph.strip("{}")
            values[ph] = request.form.get(key, "").strip()

        # авто-значения, если пусто
        if not values.get("{Дата}"):
            values["{Дата}"] = date.today().strftime("%d.%m.%Y")

        # Генерация документов
        generated_files = {}
        for key, path in TEMPLATES.items():
            doc = Document(path)
            doc = replace_in_doc(doc, values)
            out_name = "Лицензионный договор.docx" if key == "contract" else "Приложение №1.1.docx"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            doc.save(out_path)
            generated_files[key] = out_name

        return render_template(
            "success.html",
            contract=generated_files["contract"],
            appendix=generated_files["appendix"],
        )

    # GET — показать форму
    # Преобразуем вид для удобного рендера (без фигурных скобок)
    clean_list = [ph.strip("{}") for ph in placeholders]
    return render_template("form.html", variables=clean_list)

@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
