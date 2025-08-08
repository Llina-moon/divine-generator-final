from flask import Flask, render_template, request, send_from_directory
from docx import Document
from docx.shared import Pt, RGBColor
import io
import os
import uuid

app = Flask(__name__)

GENERATED_DIR = "generated"
os.makedirs(GENERATED_DIR, exist_ok=True)

VARIABLES = [
    "{ФИО}",
    "{Документ}",
    "{Адрес}",
    "{ИНН}",
    "{СНИЛС}",
    "{Банковские реквизиты}",
    "{Номер договора}",
    "{id artist / id contract}",
    "{Электронная почта лицензиара}",
    "{Дата}",
]

def _replace_in_paragraph(paragraph, mapping):
    # Простая замена по всем плейсхолдерам в рамках одного параграфа
    for run in paragraph.runs:
        text = run.text
        for k, v in mapping.items():
            if k in text:
                text = text.replace(k, v)
        run.text = text
        # (опционально) выставим шрифт
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)

def _replace_in_doc(doc, mapping):
    # Параграфы
    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)
    # Таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, mapping)
    return doc

def generate_file(template_path, values, out_prefix):
    doc = Document(template_path)
    _replace_in_doc(doc, values)
    fn = f"{out_prefix}_{uuid.uuid4().hex}.docx"
    out_path = os.path.join(GENERATED_DIR, fn)
    doc.save(out_path)
    return fn

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        values = {var: request.form.get(var.strip("{}"), "") for var in VARIABLES}
        # Генерация двух документов
        contract_name = generate_file("contract_template.docx", values, "contract")
        appendix_name = generate_file("appendix_template.docx", values, "appendix")
        # Переходим на страницу успеха с двумя ссылками
        return render_template("success.html", contract=contract_name, appendix=appendix_name)
    return render_template("form.html", variables=VARIABLES)

@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(GENERATED_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
