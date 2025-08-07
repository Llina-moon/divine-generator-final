from docx.shared import Pt, RGBColor
from flask import Flask, render_template, request, send_file
from docx import Document
import io
import re

app = Flask(__name__)

VARIABLES = [
    "{ФИО}",
    "{Паспортные данные}",
    "{Адрес}",
    "{ИНН}",
    "{СНИЛС}",
    "{Банковские реквизиты}",
    "{Номер договора}",
    "{id artist / id contract}",
    "{Электронная почта лицензиара}"
]

def replace_variables_in_docx(template_path, values_dict):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        updated_text = full_text
        for var in VARIABLES:
            updated_text = updated_text.replace(var, values_dict.get(var, var))

        # Если текст изменился, перезаписываем весь абзац одним run
        if full_text != updated_text:
            # Удаляем старые runs
            for _ in range(len(paragraph.runs)):
                paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)
            # Добавляем новый run
            run = paragraph.add_run(updated_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)

    return doc

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        values = {var: request.form.get(var.strip("{}"), "") for var in VARIABLES}

        contract_doc = replace_variables_in_docx("contract_template.docx", values)
        appendix_doc = replace_variables_in_docx("appendix_template.docx", values)

        buffer = io.BytesIO()
        contract_doc.save(buffer)
        contract_filename = "Лицензионный договор.docx"
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=contract_filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    return render_template("form.html", variables=VARIABLES)

if __name__ == "__main__":
    app.run(debug=True)
