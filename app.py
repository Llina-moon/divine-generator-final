from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt, RGBColor  # ✅ добавлено
import io

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
        for var in VARIABLES:
            if var in paragraph.text:
                inline = paragraph.runs
                for i in range(len(inline)):
                    if var in inline[i].text:
                        inline[i].text = inline[i].text.replace(var, values_dict.get(var, var))
                        inline[i].font.name = 'Times New Roman'
                        inline[i].font.size = Pt(11)  # ✅ исправлено
                        inline[i].font.color.rgb = RGBColor(0, 0, 0)  # ✅ исправлено
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
