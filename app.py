from docx.shared import Pt, RGBColor
from flask import Flask, render_template, request, send_file
from docx import Document
import io
import re

app = Flask(__name__)

def extract_variables_from_template(template_path):
    doc = Document(template_path)
    text = "\n".join([p.text for p in doc.paragraphs])
    return sorted(set(re.findall(r"\{[^}]+\}", text)))  # список уникальных переменных

def replace_variables_in_docx(template_path, values_dict):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        updated_text = full_text
        for key, val in values_dict.items():
            updated_text = updated_text.replace(key, val)

        if full_text != updated_text:
            for _ in range(len(paragraph.runs)):
                paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)
            run = paragraph.add_run(updated_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)

    return doc

@app.route("/", methods=["GET", "POST"])
def index():
    variables = extract_variables_from_template("contract_template.docx")

    if request.method == "POST":
        values = {var: request.form.get(var.strip("{}"), "") for var in variables}

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

    return render_template("form.html", variables=variables)

if __name__ == "__main__":
    app.run(debug=True)
