from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Inches
from io import BytesIO

app = Flask(__name__)

@app.route("/prepare_docx", methods=["POST"])
def prepare_docx():
    data = request.get_json(silent=True) or {}

    text = data.get("text", "")
    margins = data.get("margins", {})
    header = data.get("header", {})
    footer1 = data.get("footer1", {})
    footer2 = data.get("footer2", {})
    cover_page = data.get("cover_page", {})

    doc = Document()

    for section in doc.sections:
        if "top" in margins:
            section.top_margin = Inches(float(margins["top"]))
        if "bottom" in margins:
            section.bottom_margin = Inches(float(margins["bottom"]))
        if "left" in margins:
            section.left_margin = Inches(float(margins["left"]))
        if "right" in margins:
            section.right_margin = Inches(float(margins["right"]))

    if cover_page:
        doc.add_heading(cover_page.get("title", ""), level=1)
        doc.add_paragraph(f"Submitted to: {cover_page.get('submitted_to','')}")
        doc.add_paragraph(f"Submitted by: {cover_page.get('submitted_by','')}")
        doc.add_page_break()

    if header.get("content"):
        for sec in doc.sections:
            hdr = sec.header.paragraphs[0]
            hdr.text = header["content"]

    if footer1.get("content"):
        for sec in doc.sections:
            f = sec.footer.paragraphs[0]
            f.text = footer1["content"]

    doc.add_paragraph(text)

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="output.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    app.run(port=10000, host="0.0.0.0")
