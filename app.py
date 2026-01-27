from __future__ import annotations

import io
import os
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, flash, redirect, render_template, request
from pptx import Presentation
from werkzeug.utils import secure_filename
from docx import Document

ALLOWED_EXTENSIONS = {".pptx"}

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET", "dev-secret")


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.post("/convert")
def convert() -> Response:
    upload = request.files.get("pptx_file")
    if upload is None or upload.filename == "":
        flash("Please choose a PPTX file to upload.", "error")
        return redirect("/")

    filename = secure_filename(upload.filename)
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        flash("Only .pptx files are supported.", "error")
        return redirect("/")

    try:
        pptx_bytes = io.BytesIO(upload.read())
        docx_bytes = convert_pptx_to_docx(pptx_bytes)
    except Exception as exc:  # noqa: BLE001
        flash(f"Conversion failed: {exc}", "error")
        return redirect("/")

    timestamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
    download_name = f"{Path(filename).stem}-{timestamp}.docx"
    return Response(
        docx_bytes.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={download_name}"},
    )


def convert_pptx_to_docx(pptx_bytes: io.BytesIO) -> io.BytesIO:
    presentation = Presentation(pptx_bytes)
    document = Document()
    document.add_heading("PPTX Export", level=1)

    for slide_index, slide in enumerate(presentation.slides, start=1):
        slide_title = None
        slide_content = []
        slide_has_arabic = False

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text.strip()
            if not text:
                continue
            slide_has_arabic = slide_has_arabic or contains_arabic(text)
            if slide_title is None:
                slide_title = text
            else:
                slide_content.append(text)

        heading_text = f"Slide {slide_index}"
        if slide_title:
            heading_text += f": {slide_title}"

        heading = document.add_heading(heading_text, level=2)
        if slide_has_arabic:
            apply_rtl(heading)
        if slide_content:
            for paragraph in slide_content:
                para = document.add_paragraph(paragraph)
                if contains_arabic(paragraph):
                    apply_rtl(para)
        else:
            document.add_paragraph("(No text content found on this slide)")

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output


def contains_arabic(text: str) -> bool:
    return any("\u0600" <= char <= "\u06ff" for char in text)


def apply_rtl(paragraph) -> None:
    paragraph.paragraph_format.right_to_left = True
    for run in paragraph.runs:
        run.rtl = True


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
