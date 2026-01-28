from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, flash, redirect, render_template, request
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

    original_filename = upload.filename
    filename = secure_filename(original_filename)
    ext = Path(original_filename).suffix.lower()
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
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        pptx_path = tmp_path / "upload.pptx"
        pptx_path.write_bytes(pptx_bytes.getvalue())

        images = convert_pptx_to_images(pptx_path, tmp_path)
        if not images:
            raise RuntimeError("No slides were rendered from the PPTX.")

        document = Document()
        section = document.sections[0]
        max_width = section.page_width - section.left_margin - section.right_margin

        for index, image_path in enumerate(images):
            document.add_picture(str(image_path), width=max_width)
            if index < len(images) - 1:
                document.add_page_break()

        output = io.BytesIO()
        document.save(output)
        output.seek(0)
        return output


def convert_pptx_to_images(pptx_path: Path, output_dir: Path) -> list[Path]:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice and os.name == "nt":
        possible_paths = [
            Path("C:/Program Files/LibreOffice/program/soffice.exe"),
            Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        ]
        for path in possible_paths:
            if path.exists():
                libreoffice = str(path)
                break
    if not libreoffice:
        raise RuntimeError(
            "LibreOffice is required to render slides. Please install it on the server."
        )

    result = subprocess.run(
        [
            libreoffice,
            "--headless",
            "--convert-to",
            "png",
            "--outdir",
            str(output_dir),
            str(pptx_path),
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        raise RuntimeError(
            "Slide rendering failed. "
            f"{result.stderr.strip() or result.stdout.strip()}"
        )

    images = sorted(output_dir.glob("*.png"))
    return images


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
