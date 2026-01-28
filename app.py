from __future__ import annotations

import io
import os
import shutil
import json
import subprocess
import urllib.parse
import urllib.request
import tempfile
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
from flask import Flask, Response, flash, redirect, render_template, request
from werkzeug.utils import secure_filename
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.shared import Inches

ALLOWED_EXTENSIONS = {".pptx"}

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET", "dev-secret")
TURNSTILE_SITE_KEY = os.environ.get("TURNSTILE_SITE_KEY")
TURNSTILE_SECRET_KEY = os.environ.get("TURNSTILE_SECRET_KEY")


@app.get("/")
def index() -> str:
    return render_template("index.html", turnstile_site_key=TURNSTILE_SITE_KEY)


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

        images = convert_pptx_to_images(pptx_path, tmp_path, dpi=220)
        if not images:
            raise RuntimeError("No slides were rendered from the PPTX.")

        document = Document()
        for index, image_path in enumerate(images):
            width_px, height_px = get_image_dimensions(image_path)
            width_in = width_px / 220
            height_in = height_px / 220

            if index == 0:
                section = document.sections[0]
            else:
                section = document.add_section(WD_SECTION_START.NEW_PAGE)

            set_section_page(section, width_in, height_in)
            document.add_picture(
                str(image_path), width=Inches(width_in), height=Inches(height_in)
            )

        output = io.BytesIO()
        document.save(output)
        output.seek(0)
    return output


def verify_turnstile(token: str) -> bool:
    data = urllib.parse.urlencode(
        {"secret": TURNSTILE_SECRET_KEY, "response": token}
    ).encode()
    try:
        with urllib.request.urlopen(
            "https://challenges.cloudflare.com/turnstile/v0/siteverify", data=data
        ) as response:
            payload = json.loads(response.read().decode("utf-8"))
            return bool(payload.get("success"))
    except Exception:  # noqa: BLE001
        return False


def convert_pptx_to_images(
    pptx_path: Path, output_dir: Path, *, dpi: int
) -> list[Path]:
    libreoffice = find_libreoffice()
    pdf_path = convert_pptx_to_pdf(libreoffice, pptx_path, output_dir)
    return pdf_to_pngs(pdf_path, output_dir, dpi)


def find_libreoffice() -> str:
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
    return libreoffice


def convert_pptx_to_pdf(libreoffice: str, pptx_path: Path, output_dir: Path) -> Path:
    result = subprocess.run(
        [
            libreoffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
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

    pdf_path = output_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        pdfs = list(output_dir.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice did not produce a PDF.")
        pdf_path = pdfs[0]
    return pdf_path


def pdf_to_pngs(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    scale = dpi / 72.0
    matrix = fitz.Matrix(scale, scale)

    images: list[Path] = []
    with fitz.open(str(pdf_path)) as doc:
        for index in range(doc.page_count):
            page = doc.load_page(index)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = output_dir / f"slide_{index + 1:03d}.png"
            pixmap.save(str(image_path))
            images.append(image_path)
    return images


def get_image_dimensions(image_path: Path) -> tuple[int, int]:
    with fitz.open(str(image_path)) as doc:
        page = doc.load_page(0)
        pixmap = page.get_pixmap(alpha=False)
        return pixmap.width, pixmap.height


def set_section_page(section, width_in: float, height_in: float) -> None:
    section.page_width = Inches(width_in)
    section.page_height = Inches(height_in)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
    if TURNSTILE_SITE_KEY and TURNSTILE_SECRET_KEY:
        token = request.form.get("cf-turnstile-response")
        if not token:
            flash("Please complete the captcha challenge.", "error")
            return redirect("/")
        if not verify_turnstile(token):
            flash("Captcha verification failed. Please try again.", "error")
            return redirect("/")
