# PPTX to Word Converter

A simple Flask web app that converts PPTX presentations into Word documents so they can be viewed or shared more easily.

## Features

- Upload a `.pptx` file and download a `.docx` file.
- Slides are rendered as images so layout and graphics stay intact.
- Uses LibreOffice + PDF rendering to keep every slide included.
- Lightweight UI for quick conversions.

## Local setup

This app requires **LibreOffice** to render PPTX slides into images.
To enable Cloudflare Turnstile captcha, set `TURNSTILE_SITE_KEY` and
`TURNSTILE_SECRET_KEY` in your environment.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Then open `http://localhost:8000` in your browser.

## Install LibreOffice

You need LibreOffice on the machine running the app. Use one of the options below:

- **macOS (Homebrew)**:
  ```bash
  brew install --cask libreoffice
  ```
- **Ubuntu/Debian**:
  ```bash
  sudo apt-get update
  sudo apt-get install libreoffice
  ```
- **Windows**:
  Download the installer from https://www.libreoffice.org/download/download/ and run it.
  If the app still cannot find LibreOffice, make sure it's installed in the default
  `C:\Program Files\LibreOffice\program\soffice.exe` path or update your PATH.

## Getting the files

Use either method below:

1. **Git (recommended)**:
   ```bash
   git clone <your-repo-url>
   cd pptx-to-word
   ```
2. **Download ZIP**:
   - Click the green **Code** button on GitHub.
   - Choose **Download ZIP**, then unzip it on your computer.

## Free hosting options

- **Render / Railway**: Deploy this Flask app with a free web service and connect to a GitHub repo.
- **Fly.io**: Free tier for running a containerized Flask app.
- **Cloudflare Pages + Workers**: Good if you want to rework it into a serverless architecture.

Each platform has limits on free tiers, so choose based on expected usage.
