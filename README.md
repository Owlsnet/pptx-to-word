# PPTX to Word Converter

A simple Flask web app that converts PPTX presentations into Word documents so they can be viewed or shared more easily.

## Features

- Upload a `.pptx` file and download a `.docx` file.
- Each slide becomes a Word heading with its text content.
- Arabic text is marked as right-to-left when detected.
- Lightweight UI for quick conversions.

## Local setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Then open `http://localhost:8000` in your browser.

## Free hosting options

- **Render / Railway**: Deploy this Flask app with a free web service and connect to a GitHub repo.
- **Fly.io**: Free tier for running a containerized Flask app.
- **Cloudflare Pages + Workers**: Good if you want to rework it into a serverless architecture.

Each platform has limits on free tiers, so choose based on expected usage.
