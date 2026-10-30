# Markdown to Word Flask App

![Alt text](./image.png)

Upload multiple Markdown files and download them as Word documents.

## Features

- Upload multiple `.md`, `.markdown`, or `.txt` files
- Convert all files into one merged `.docx`
- Convert each file into a separate `.docx` inside a ZIP
- Download as editable Word (`.docx`) or themed PDF (`.pdf`)
- Handwritten-notes Word theme with large, roomy MV Boli typography
- Colorful headings, bullets, highlighted emphasis, tables, and sketch-style callouts
- Deterministic visual variation for a human touch without changing Markdown content
- Separate PDF tool that merges two or more PDFs locally without rasterizing pages

## Setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

On Windows:

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:5000
```
