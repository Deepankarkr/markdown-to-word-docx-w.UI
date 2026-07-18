import io
import re
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from flask import Flask, render_template, request, send_file
from markdown import markdown
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024
ALLOWED_EXTENSIONS = {"md", "markdown", "txt"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def clean_name(filename: str) -> str:
    name = secure_filename(filename) or "document.md"
    return Path(name).stem or "document"


def add_bottom_border(paragraph, color="4F81BD", size="12"):
    p = paragraph._p
    p_pr = p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def shade_cell(cell, fill="D9EAF7"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.bold = bold
    paragraph.paragraph_format.space_after = Pt(0)


def setup_document(doc: Document):
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.12

    for level in range(1, 4):
        style = styles[f"Heading {level}"]
        style.font.name = "Calibri"
        style.font.bold = True
        style.font.color.rgb = RGBColor(31, 78, 121)
        style.paragraph_format.space_before = Pt(14 if level == 1 else 10)
        style.paragraph_format.space_after = Pt(6)

    if "Code Block" not in [s.name for s in styles]:
        code = styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
        code.font.name = "Consolas"
        code.font.size = Pt(9)
        code.font.color.rgb = RGBColor(45, 45, 45)
        code.paragraph_format.left_indent = Inches(0.25)
        code.paragraph_format.right_indent = Inches(0.15)
        code.paragraph_format.space_before = Pt(6)
        code.paragraph_format.space_after = Pt(8)

    if "Quote Block" not in [s.name for s in styles]:
        quote = styles.add_style("Quote Block", WD_STYLE_TYPE.PARAGRAPH)
        quote.font.italic = True
        quote.font.color.rgb = RGBColor(95, 95, 95)
        quote.paragraph_format.left_indent = Inches(0.3)
        quote.paragraph_format.space_before = Pt(6)
        quote.paragraph_format.space_after = Pt(8)


def add_title_page(doc: Document, title: str, subtitle: str | None = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(31, 78, 121)

    if subtitle:
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.space_after = Pt(20)
        sr = sp.add_run(subtitle)
        sr.font.size = Pt(12)
        sr.font.color.rgb = RGBColor(100, 100, 100)

    add_bottom_border(p)


def add_inline_text(paragraph, element):
    for child in element.children:
        if isinstance(child, NavigableString):
            paragraph.add_run(str(child))
        elif isinstance(child, Tag):
            if child.name == "br":
                paragraph.add_run().add_break()
                continue
            text = child.get_text()
            run = paragraph.add_run(text)
            if child.name in ["strong", "b"]:
                run.bold = True
            if child.name in ["em", "i"]:
                run.italic = True
            if child.name == "code":
                run.font.name = "Consolas"
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(168, 38, 38)


def add_table(doc: Document, table_tag: Tag):
    rows = table_tag.find_all("tr")
    if not rows:
        return

    max_cols = max(len(row.find_all(["th", "td"])) for row in rows)
    table = doc.add_table(rows=len(rows), cols=max_cols)
    table.style = "Table Grid"
    table.autofit = True

    for r_idx, row in enumerate(rows):
        cells = row.find_all(["th", "td"])
        for c_idx, cell in enumerate(cells):
            is_header = cell.name == "th" or r_idx == 0 and row.find("th") is not None
            target = table.cell(r_idx, c_idx)
            set_cell_text(target, cell.get_text(" ", strip=True), bold=is_header)
            if is_header:
                shade_cell(target)

    doc.add_paragraph()


def add_list(doc: Document, list_tag: Tag, ordered=False):
    style = "List Number" if ordered else "List Bullet"
    for li in list_tag.find_all("li", recursive=False):
        p = doc.add_paragraph(style=style)
        add_inline_text(p, li)


def add_markdown_to_doc(doc: Document, md_text: str, source_title: str | None = None):
    html = markdown(
        md_text,
        extensions=["tables", "fenced_code", "sane_lists", "nl2br"],
        output_format="html5",
    )
    soup = BeautifulSoup(html, "html.parser")

    if source_title:
        heading = doc.add_heading(source_title, level=1)
        add_bottom_border(heading, size="8")

    for element in soup.children:
        if not isinstance(element, Tag):
            continue

        if re.fullmatch(r"h[1-6]", element.name):
            level = min(int(element.name[1]), 3)
            p = doc.add_heading(level=level)
            add_inline_text(p, element)

        elif element.name == "p":
            p = doc.add_paragraph()
            add_inline_text(p, element)

        elif element.name == "ul":
            add_list(doc, element, ordered=False)

        elif element.name == "ol":
            add_list(doc, element, ordered=True)

        elif element.name == "pre":
            p = doc.add_paragraph(style="Code Block")
            p.add_run(element.get_text())

        elif element.name == "blockquote":
            p = doc.add_paragraph(style="Quote Block")
            p.add_run(element.get_text(" ", strip=True))

        elif element.name == "table":
            add_table(doc, element)

        elif element.name == "hr":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run("—" * 28)


def make_docx(md_text: str, title: str, include_title=True) -> io.BytesIO:
    doc = Document()
    setup_document(doc)
    if include_title:
        add_title_page(doc, title)
    add_markdown_to_doc(doc, md_text)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def make_merged_docx(files) -> io.BytesIO:
    doc = Document()
    setup_document(doc)
    add_title_page(doc, "Merged Markdown Document", f"Generated from {len(files)} Markdown file(s)")

    for index, file in enumerate(files):
        filename = file.filename
        md_text = file.read().decode("utf-8", errors="replace")
        if index > 0:
            doc.add_section(WD_SECTION.NEW_PAGE)
        add_markdown_to_doc(doc, md_text, source_title=clean_name(filename))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    mode = request.form.get("mode", "separate")
    valid_files = [f for f in files if f and f.filename and allowed_file(f.filename)]

    if not valid_files:
        return render_template("index.html", error="Please upload at least one .md, .markdown, or .txt file."), 400

    if mode == "merged":
        output = make_merged_docx(valid_files)
        return send_file(
            output,
            as_attachment=True,
            download_name="merged_markdown_document.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in valid_files:
            md_text = file.read().decode("utf-8", errors="replace")
            base = clean_name(file.filename)
            docx_buffer = make_docx(md_text, title=base)
            zf.writestr(f"{base}.docx", docx_buffer.getvalue())

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="converted_docx_files.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
