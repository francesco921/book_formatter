# book_formatter.py
import argparse
import datetime
import os
import re
import sys
from typing import List, Tuple, Optional

# DOCX
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# PDF input opzionale
try:
    from pdfminer.high_level import extract_text as pdf_extract_text
except Exception:
    pdf_extract_text = None

# ReportLab per PDF nativo
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm, inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, PageBreak, Spacer, TableOfContents, Frame, PageTemplate
from reportlab.platypus.flowables import KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import fonts

HeadingItem = Tuple[int, str]  # (level, text)

# -------------------------------
# Parsing sorgente DOCX e PDF
# -------------------------------
def parse_docx_input(path: str) -> Tuple[List[HeadingItem], List[Tuple[int, List[str]]]]:
    doc = Document(path)
    headings: List[HeadingItem] = []
    content_blocks: List[Tuple[int, List[str]]] = []
    current_level: Optional[int] = None
    current_buffer: List[str] = []

    def flush():
        nonlocal current_level, current_buffer
        if current_level is not None:
            content_blocks.append((current_level, current_buffer))
        current_level = None
        current_buffer = []

    h1_pattern = re.compile(r"^\s*\d+\.\s+.+")
    h2_pattern = re.compile(r"^\s*\d+\.\d+\s+.+")
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue
        style_name = (p.style.name if p.style is not None else "")
        level_detected: Optional[int] = None
        if style_name in ("Heading 1", "Titolo 1"):
            level_detected = 1
        elif style_name in ("Heading 2", "Titolo 2"):
            level_detected = 2
        else:
            if h2_pattern.match(text):
                level_detected = 2
            elif h1_pattern.match(text):
                level_detected = 1

        if level_detected:
            flush()
            headings.append((level_detected, text))
            current_level = level_detected
        else:
            if current_level is None:
                current_level = 1
                headings.append((1, "Capitolo introduttivo"))
            current_buffer.append(text)
    flush()
    return headings, content_blocks


def parse_pdf_input(path: str) -> Tuple[List[HeadingItem], List[Tuple[int, List[str]]]]:
    if pdf_extract_text is None:
        raise RuntimeError("pdfminer.six non installato. Esegui: pip install pdfminer.six")
    raw = pdf_extract_text(path)
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    h1 = re.compile(r"^\s*\d+\.\s+.+")
    h2 = re.compile(r"^\s*\d+\.\d+\s+.+")
    headings: List[HeadingItem] = []
    content_blocks: List[Tuple[int, List[str]]] = []
    current_level: Optional[int] = None
    current_buffer: List[str] = []

    def flush():
        nonlocal current_level, current_buffer
        if current_level is not None:
            content_blocks.append((current_level, current_buffer))
        current_level = None
        current_buffer = []

    for ln in lines:
        if h2.match(ln):
            flush()
            headings.append((2, ln))
            current_level = 2
        elif h1.match(ln):
            flush()
            headings.append((1, ln))
            current_level = 1
        else:
            if current_level is None:
                current_level = 1
                headings.append((1, "Capitolo introduttivo"))
            current_buffer.append(ln)
    flush()
    return headings, content_blocks

# -------------------------------
# Costruzione DOCX
# -------------------------------
def set_page_size_and_margins(doc: Document, page: str):
    if page not in ("6x9", "8.5x11"):
        raise ValueError("Formato pagina non valido. Usa 6x9 o 8.5x11.")
    width_in, height_in = (6.0, 9.0) if page == "6x9" else (8.5, 11.0)
    section = doc.sections[0]
    section.page_width = Inches(width_in)
    section.page_height = Inches(height_in)
    margin = Cm(2.0)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin


def add_title_page(doc: Document, title: str, subtitle: str, author: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title.strip())
    r.font.name = "Calibri"
    r.font.size = Pt(24)
    r.bold = True

    doc.add_paragraph()
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(subtitle.strip())
    r2.font.name = "Calibri"
    r2.font.size = Pt(14)

    for _ in range(3):
        doc.add_paragraph()

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run(author.strip())
    r3.font.name = "Calibri"
    r3.font.size = Pt(12)
    r3.italic = True

    doc.add_page_break()


def add_copyright_page(doc: Document, author: str):
    year = datetime.datetime.now().year
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"© {year} {author}. Tutti i diritti riservati.")
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    doc.add_page_break()


def add_toc_field(doc: Document, levels: str = "1-2"):
    p = doc.add_paragraph()
    r = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f'TOC \\o "{levels}" \\h \\z \\u')
    r._r.append(fld)
    doc.add_page_break()


def add_footer_page_numbers(doc: Document):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r._r.append(fld)


def add_heading(doc: Document, text: str, level: int):
    p = doc.add_paragraph()
    try:
        p.style = doc.styles[f"Heading {level}"]
        p.add_run(text.strip())
    except Exception:
        r = p.add_run(text.strip())
        r.font.name = "Calibri"
        r.font.size = Pt(16 if level == 1 else 13)
        r.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_paragraph(doc: Document, text: str):
    p = doc.add_paragraph()
    r = p.add_run(text.strip())
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def build_docx(
    title: str,
    subtitle: str,
    author: str,
    page_format: str,
    headings: List[HeadingItem],
    content_blocks: List[Tuple[int, List[str]]],
) -> Document:
    doc = Document()
    set_page_size_and_margins(doc, page_format)
    add_footer_page_numbers(doc)
    add_title_page(doc, title, subtitle, author)
    add_copyright_page(doc, author)
    add_toc_field(doc, "1-2")

    idx_block = 0
    for level, heading_text in headings:
        add_heading(doc, heading_text, level)
        if idx_block < len(content_blocks):
            b_level, paragraphs = content_blocks[idx_block]
            if b_level == level:
                for t in paragraphs:
                    add_paragraph(doc, t)
                idx_block += 1
        if level == 1:
            doc.add_page_break()
    return doc

# -------------------------------
# PDF nativo con ReportLab
# -------------------------------
def _pagesize_for(page_format: str):
    if page_format == "6x9":
        return (6.0 * inch, 9.0 * inch)
    return (8.5 * inch, 11.0 * inch)

class TOCDoc(SimpleDocTemplate):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._toc_entries = []

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph):
            text = flowable.getPlainText()
            style_name = flowable.style.name
            if style_name == "H1":
                level = 1
            elif style_name == "H2":
                level = 2
            else:
                return
            self.notify("TOCEntry", (level, text, self.page))

def build_pdf_reportlab(
    out_pdf_path: str,
    title: str,
    subtitle: str,
    author: str,
    page_format: str,
    headings: List[HeadingItem],
    content_blocks: List[Tuple[int, List[str]]],
):
    pagesize = _pagesize_for(page_format)
    left_margin = right_margin = top_margin = bottom_margin = 2 * cm

    doc = TOCDoc(
        out_pdf_path,
        pagesize=pagesize,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
        title=title,
        author=author,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleCenter", parent=styles["Title"], alignment=1, fontName="Helvetica"))
    styles.add(ParagraphStyle(name="Subtitle", parent=styles["Heading2"], alignment=1, fontName="Helvetica"))
    styles.add(ParagraphStyle(name="Author", parent=styles["Normal"], alignment=1, fontName="Helvetica-Oblique", fontSize=11))
    styles.add(ParagraphStyle(name="Copyright", parent=styles["Normal"], alignment=1, fontName="Helvetica", fontSize=9))
    styles.add(ParagraphStyle(name="H1", parent=styles["Heading1"], fontName="Helvetica-Bold", spaceBefore=12, spaceAfter=6))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4))
    styles.add(ParagraphStyle(name="Body", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=14))

    story = []

    # Pagina titolo
    story.append(Spacer(1, 3 * cm))
    story.append(Paragraph(title, styles["TitleCenter"]))
    if subtitle.strip():
        story.append(Spacer(1, 0.5 * cm))
        story.append(Paragraph(subtitle, styles["Subtitle"]))
    story.append(Spacer(1, 2 * cm))
    story.append(Paragraph(author, styles["Author"]))
    story.append(PageBreak())

    # Copyright
    year = datetime.datetime.now().year
    story.append(Spacer(1, 10 * cm))
    story.append(Paragraph(f"© {year} {author}. Tutti i diritti riservati.", styles["Copyright"]))
    story.append(PageBreak())

    # TOC
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(name="TOCHeading1", fontName="Helvetica", fontSize=11, leftIndent=0, firstLineIndent=0, spaceBefore=2, leading=12),
        ParagraphStyle(name="TOCHeading2", fontName="Helvetica", fontSize=10, leftIndent=16, firstLineIndent=0, spaceBefore=0, leading=12),
    ]
    story.append(Paragraph("Indice", styles["H1"]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(toc)
    story.append(PageBreak())

    # Contenuti
    idx_block = 0
    for level, text in headings:
        style = styles["H1"] if level == 1 else styles["H2"]
        story.append(Paragraph(text.strip(), style))
        if idx_block < len(content_blocks):
            b_level, paragraphs = content_blocks[idx_block]
            if b_level == level:
                for t in paragraphs:
                    story.append(Paragraph(t.strip(), styles["Body"]))
                    story.append(Spacer(1, 0.15 * cm))
                idx_block += 1
        if level == 1:
            story.append(PageBreak())

    def on_page(canvas, doc_):
        canvas.saveState()
        w, h = pagesize
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(w / 2.0, 1.0 * cm, str(doc_.page))
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

# -------------------------------
# Runner CLI
# -------------------------------
def cli_main():
    ap = argparse.ArgumentParser(description="Genera DOCX e PDF formattati da DOCX o PDF in ingresso")
    ap.add_argument("--input", required=True, help="File sorgente .docx o .pdf")
    ap.add_argument("--title", required=True, help="Titolo")
    ap.add_argument("--subtitle", default="", help="Sottotitolo")
    ap.add_argument("--author", required=True, help="Autore")
    ap.add_argument("--page", choices=["6x9", "8.5x11"], default="6x9", help="Formato pagina")
    ap.add_argument("--out-docx", default="output_formattato.docx", help="Output DOCX")
    ap.add_argument("--out-pdf", default="output_formattato.pdf", help="Output PDF")
    ap.add_argument("--no-pdf", action="store_true", help="Non generare il PDF")
    args = ap.parse_args()

    src = args.input.lower()
    if src.endswith(".docx"):
        headings, blocks = parse_docx_input(args.input)
    elif src.endswith(".pdf"):
        headings, blocks = parse_pdf_input(args.input)
    else:
        raise ValueError("Formato input non supportato. Usa .docx o .pdf")

    docx_obj = build_docx(args.title, args.subtitle, args.author, args.page, headings, blocks)
    docx_obj.save(args.out_docx)
    print(f"[OK] DOCX creato: {args.out_docx}")

    if not args.no_pdf:
        build_pdf_reportlab(args.out_pdf, args.title, args.subtitle, args.author, args.page, headings, blocks)
        print(f"[OK] PDF creato: {args.out_pdf}")

# -------------------------------
# UI Streamlit integrata
# -------------------------------
def ui_main():
    import tempfile
    import streamlit as st

    st.set_page_config(page_title="Formattatore Libro", page_icon="📘", layout="centered")
    st.title("Formattatore Libro")
    st.caption("Genera DOCX e PDF da un DOCX o PDF con capitoli e sottocapitoli.")

    with st.sidebar:
        st.header("Impostazioni")
        page_format = st.selectbox("Formato pagina", ("6x9", "8.5x11"))
        title = st.text_input("Titolo", "")
        subtitle = st.text_input("Sottotitolo", "")
        author = st.text_input("Autore", "")
        gen_pdf = st.checkbox("Genera anche PDF", value=True)

    uploaded = st.file_uploader("Carica un DOCX o un PDF", type=["docx", "pdf"])
    run = st.button("Genera")

    if run:
        if not uploaded:
            st.error("Carica prima un file DOCX o PDF.")
            st.stop()
        if not title or not author:
            st.error("Compila Titolo e Autore.")
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            suffix = ".docx" if uploaded.name.lower().endswith(".docx") else ".pdf"
            in_path = os.path.join(tmpdir, "input" + suffix)
            with open(in_path, "wb") as f:
                f.write(uploaded.getbuffer())

            try:
                if in_path.endswith(".docx"):
                    headings, blocks = parse_docx_input(in_path)
                else:
                    headings, blocks = parse_pdf_input(in_path)

                # DOCX
                doc = build_docx(title, subtitle, author, page_format, headings, blocks)
                out_docx = os.path.join(tmpdir, "output_formattato.docx")
                doc.save(out_docx)
                with open(out_docx, "rb") as f:
                    docx_bytes = f.read()
                st.success("DOCX generato")
                st.download_button(
                    label="Scarica DOCX",
                    data=docx_bytes,
                    file_name=f"{title.strip().replace(' ','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                # PDF nativo
                if gen_pdf:
                    out_pdf = os.path.join(tmpdir, "output_formattato.pdf")
                    build_pdf_reportlab(out_pdf, title, subtitle, author, page_format, headings, blocks)
                    with open(out_pdf, "rb") as f:
                        pdf_bytes = f.read()
                    st.success("PDF generato")
                    st.download_button(
                        label="Scarica PDF",
                        data=pdf_bytes,
                        file_name=f"{title.strip().replace(' ','_')}.pdf",
                        mime="application/pdf",
                    )

            except Exception as e:
                st.error(f"Errore: {e}")

# -------------------------------
# Entry point
# -------------------------------
if __name__ == "__main__":
    wants_cli = any(arg.startswith("--") for arg in sys.argv[1:])
    if wants_cli:
        cli_main()
    else:
        try:
            import streamlit  # noqa
            ui_main()
        except Exception:
            print("Per la UI: streamlit run book_formatter.py")
            print("Oppure usa la CLI con: python book_formatter.py --input file.docx --title 'Titolo' --author 'Autore'")
            sys.exit(1)
