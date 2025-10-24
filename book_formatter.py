# book_formatter.py
import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys
from typing import List, Tuple, Optional

# DOCX
from docx import Document
from docx.shared import Pt, Inches, Cm  # Cm necessario
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# PDF input opzionale
try:
    from pdfminer.high_level import extract_text as pdf_extract_text
except Exception:
    pdf_extract_text = None

# Tipi
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
        nonlocal current_level, current_buffer, content_blocks
        if current_level is not None:
            content_blocks.append((current_level, current_buffer))
        current_level = None
        current_buffer = []

    h1_pattern = re.compile(r"^\s*(\d+)\.\s+.+")       # 1. Titolo
    h2_pattern = re.compile(r"^\s*(\d+)\.(\d+)\s+.+")  # 1.1 Sottotitolo

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

    h1_pattern = re.compile(r"^\s*(\d+)\.\s+(.+)")
    h2_pattern = re.compile(r"^\s*(\d+)\.(\d+)\s+(.+)")

    headings: List[HeadingItem] = []
    content_blocks: List[Tuple[int, List[str]]] = []
    current_level: Optional[int] = None
    current_buffer: List[str] = []

    def flush():
        nonlocal current_level, current_buffer, content_blocks
        if current_level is not None:
            content_blocks.append((current_level, current_buffer))
        current_level = None
        current_buffer = []

    for ln in lines:
        if h2_pattern.match(ln):
            flush()
            headings.append((2, ln))
            current_level = 2
        elif h1_pattern.match(ln):
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
        raise ValueError("Formato pagina non valido. Usa '6x9' o '8.5x11'.")

    if page == "6x9":
        width_in, height_in = 6.0, 9.0
    else:
        width_in, height_in = 8.5, 11.0

    section = doc.sections[0]
    section.page_width = Inches(width_in)
    section.page_height = Inches(height_in)
    margin = Cm(2.0)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin


def add_title_page(doc: Document, title: str, subtitle: str, author: str):
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run(title.strip())
    r.font.name = "Calibri"
    r.font.size = Pt(24)
    r.bold = True

    doc.add_paragraph()
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_sub.add_run(subtitle.strip())
    r.font.name = "Calibri"
    r.font.size = Pt(14)

    for _ in range(3):
        doc.add_paragraph()

    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_auth.add_run(author.strip())
    r.font.name = "Calibri"
    r.font.size = Pt(12)
    r.italic = True

    doc.add_page_break()


def add_copyright_page(doc: Document, author: str):
    year = datetime.datetime.now().year
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"© {year} {author}. Tutti i diritti riservati.")
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    doc.add_page_break()


def add_toc(doc: Document, levels: str = "1-2"):
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
    # tenta lo stile di Word se esiste, altrimenti fallback
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


def build_document(
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
    add_toc(doc, "1-2")

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
# Export PDF
# -------------------------------
def export_pdf(docx_path: str, pdf_path: str) -> bool:
    # LibreOffice
    soffice = shutil.which("soffice") or shutil.which("soffice.bin")
    if soffice:
        try:
            outdir = os.path.dirname(os.path.abspath(pdf_path)) or "."
            subprocess.check_call([soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path])
            base = os.path.splitext(os.path.basename(docx_path))[0]
            created = os.path.join(outdir, base + ".pdf")
            if created != pdf_path:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                os.replace(created, pdf_path)
            return True
        except Exception:
            pass
    # docx2pdf su Windows
    try:
        import platform
        if platform.system().lower().startswith("win"):
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            return True
    except Exception:
        pass
    return False

# -------------------------------
# Modalità CLI
# -------------------------------
def cli_main():
    ap = argparse.ArgumentParser(description="Genera DOCX/PDF formattato da DOCX/PDF di ingresso")
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

    doc = build_document(
        title=args.title,
        subtitle=args.subtitle,
        author=args.author,
        page_format=args.page,
        headings=headings,
        content_blocks=blocks,
    )
    doc.save(args.out_docx)
    print(f"[OK] DOCX creato: {args.out_docx}")

    if not args.no_pdf:
        ok = export_pdf(args.out_docx, args.out_pdf)
        if ok:
            print(f"[OK] PDF creato: {args.out_pdf}")
        else:
            print("Conversione PDF non riuscita. Installa LibreOffice o usa docx2pdf su Windows.")

# -------------------------------
# Modalità UI Streamlit integrata
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
                with st.spinner("Analisi del sorgente"):
                    if in_path.endswith(".docx"):
                        headings, blocks = parse_docx_input(in_path)
                    else:
                        headings, blocks = parse_pdf_input(in_path)
                with st.spinner("Costruzione DOCX"):
                    doc = build_document(
                        title=title,
                        subtitle=subtitle,
                        author=author,
                        page_format=page_format,
                        headings=headings,
                        content_blocks=blocks,
                    )
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
                if gen_pdf:
                    with st.spinner("Conversione in PDF"):
                        out_pdf = os.path.join(tmpdir, "output_formattato.pdf")
                        ok = export_pdf(out_docx, out_pdf)
                        if ok and os.path.exists(out_pdf):
                            with open(out_pdf, "rb") as f:
                                pdf_bytes = f.read()
                            st.success("PDF generato")
                            st.download_button(
                                label="Scarica PDF",
                                data=pdf_bytes,
                                file_name=f"{title.strip().replace(' ','_')}.pdf",
                                mime="application/pdf",
                            )
                        else:
                            st.warning("PDF non generato automaticamente. Installa LibreOffice in PATH o docx2pdf su Windows.")
            except Exception as e:
                st.error(f"Errore: {e}")

# -------------------------------
# Entry point
# -------------------------------
if __name__ == "__main__":
    # Se lo si lancia con streamlit run, sys.argv non contiene i parametri richiesti
    # Regola: se sono presenti argomenti espliciti per CLI usa CLI, altrimenti prova UI
    wants_cli = any(arg.startswith("--") for arg in sys.argv[1:])
    if wants_cli:
        cli_main()
    else:
        # prova UI se Streamlit è disponibile
        try:
            import streamlit  # noqa
            ui_main()
        except Exception:
            # fallback CLI con help
            print("Per la UI: streamlit run book_formatter.py")
            print("Oppure usa la CLI con argomenti. Esempio:")
            print("python book_formatter.py --input sorgente.docx --title 'Titolo' --author 'Autore' --page 6x9")
            sys.exit(1)
