import argparse
import datetime
import os
import re
import shutil
import subprocess
from typing import List, Tuple, Optional

# --- DOCX ---
from docx import Document
from docx.shared import Pt, Inches, Cm  # Cm è IMPORTANTE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- PDF (input opzionale) ---
try:
    from pdfminer.high_level import extract_text as pdf_extract_text
except Exception:
    pdf_extract_text = None


# =========================
# Utility: parsing input
# =========================
HeadingItem = Tuple[int, str]  # (level, text)

def parse_docx_input(path: str) -> Tuple[List[HeadingItem], List[Tuple[int, List[str]]]]:
    """
    Ritorna:
      - headings: [(level, text)]
      - content_blocks: [(level, [paragraphs...])] in cui level 1=capitolo, 2=sottocapitolo
    Riconosce: stili 'Heading 1'/'Heading 2' o titoli numerati (es: '1.', '1.1').
    """
    doc = Document(path)
    headings: List[HeadingItem] = []
    content_blocks: List[Tuple[int, List[str]]] = []

    current_level = None
    current_buffer: List[str] = []

    def flush():
        nonlocal current_level, current_buffer, content_blocks
        if current_level is not None:
            content_blocks.append((current_level, current_buffer))
        current_level = None
        current_buffer = []

    h1_pattern = re.compile(r"^\s*(\d+)\.\s+.+")          # 1. Titolo
    h2_pattern = re.compile(r"^\s*(\d+)\.(\d+)\s+.+")      # 1.1 Sottotitolo

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
            # fallback su pattern numerati
            if h2_pattern.match(text):
                level_detected = 2
            elif h1_pattern.match(text):
                level_detected = 1

        if level_detected:
            # chiudi blocco precedente
            flush()
            headings.append((level_detected, text))
            current_level = level_detected
        else:
            if current_level is None:
                # se inizia con testo senza heading, assegnalo al capitolo 1 "Implicito"
                current_level = 1
                headings.append((1, "Capitolo introduttivo"))
            current_buffer.append(text)

    flush()
    return headings, content_blocks


def parse_pdf_input(path: str) -> Tuple[List[HeadingItem], List[Tuple[int, List[str]]]]:
    """
    Parsing basilare da PDF: serve testo strutturato (capitoli numerati tipo '1.' e sottocapitoli '1.1').
    """
    if pdf_extract_text is None:
        raise RuntimeError("pdfminer.six non è installato. Esegui: pip install pdfminer.six")

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
        m2 = h2_pattern.match(ln)
        m1 = h1_pattern.match(ln)
        if m2:
            flush()
            title = ln
            headings.append((2, title))
            current_level = 2
        elif m1:
            flush()
            title = ln
            headings.append((1, title))
            current_level = 1
        else:
            if current_level is None:
                current_level = 1
                headings.append((1, "Capitolo introduttivo"))
            current_buffer.append(ln)

    flush()
    return headings, content_blocks


# =========================
# DOCX building
# =========================
def set_page_size_and_margins(doc: Document, page: str):
    """
    page: '6x9' o '8.5x11'
    Margini conservativi: 2.0 cm.
    """
    if page not in ("6x9", "8.5x11"):
        raise ValueError("Formato pagina non valido. Usa '6x9' oppure '8.5x11'.")

    if page == "6x9":
        width_in, height_in = 6.0, 9.0
    else:
        width_in, height_in = 8.5, 11.0

    section = doc.sections[0]
    section.page_width = Inches(width_in)
    section.page_height = Inches(height_in)
    # margini
    margin = Cm(2.0)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin


def add_title_page(doc: Document, title: str, subtitle: str, author: str):
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run(title.strip())
    run.font.name = "Calibri"
    run.font.size = Pt(24)
    run.bold = True

    doc.add_paragraph()  # spazio
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run(subtitle.strip())
    run.font.name = "Calibri"
    run.font.size = Pt(14)

    for _ in range(3):
        doc.add_paragraph()

    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_auth.add_run(author.strip())
    run.font.name = "Calibri"
    run.font.size = Pt(12)
    run.italic = True

    doc.add_page_break()


def add_copyright_page(doc: Document, author: str):
    year = datetime.datetime.now().year
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"© {year} {author}. Tutti i diritti riservati.")
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    doc.add_page_break()


def add_toc(doc: Document, levels: str = "1-2"):
    """
    Inserisce un campo TOC aggiornabile in Word:
    L'utente (o automazione) dovrà aggiornare i campi prima dell'esportazione definitiva.
    """
    p = doc.add_paragraph()
    r = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f'TOC \\o "{levels}" \\h \\z \\u')
    r._r.append(fld)
    doc.add_page_break()


def add_footer_page_numbers(doc: Document):
    """
    Numeri di pagina centrati nel piè di pagina: {PAGE}
    """
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Inserisci campo PAGE
    r = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r._r.append(fld)


def add_heading(doc: Document, text: str, level: int):
    p = doc.add_paragraph()
    p.style = doc.styles.get(f"Heading {level}", None)
    if p.style is None:
        # fallback manuale
        run = p.add_run(text.strip())
        run.font.name = "Calibri"
        run.font.size = Pt(16 if level == 1 else 13)
        run.bold = True
    else:
        p.add_run(text.strip())
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_paragraph(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text.strip())
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    # Giustificato
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

    # Front matter
    add_title_page(doc, title, subtitle, author)
    add_copyright_page(doc, author)
    add_toc(doc, "1-2")

    # Contenuti: per semplicità, accoppiamo headings e blocchi nella sequenza in cui sono arrivati.
    # content_blocks è una lista di blocchi per livello (1/2). Inseriamo interruzione pagina alla FINE dei level 1.
    idx_block = 0
    for level, heading_text in headings:
        add_heading(doc, heading_text, level)

        # Se esiste un blocco testo a questo punto e di livello compatibile, inseriscilo
        if idx_block < len(content_blocks):
            b_level, paragraphs = content_blocks[idx_block]
            if b_level == level:
                for t in paragraphs:
                    add_paragraph(doc, t)
                idx_block += 1

        # Interruzione pagina a fine capitolo (level 1)
        if level == 1:
            doc.add_page_break()

    return doc


# =========================
# PDF export
# =========================
def export_pdf(docx_path: str, pdf_path: str) -> bool:
    """
    Prova 1: LibreOffice headless ('soffice')
    Prova 2: docx2pdf (solo Windows con Word)
    Ritorna True se PDF creato.
    """
    # 1) LibreOffice
    soffice = shutil.which("soffice") or shutil.which("soffice.bin")
    if soffice:
        try:
            outdir = os.path.dirname(os.path.abspath(pdf_path)) or "."
            subprocess.check_call([
                soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path
            ])
            # LibreOffice crea <same-name>.pdf
            base = os.path.splitext(os.path.basename(docx_path))[0]
            created = os.path.join(outdir, base + ".pdf")
            if created != pdf_path:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                os.replace(created, pdf_path)
            return True
        except Exception:
            pass

    # 2) docx2pdf (Windows/MS Word)
    try:
        import platform
        if platform.system().lower().startswith("win"):
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            return True
    except Exception:
        pass

    return False


# =========================
# Main CLI
# =========================
def main():
    ap = argparse.ArgumentParser(description="Genera DOCX/PDF formattato da DOCX/PDF d’ingresso.")
    ap.add_argument("--input", required=True, help="Percorso file sorgente (.docx o .pdf)")
    ap.add_argument("--title", required=True, help="Titolo per la pagina iniziale")
    ap.add_argument("--subtitle", default="", help="Sottotitolo per la pagina iniziale")
    ap.add_argument("--author", required=True, help="Autore (per pagina titolo e copyright)")
    ap.add_argument("--page", choices=["6x9", "8.5x11"], default="6x9", help="Formato pagina")
    ap.add_argument("--out-docx", default="output_formattato.docx", help="Percorso DOCX di output")
    ap.add_argument("--out-pdf", default="output_formattato.pdf", help="Percorso PDF di output")
    ap.add_argument("--no-pdf", action="store_true", help="Non generare il PDF")
    args = ap.parse_args()

    src = args.input.lower()
    if src.endswith(".docx"):
        headings, blocks = parse_docx_input(args.input)
    elif src.endswith(".pdf"):
        headings, blocks = parse_pdf_input(args.input)
    else:
        raise ValueError("Formato input non supportato. Usa .docx o .pdf")

    # Costruisci DOCX
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

    # PDF
    if not args.no_pdf:
        ok = export_pdf(args.out_docx, args.out_pdf)
        if ok:
            print(f"[OK] PDF creato: {args.out_pdf}")
        else:
            print("[AVVISO] Conversione PDF non riuscita automaticamente. "
                  "Installa LibreOffice (soffice in PATH) oppure usa docx2pdf su Windows. "
                  "In alternativa, apri il DOCX e salva come PDF manualmente.")


if __name__ == "__main__":
    main()
