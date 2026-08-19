"""
Builds a .docx for a tender-required document we don't have a hand-authored
template for, by reproducing the tender's own text for that document
verbatim -- under Airox's letterhead/footer -- instead of asking an LLM to
invent content. Many bundled tender documents (Inspection & Test Plans,
Quality Assurance Plan forms, an Approved Makes list) are already complete,
ready-to-sign documents inside the tender PDF; the bidder's job is to
countersign them, not rewrite them. This keeps that guarantee: nothing in
the body text comes from anywhere but the tender itself.
"""
import re

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX

from . import branding
from .branding import docx_safe as _docx_safe

FONT_NAME = "Times New Roman"
FONT_SIZE = 11


def slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")
    return (slug or "document") + ".docx"


def _render_table(doc, table_data: list[list[str]]) -> None:
    """Render a pdfplumber-extracted table as a real Word table. Reproducing
    a tender page with side-by-side columns as flat paragraph text scrambles
    it into unreadable word-soup (no column boundaries survive); an actual
    table keeps each cell where it belongs."""
    if not table_data or not table_data[0]:
        return
    n_cols = max(len(row) for row in table_data)
    table = doc.add_table(rows=0, cols=n_cols)
    table.style = "Table Grid"
    for row in table_data:
        cells = table.add_row().cells
        for i in range(n_cols):
            value = row[i] if i < len(row) else ""
            cells[i].text = _docx_safe(value.strip())
    doc.add_paragraph()


def build_verbatim_document(
    title: str,
    page_entries: list,
    output_path: str,
    company: dict,
    signing_date: str,
    signing_place: str,
) -> None:
    """page_entries: list of pages this document spans in the source tender
    PDF, in order. Each entry is either a plain (page_number, text) tuple
    (legacy form, always rendered as flat paragraphs) or a dict
    {"page": n, "text": str, "tables": list[list[list[str]]]} as returned by
    extraction.parse_pdf_page_range_detailed -- pages with genuine tables are
    rendered as real Word tables instead of flattened paragraph text."""
    doc = Document()
    branding.set_default_font(doc, FONT_NAME, FONT_SIZE)
    branding.add_letterhead(doc)
    branding.add_footer(doc, company)
    branding.add_title(doc, _docx_safe(title))

    note = doc.add_paragraph()
    note_run = note.add_run(
        "Reproduced verbatim from the tender document (no template was available "
        "for this item -- review against the source before signing)."
    )
    note_run.italic = True
    note_run.font.size = Pt(9)
    note_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.add_paragraph()

    for entry in page_entries:
        if isinstance(entry, dict):
            page_no, text, tables = entry["page"], entry["text"], entry.get("tables") or []
        else:
            page_no, text, tables = entry[0], entry[1], []

        caption = doc.add_paragraph()
        caption_run = caption.add_run(f"[Tender page {page_no}]")
        caption_run.italic = True
        caption_run.font.size = Pt(8)
        caption_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        # A page whose only detected "table" is a single-row box is really
        # just the page's title banner caught by pdfplumber's border
        # detection, not an actual data table -- narrative/clause pages
        # (e.g. numbered SCOPE/METHODOLOGY paragraphs) commonly get exactly
        # that. Treating it as "this page is a table page" would skip the
        # flat-text branch below and silently drop the entire page's real
        # content, leaving only the banner. A genuine checklist/ITP table
        # always has more than one row, so that's the signal used here.
        has_real_table = any(len(t) > 1 for t in tables)
        if has_real_table:
            # The page's tables already carry essentially all of its
            # meaningful content (checklist rows, form fields, headers);
            # also flattening the same text as paragraphs would duplicate it
            # and reintroduce the scrambled-column problem this exists to
            # avoid, so skip the flat-text rendering for these pages --
            # unless the table is nowhere near the whole page (see below).
            for table_data in tables:
                _render_table(doc, table_data)

        # Some pages mix a small real table with substantial narrative
        # clauses that live entirely outside it (a numbered-methodology
        # page with a short "facility list" table dropped in the middle,
        # say) -- there the table is nowhere near the whole page, and
        # skipping flat text would silently drop that narrative. Measured
        # on real MRPL ITP/procedure pages: pages where the table
        # genuinely *is* the page have extract_text() length within
        # ~1.2-1.6x the table's own character count (extract_text() just
        # re-flattens the same cells); pages with real narrative alongside
        # a small table run 4.5-7x -- so a wide-margin ratio of 2.5 tells
        # the two apart without re-triggering on table-only pages.
        table_chars = sum(len(cell or "") for t in tables for row in t for cell in row)
        needs_flat_text = not has_real_table or len(text) > table_chars * 2.5
        if needs_flat_text:
            for line in text.splitlines():
                line = _docx_safe(line.strip())
                if not line:
                    continue
                doc.add_paragraph(line)
        doc.add_paragraph()

    branding.add_signoff_tail(doc, company, signing_place, signing_date)
    branding.ensure_unique_drawing_ids(doc)
    doc.save(output_path)
