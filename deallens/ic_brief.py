"""Readable black-and-white investment committee memo, 11pt minimum font.

Each substantive narrative section uses at most five complete CIM-backed
sentences, not arbitrary character truncation. Pagination expands as needed.
"""
from io import BytesIO
import re

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT = "Times New Roman"
INK = RGBColor(0, 0, 0)
BODY_SIZE = Pt(11)
TITLE_SIZE = Pt(13)
MISSING = "Not established in the uploaded CIM. Further diligence required."

# Sentence breaks are defined at punctuation + uppercase word starts rather
# than character counts. A period after the abbreviation "p." in [CIM p. N]
# must not be treated as a sentence break.
_SENTENCE_BREAK = re.compile(r"(?<=[.!?])\s+(?=[A-Z(])")
_CITATION_BREAK = re.compile(r"(\[CIM p\.)\s+(\d+\])")


def _clean(value):
    text = str(value or MISSING)
    text = re.sub(r"(?m)^\s*(?:[-*•]\s+|\d+[.)]\s+)", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip() or MISSING


def _sentences(value, max_sentences=5):
    """Preserve complete thoughts and CIM page citations without truncation."""
    text = _clean(value)
    # Protect [CIM p. 14] from sentence segmentation, restoring it after.
    protected = _CITATION_BREAK.sub(r"\1\2", text)
    chunks = [p.strip() for p in _SENTENCE_BREAK.split(protected) if p.strip()]
    # Source models occasionally return comma/semicolon-separated bullets.
    # Preserve those intact instead of inventing a narrative.
    if not chunks:
        return MISSING
    return " ".join(chunks[:max_sentences]).replace("[CIM p.", "[CIM p. ")


def _run(p, value, bold=False, size=BODY_SIZE):
    run = p.add_run(str(value))
    run.bold = bold
    run.font.name = FONT
    run.font.size = size
    run.font.color.rgb = INK
    return run


def _paragraph(doc, after=3, before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.04
    p.paragraph_format.widow_control = True
    return p


def _heading(doc, name):
    p = _paragraph(doc, after=2, before=7)
    p.paragraph_format.keep_with_next = True
    _run(p, name, bold=True)


def _body(doc, content, max_sentences=5):
    p = _paragraph(doc, after=5)
    _run(p, _sentences(content, max_sentences))


def _label_line(doc, label, content):
    p = _paragraph(doc, after=3)
    _run(p, label + ": ", bold=True)
    _run(p, _sentences(content, 1))


def _border(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    edge = OxmlElement("w:bottom")
    edge.set(qn("w:val"), "single")
    edge.set(qn("w:sz"), "4")
    edge.set(qn("w:color"), "000000")
    borders.append(edge)


def _table_cell(cell, text, bold=False):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    _run(p, text, bold=bold)
    _border(cell)


def export_two_page_ic(company, memo, moic=None, irr=None, verified=False, issues=None):
    """Export a concise IC memo; expands beyond two pages if needed for 11pt text."""
    from .full_model import recommendation

    status, rationale = recommendation(moic, irr, verified, issues)
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin, section.bottom_margin = Inches(0.68), Inches(0.61)
    section.left_margin, section.right_margin = Inches(0.75), Inches(0.75)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = BODY_SIZE
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(3)
    normal.paragraph_format.line_spacing = 1.04

    p = _paragraph(doc, after=3)
    _run(p, _clean(company), bold=True, size=TITLE_SIZE)
    p = _paragraph(doc, after=6)
    _run(p, "Investment Committee Memorandum  |  ", bold=True)
    _run(p, f"Recommendation: {status}")

    summary = doc.add_table(rows=2, cols=3)
    summary.autofit = True
    for i, title in enumerate(("Five-year sponsor MOIC", "Five-year sponsor IRR", "Investment hurdle")):
        _table_cell(summary.cell(0, i), title, bold=True)
    for i, value in enumerate((
        f"{moic:.2f}x" if moic is not None else "Not verified",
        f"{irr:.1%}" if irr is not None else "Not verified",
        "25% IRR / ~3.05x MOIC",
    )):
        _table_cell(summary.cell(1, i), value)

    _heading(doc, "Business Overview")
    for label, key in (
        ("Management and Employees", "management_employees"),
        ("Products and Services", "products_services"),
        ("Operations", "general_operations"),
        ("Customers", "customer_base"),
    ):
        _label_line(doc, label, memo.get(key))

    for heading, key in (
        ("Historical Financial Performance", "historical_financial_performance"),
        ("Market and Competition", "market_overview"),
        ("Growth Opportunities", "growth_opportunities"),
        ("Investment Merits, Risks and Mitigants", "investment_merits_risks_mitigants"),
        ("Investment Thesis", "investment_thesis"),
        ("Potential Exits and Buyer Universe", "potential_exit_strategies_buyer_universe"),
    ):
        _heading(doc, heading)
        _body(doc, memo.get(key), 5)

    _heading(doc, "Recommendation")
    # This decision comes from verified sponsor returns, not the seller's
    # recommendation or an LLM-generated enthusiasm about the business.
    _body(doc, f"{status}. {rationale}", 5)

    _heading(doc, "Outstanding Diligence")
    if issues:
        _body(doc, "; ".join(str(issue) for issue in issues[:5]) + ". "
              "All unresolved items require independent verification before an investment decision.", 5)
    else:
        _body(
            doc,
            "Verify reported financials, adjusted EBITDA addbacks, customer concentration, "
            "maintenance capital expenditure, net working capital, and financing terms. "
            "Reconcile cash flow and the equity waterfall before relying on the five-year return estimate.",
            5,
        )

    output = BytesIO()
    doc.save(output)
    return output.getvalue()
