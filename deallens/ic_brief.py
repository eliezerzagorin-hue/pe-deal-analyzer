"""Two-page, black-and-white IC memorandum with 11pt minimum type."""
from io import BytesIO
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT = "Times New Roman"
INK = RGBColor(0, 0, 0)
BODY_SIZE = Pt(11)
TITLE_SIZE = Pt(13)


def _tidy(value, limit=420):
    text = re.sub(r"\s+", " ", str(value or "Not disclosed in the CIM; verify in diligence.")).strip()
    if len(text) <= limit:
        return text
    cut = text[:limit].rfind(". ")
    if cut < int(limit * 0.60):
        cut = limit
    return text[:cut].rstrip(" ,;") + ("" if text[:cut].endswith(".") else ".")


def _run(p, value, *, bold=False, size=BODY_SIZE):
    run = p.add_run(str(value))
    run.bold = bold
    run.font.name = FONT
    run.font.size = size
    run.font.color.rgb = INK
    return run


def _paragraph(doc, value="", *, after=2, before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0
    if value:
        _run(p, value)
    return p


def _heading(doc, name):
    p = _paragraph(doc, after=2, before=6)
    p.paragraph_format.keep_with_next = True
    _run(p, name, bold=True)


def _body(doc, text, limit=430):
    p = _paragraph(doc, after=3)
    _run(p, _tidy(text, limit))


def _label_line(doc, label, text, limit=150):
    p = _paragraph(doc, after=2)
    _run(p, label + ": ", bold=True)
    _run(p, _tidy(text, limit))


def _border(cell, position="bottom"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    edge = OxmlElement(f"w:{position}")
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
    """Compact, strictly monochrome two-page investment committee document."""
    from .full_model import recommendation
    status, rationale = recommendation(moic, irr, verified, issues)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.63)
    section.bottom_margin = Inches(0.58)
    section.left_margin = Inches(0.74)
    section.right_margin = Inches(0.74)
    section.header_distance = Inches(0.25)
    section.footer_distance = Inches(0.25)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = BODY_SIZE
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.0

    p = _paragraph(doc, after=3)
    _run(p, _tidy(company, 100), bold=True, size=TITLE_SIZE)
    p = _paragraph(doc, after=5)
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
    for label, field in (
        ("Management and Employees", "management_employees"),
        ("Products and Services", "products_services"),
        ("Operations", "general_operations"),
        ("Customers", "customer_base"),
    ):
        _label_line(doc, label, memo.get(field), limit=165)

    _heading(doc, "Historical Financial Performance")
    _body(doc, memo.get("historical_financial_performance"), limit=400)

    _heading(doc, "Market and Competition")
    _body(doc, memo.get("market_overview"), limit=350)

    _heading(doc, "Growth Opportunities")
    _body(doc, memo.get("growth_opportunities"), limit=390)

    doc.add_page_break()

    _heading(doc, "Investment Merits, Risks and Mitigants")
    _body(doc, memo.get("investment_merits_risks_mitigants"), limit=650)

    _heading(doc, "Investment Thesis")
    _body(doc, memo.get("investment_thesis"), limit=490)

    _heading(doc, "Potential Exits and Buyer Universe")
    _body(doc, memo.get("potential_exit_strategies_buyer_universe"), limit=450)

    _heading(doc, "Recommendation")
    _body(doc, status + ". " + rationale, limit=365)

    _heading(doc, "Outstanding Diligence")
    if issues:
        _body(doc, "; ".join(str(s) for s in issues[:4]), limit=485)
    else:
        _body(doc,
              "Verify CIM figures, EBITDA adjustments, customer concentration, "
              "maintenance capital expenditure, debt terms, working capital, "
              "exit assumptions, and financial statement reconciliations.", limit=430)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()
