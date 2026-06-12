"""
Codesino Agreement Generator — Python / python-docx implementation
===================================================================
File: agreement/generator.py

Direct Python port of generateAgreement.js using python-docx.
Call generate_agreement_buffer(data: dict) → bytes (the .docx buffer).
"""

import os
import random
from datetime import datetime
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches, Cm, Twips
from docx.enum.style import WD_STYLE_TYPE

# ─────────────────────────────────────────────
# BRAND COLORS (as (R, G, B) tuples)
# ─────────────────────────────────────────────
BRAND_PURPLE  = RGBColor(0x5B, 0x2D, 0x8E)
MID_PURPLE    = RGBColor(0x7B, 0x4B, 0xBF)
LIGHT_PURPLE  = RGBColor(0xED, 0xE7, 0xF6)
LAVENDER      = RGBColor(0xD1, 0xC4, 0xE9)
DARK_GRAY     = RGBColor(0x33, 0x33, 0x33)
WHITE_COLOR   = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_ACCENT  = RGBColor(0x27, 0xAE, 0x60)
LIGHT_GRAY_C  = RGBColor(0xF8, 0xF7, 0xFC)
MID_GRAY_C    = RGBColor(0xCC, 0xCC, 0xCC)
MUTED_PURPLE  = RGBColor(0x7B, 0x68, 0xAA)
PLACEHOLDER   = RGBColor(0xAA, 0xAA, 0xAA)

# Hex strings for XML shading
HEX_BRAND_PURPLE  = "5B2D8E"
HEX_MID_PURPLE    = "7B4BBF"
HEX_LIGHT_PURPLE  = "EDE7F6"
HEX_LAVENDER      = "D1C4E9"
HEX_DARK_GRAY     = "333333"
HEX_WHITE         = "FFFFFF"
HEX_LIGHT_GRAY    = "F8F7FC"
HEX_GREEN         = "27AE60"
HEX_C5B8E0        = "C5B8E0"


# ─────────────────────────────────────────────
# XML HELPERS
# ─────────────────────────────────────────────

def hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_cell_shading(cell, fill_hex):
    """Set table cell background shading via XML."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.upper())
    # Remove existing shading
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    tcPr.append(shd)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """Set individual border sides on a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tbl_borders = OxmlElement("w:tcBorders")
    for side, cfg in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        if cfg is None:
            el.set(qn("w:val"), "none")
        else:
            el.set(qn("w:val"), cfg.get("val", "single"))
            el.set(qn("w:sz"), str(cfg.get("sz", 4)))
            el.set(qn("w:color"), cfg.get("color", "000000"))
        tbl_borders.append(el)
    # Remove existing borders element
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    tcPr.append(tbl_borders)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Set cell internal padding in twentieths of a point (twips)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    for existing in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(existing)
    tcPr.append(mar)


def set_cell_vertical_align(cell, align="center"):
    """Set vertical alignment: top / center / bottom."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), align)
    for existing in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(existing)
    tcPr.append(vAlign)


def set_cell_width(cell, width_dxa):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    for ex in tcPr.findall(qn("w:tcW")):
        tcPr.remove(ex)
    tcPr.append(tcW)


def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    cs = OxmlElement("w:cantSplit")
    trPr.append(cs)


def add_paragraph_border_bottom(para, color=HEX_C5B8E0, sz=8, space=4):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    for ex in pPr.findall(qn("w:pBdr")):
        pPr.remove(ex)
    pPr.append(pBdr)


def add_paragraph_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    if before:
        spacing.set(qn("w:before"), str(before))
    if after:
        spacing.set(qn("w:after"), str(after))
    for ex in pPr.findall(qn("w:spacing")):
        pPr.remove(ex)
    pPr.append(spacing)


def set_run_font(run, font_name="Arial"):
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    for ex in rPr.findall(qn("w:rFonts")):
        rPr.remove(ex)
    rPr.insert(0, rFonts)


def set_table_width(table, width_dxa):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(width_dxa))
    tblW.set(qn("w:type"), "dxa")
    for ex in tblPr.findall(qn("w:tblW")):
        tblPr.remove(ex)
    tblPr.append(tblW)


def set_column_widths(table, widths):
    """Set column widths in DXA using tblGrid."""
    tbl = table._tbl
    for ex in tbl.findall(qn("w:tblGrid")):
        tbl.remove(ex)
    tblGrid = OxmlElement("w:tblGrid")
    for w in widths:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        tblGrid.append(gridCol)
    # Insert after tblPr
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        tblPr.addnext(tblGrid)
    else:
        tbl.insert(0, tblGrid)


def set_table_no_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    for ex in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(ex)
    tblPr.append(tblBorders)


def no_space_after(para):
    add_paragraph_spacing(para, before=0, after=0)


# ─────────────────────────────────────────────
# MISC HELPERS
# ─────────────────────────────────────────────

def generate_ref():
    now = datetime.now()
    y = str(now.year)[-2:]
    m = str(now.month).zfill(2)
    d = str(now.day).zfill(2)
    rand = random.randint(1000, 9999)
    return f"CD-AGR-{y}{m}{d}-{rand}"


def format_date(date_str=None):
    if not date_str:
        return datetime.now().strftime("%d %B %Y")
    try:
        dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
        return dt.strftime("%d %B %Y")
    except Exception:
        return str(date_str)


def format_datetime():
    return datetime.now().strftime("%d %B %Y, %H:%M")


def cb(checked):
    return "☑" if checked else "☐"


def val_text(v):
    return v if v else "Not Provided"


# ─────────────────────────────────────────────
# DOCUMENT BUILDING HELPERS
# ─────────────────────────────────────────────

def add_run(para, text, bold=False, italic=False, size_pt=11, color=None, font="Arial"):
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    run.font.name = font
    if color:
        run.font.color.rgb = color if isinstance(color, RGBColor) else hex_to_rgb(color)
    set_run_font(run, font)
    return run


def heading1(doc, text):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=280, after=100)
    add_paragraph_border_bottom(para, color=HEX_BRAND_PURPLE, sz=8, space=4)
    add_run(para, text, bold=True, size_pt=14, color=BRAND_PURPLE)
    return para


def heading2(doc, text):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=180, after=60)
    add_run(para, text, bold=True, size_pt=12, color=MID_PURPLE)
    return para


def body_text(doc, text, italic=False):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=40, after=40)
    add_run(para, text, italic=italic, size_pt=11, color=DARK_GRAY)
    return para


def italic_note(doc, text):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=40, after=80)
    add_run(para, text, italic=True, size_pt=10, color=MUTED_PURPLE)
    return para


def spacer(doc, before=80, after=80):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=before, after=after)
    return para


def divider(doc):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=160, after=160)
    add_paragraph_border_bottom(para, color=HEX_C5B8E0, sz=2, space=1)
    return para


def label_value_para(doc, label, value, label_color=None, indent_cell=None):
    """Adds a label: value paragraph to doc or into a table cell container."""
    if label_color is None:
        label_color = BRAND_PURPLE
    container = indent_cell if indent_cell is not None else doc
    para = container.add_paragraph()
    add_paragraph_spacing(para, before=0, after=60)
    r1 = para.add_run(f"{label}: ")
    r1.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = label_color
    set_run_font(r1)
    is_np = not value or value == "Not Provided"
    r2 = para.add_run("Not Provided" if is_np else value)
    r2.font.size = Pt(11)
    r2.italic = is_np
    r2.font.color.rgb = PLACEHOLDER if is_np else DARK_GRAY
    set_run_font(r2)
    return para


def section_box_table(doc, paragraphs_fn, fill_color=HEX_WHITE):
    """
    Creates a 1-column full-width table with purple left border.
    paragraphs_fn(cell) should add paragraphs into the cell.
    """
    table = doc.add_table(rows=1, cols=1)
    set_table_width(table, 9360)
    set_column_widths(table, [9360])
    set_table_no_borders(table)

    cell = table.cell(0, 0)
    set_cell_width(cell, 9360)
    set_cell_shading(cell, fill_color)
    set_cell_margins(cell, top=120, bottom=120, left=200, right=120)
    set_cell_borders(
        cell,
        top={"val": "single", "sz": 2, "color": HEX_C5B8E0},
        bottom={"val": "single", "sz": 2, "color": HEX_C5B8E0},
        left={"val": "single", "sz": 12, "color": HEX_BRAND_PURPLE},
        right=None,
    )
    # Remove default empty paragraph
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)

    paragraphs_fn(cell)
    return table


def add_lv_to_cell(cell, label, value, label_color=None):
    label_value_para(None, label, value, label_color=label_color, indent_cell=cell)


def data_table(doc, headers, rows, col_widths):
    """Creates a styled data table with purple header row."""
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    set_table_width(table, sum(col_widths))
    set_column_widths(table, col_widths)
    set_table_no_borders(table)

    purple_border = {"val": "single", "sz": 2, "color": HEX_C5B8E0}

    # Header row
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_width(cell, col_widths[i])
        set_cell_shading(cell, HEX_BRAND_PURPLE)
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        set_cell_borders(cell, top=purple_border, bottom=purple_border, left=purple_border, right=purple_border)
        for p in list(cell.paragraphs):
            p._element.getparent().remove(p._element)
        para = cell.add_paragraph()
        add_run(para, h, bold=True, size_pt=10, color=WHITE_COLOR)

    # Data rows
    for ri, row_data in enumerate(rows):
        row = table.rows[ri + 1]
        fill = HEX_WHITE if ri % 2 == 0 else HEX_LIGHT_PURPLE
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            set_cell_width(cell, col_widths[ci])
            set_cell_shading(cell, fill)
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            set_cell_borders(cell, top=purple_border, bottom=purple_border, left=purple_border, right=purple_border)
            for p in list(cell.paragraphs):
                p._element.getparent().remove(p._element)
            para = cell.add_paragraph()
            add_run(para, str(val) if val is not None else "", size_pt=10, color=DARK_GRAY)

    return table


def add_bullet(cell_or_doc, text):
    para = cell_or_doc.add_paragraph(style="List Bullet")
    add_paragraph_spacing(para, before=40, after=40)
    # Remove default run and add styled one
    for r in list(para.runs):
        r._r.getparent().remove(r._r)
    add_run(para, text, size_pt=11, color=DARK_GRAY)
    return para


def add_numbered(doc_or_cell, text):
    para = doc_or_cell.add_paragraph(style="List Number")
    add_paragraph_spacing(para, before=40, after=40)
    for r in list(para.runs):
        r._r.getparent().remove(r._r)
    add_run(para, text, size_pt=11, color=DARK_GRAY)
    return para


def sig_lines(cell):
    """Add 3 blank lines + underline for signature."""
    for _ in range(2):
        p = cell.add_paragraph()
        no_space_after(p)
    p = cell.add_paragraph()
    add_paragraph_spacing(p, before=0, after=60)
    add_paragraph_border_bottom(p, color=HEX_BRAND_PURPLE, sz=4, space=1)


# ─────────────────────────────────────────────
# MAIN GENERATOR
# ─────────────────────────────────────────────

def generate_agreement_buffer(data: dict) -> bytes:
    doc = Document()

    # ── Page setup ──
    for section in doc.sections:
        section.page_width = Twips(12240)
        section.page_height = Twips(15840)
        section.left_margin = Twips(1440)
        section.right_margin = Twips(1440)
        section.top_margin = Twips(1080)
        section.bottom_margin = Twips(1080)

    # ── Ensure List styles exist ──
    try:
        doc.styles["List Bullet"]
    except KeyError:
        doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH)
    try:
        doc.styles["List Number"]
    except KeyError:
        doc.styles.add_style("List Number", WD_STYLE_TYPE.PARAGRAPH)

    agreement_ref = generate_ref()
    agreement_date = format_datetime()

    # ═══════════════════════════════════════════
    # COVER HEADER
    # ═══════════════════════════════════════════
    header_table = doc.add_table(rows=1, cols=2)
    set_table_width(header_table, 9360)
    set_column_widths(header_table, [2000, 7360])
    set_table_no_borders(header_table)

    left_cell = header_table.cell(0, 0)
    right_cell = header_table.cell(0, 1)

    set_cell_width(left_cell, 2000)
    set_cell_shading(left_cell, HEX_BRAND_PURPLE)
    set_cell_margins(left_cell, top=200, bottom=200, left=300, right=200)
    set_cell_borders(left_cell, top=None, bottom=None, left=None, right=None)
    set_cell_vertical_align(left_cell, "center")
    for p in list(left_cell.paragraphs):
        p._element.getparent().remove(p._element)
    lp = left_cell.add_paragraph()
    add_run(lp, "CODESINO", bold=True, size_pt=12, color=WHITE_COLOR)

    set_cell_width(right_cell, 7360)
    set_cell_shading(right_cell, HEX_BRAND_PURPLE)
    set_cell_margins(right_cell, top=240, bottom=240, left=200, right=400)
    set_cell_borders(right_cell, top=None, bottom=None, left=None, right=None)
    set_cell_vertical_align(right_cell, "center")
    for p in list(right_cell.paragraphs):
        p._element.getparent().remove(p._element)

    p1 = right_cell.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_paragraph_spacing(p1, before=0, after=40)
    add_run(p1, "CODESINO SOFTWARE DEVELOPMENT SERVICES", bold=True, size_pt=20, color=WHITE_COLOR)

    p2 = right_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_paragraph_spacing(p2, before=0, after=40)
    add_run(p2, "Client Project Agreement", size_pt=13, color=hex_to_rgb("D4C8F0"))

    p3 = right_cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    no_space_after(p3)
    add_run(p3, "www.codesinodev.com", italic=True, size_pt=9, color=hex_to_rgb("B39DDB"))

    spacer(doc, 120, 40)

    # Ref + Date bar
    ref_table = doc.add_table(rows=1, cols=2)
    set_table_width(ref_table, 9360)
    set_column_widths(ref_table, [4680, 4680])
    set_table_no_borders(ref_table)

    ref_cell = ref_table.cell(0, 0)
    date_cell = ref_table.cell(0, 1)
    for cell in [ref_cell, date_cell]:
        set_cell_shading(cell, HEX_LIGHT_PURPLE)
        set_cell_margins(cell, top=80, bottom=80, left=160, right=100)
        set_cell_borders(cell, top=None, bottom=None, left=None, right=None)
        for p in list(cell.paragraphs):
            p._element.getparent().remove(p._element)

    rp = ref_cell.add_paragraph()
    no_space_after(rp)
    add_run(rp, "Agreement Ref:  ", bold=True, size_pt=10, color=BRAND_PURPLE)
    add_run(rp, agreement_ref, bold=True, size_pt=10, color=DARK_GRAY)

    set_cell_width(date_cell, 4680)
    dp = date_cell.add_paragraph()
    dp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    no_space_after(dp)
    add_run(dp, "Date & Time:  ", bold=True, size_pt=10, color=BRAND_PURPLE)
    add_run(dp, agreement_date, size_pt=10, color=DARK_GRAY)

    spacer(doc, 140, 60)

    # ═══════════════════════════════════════════
    # SECTION 1: PARTIES
    # ═══════════════════════════════════════════
    heading1(doc, "1. PARTIES TO THIS AGREEMENT")
    spacer(doc, 40, 40)

    heading2(doc, "1.1 Service Provider")

    def provider_content(cell):
        add_lv_to_cell(cell, "Company Name", "Codesino Software Development Services")
        add_lv_to_cell(cell, "Email", "contact@codesinodev.com")
        add_lv_to_cell(cell, "Phone / WhatsApp", "+2349036206457")
        add_lv_to_cell(cell, "Website", "https://www.codesinodev.com")

    section_box_table(doc, provider_content, HEX_WHITE)
    spacer(doc, 80, 60)

    heading2(doc, "1.2 Client")

    def client_content(cell):
        add_lv_to_cell(cell, "Client Full Name", data.get("clientName", ""))
        add_lv_to_cell(cell, "Company / Organization", data.get("clientCompany", ""))
        add_lv_to_cell(cell, "Email Address", data.get("clientEmail", ""))
        add_lv_to_cell(cell, "Phone / WhatsApp", data.get("clientPhone", ""))

    section_box_table(doc, client_content, HEX_WHITE)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 2: PROJECT OVERVIEW
    # ═══════════════════════════════════════════
    heading1(doc, "2. PROJECT OVERVIEW")
    spacer(doc, 40, 40)

    project_types = ["Website", "Web Application", "E-Commerce", "Web Portal", "Other"]
    selected_type = data.get("projectType", "")
    type_checkboxes = "        ".join(f"{cb(selected_type == t)}  {t}" for t in project_types)

    def overview_content(cell):
        add_lv_to_cell(cell, "Project Name", data.get("projectName", ""))
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=60)
        add_run(p, "Project Type:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(p, type_checkboxes, size_pt=11, color=DARK_GRAY)
        if selected_type == "Other" and data.get("projectTypeOther"):
            add_lv_to_cell(cell, "Other (specified)", data.get("projectTypeOther", ""))
        dp = cell.add_paragraph()
        add_paragraph_spacing(dp, before=0, after=40)
        add_run(dp, "Project Description:", bold=True, size_pt=11, color=BRAND_PURPLE)
        desc = data.get("projectDescription", "")
        ddp = cell.add_paragraph()
        add_paragraph_spacing(ddp, before=0, after=60)
        add_run(ddp, desc if desc else "Not Provided", italic=not bool(desc), size_pt=11,
                color=DARK_GRAY if desc else PLACEHOLDER)
        add_lv_to_cell(cell, "Target Audience / End Users", data.get("targetAudience", ""))

    section_box_table(doc, overview_content, HEX_WHITE)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 3: SCOPE OF WORK
    # ═══════════════════════════════════════════
    heading1(doc, "3. SCOPE OF WORK")
    italic_note(doc, "All features listed below have been agreed upon between both parties. Any feature not listed is considered out of scope and may attract additional charges.")
    spacer(doc, 40, 40)

    heading2(doc, "3.1 Core / Main Features")
    core_features = [f for f in (data.get("coreFeatures") or []) if f.get("feature")]
    if core_features:
        data_table(doc,
                   ["#", "Feature / Deliverable", "Description"],
                   [[str(i+1), f["feature"], f.get("description", "—")] for i, f in enumerate(core_features)],
                   [480, 3840, 5040])
    else:
        italic_note(doc, "No core features specified.")

    spacer(doc, 120, 60)
    heading2(doc, "3.2 Extra / Add-on Features")
    add_on_features = [f for f in (data.get("addOnFeatures") or []) if f.get("feature")]
    if add_on_features and data.get("hasAddOns"):
        data_table(doc,
                   ["#", "Extra Feature", "Notes / Condition"],
                   [[str(i+1), f["feature"], f.get("notes", "—")] for i, f in enumerate(add_on_features)],
                   [480, 3840, 5040])
    else:
        italic_note(doc, "No add-on features specified for this agreement.")

    spacer(doc, 120, 60)
    heading2(doc, "3.3 Explicitly Out of Scope")
    italic_note(doc, "These items are NOT included in this agreement. Additional requests for these will be quoted separately.")
    out_of_scope = [x for x in (data.get("outOfScope") or []) if x]
    if out_of_scope:
        for item in out_of_scope:
            add_numbered(doc, item)
    else:
        italic_note(doc, "No explicit out-of-scope items listed.")

    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 4: UI / DESIGN SPECIFICATIONS
    # ═══════════════════════════════════════════
    heading1(doc, "4. UI / DESIGN SPECIFICATIONS")
    spacer(doc, 40, 40)

    design_styles = ["Minimal", "Corporate", "Bold/Vibrant", "Elegant", "Playful", "Other"]
    selected_style = data.get("designStyle", "")
    style_checkboxes = "      ".join(f"{cb(selected_style == s)}  {s}" for s in design_styles)
    logo_provided = data.get("logoProvided", "")
    brand_assets = data.get("brandAssetsProvided", "")

    def design_content(cell):
        add_lv_to_cell(cell, "Primary Brand Color", data.get("primaryColor", ""))
        add_lv_to_cell(cell, "Secondary Color", data.get("secondaryColor", ""))
        add_lv_to_cell(cell, "Accent Color", data.get("accentColor", ""))
        add_lv_to_cell(cell, "Preferred Font(s)", data.get("preferredFonts", ""))
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=60)
        add_run(p, "Design Style / Mood:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(p, style_checkboxes, size_pt=11, color=DARK_GRAY)
        if selected_style == "Other" and data.get("designStyleOther"):
            add_lv_to_cell(cell, "Other (specified)", data.get("designStyleOther", ""))
        add_lv_to_cell(cell, "Reference / Inspiration Websites", data.get("inspirationSites", ""))
        lp = cell.add_paragraph()
        add_paragraph_spacing(lp, before=0, after=60)
        add_run(lp, "Logo Provided by Client:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(lp, f"{cb(logo_provided == 'yes')}  Yes      {cb(logo_provided == 'no')}  No (Codesino will create basic logo — quoted separately if required)", size_pt=11, color=DARK_GRAY)
        bp = cell.add_paragraph()
        add_paragraph_spacing(bp, before=0, after=60)
        add_run(bp, "Brand Assets / Style Guide Provided:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(bp, f"{cb(brand_assets == 'yes')}  Yes      {cb(brand_assets == 'no')}  No", size_pt=11, color=DARK_GRAY)
        rev = data.get("revisionRounds", "")
        add_lv_to_cell(cell, "Number of Design Revision Rounds",
                       f"{rev} (additional revisions billed separately)" if rev else "")

    section_box_table(doc, design_content, HEX_WHITE)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 5: PAGES & SCREENS
    # ═══════════════════════════════════════════
    heading1(doc, "5. PAGES & SCREENS INCLUDED")
    italic_note(doc, "All pages listed below are included in this agreement. Any page not listed is out of scope.")
    spacer(doc, 40, 40)

    default_pages = [
        {"name": "Home / Landing Page", "notes": "Default"},
        {"name": "About Page", "notes": "Default"},
        {"name": "Services / Products Page", "notes": "Default"},
        {"name": "Contact Page", "notes": "Default"},
    ]
    extra_pages = [p for p in (data.get("extraPages") or []) if p.get("name")]
    all_pages = default_pages + [{"name": p["name"], "notes": p.get("notes", "Additional")} for p in extra_pages]
    data_table(doc,
               ["#", "Page / Screen Name", "Notes"],
               [[str(i+1), p["name"], p["notes"]] for i, p in enumerate(all_pages)],
               [480, 4440, 4440])
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 6: TECHNICAL SPECIFICATIONS
    # ═══════════════════════════════════════════
    heading1(doc, "6. TECHNICAL SPECIFICATIONS")
    spacer(doc, 40, 40)

    domain_provided = data.get("domainProvidedBy", "")
    ssl_status = data.get("sslCertificate", "")
    mobile_resp = data.get("mobileResponsive", "")

    def tech_content(cell):
        add_lv_to_cell(cell, "Frontend Technology", data.get("frontendTech", ""))
        add_lv_to_cell(cell, "Backend Technology", data.get("backendTech", ""))
        add_lv_to_cell(cell, "Database", data.get("database", ""))
        add_lv_to_cell(cell, "Hosting Platform", data.get("hostingPlatform", ""))
        add_lv_to_cell(cell, "Domain Name", data.get("domainName", ""))
        dp = cell.add_paragraph()
        add_paragraph_spacing(dp, before=0, after=60)
        add_run(dp, "Domain Provided By:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(dp, f"{cb(domain_provided == 'client')}  Client      {cb(domain_provided == 'codesino')}  Codesino Development", size_pt=11, color=DARK_GRAY)
        sp = cell.add_paragraph()
        add_paragraph_spacing(sp, before=0, after=60)
        add_run(sp, "SSL Certificate:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(sp, f"{cb(ssl_status == 'included')}  Included      {cb(ssl_status == 'not-included')}  Not Included      {cb(ssl_status == 'client')}  Client to Arrange", size_pt=11, color=DARK_GRAY)
        mp = cell.add_paragraph()
        no_space_after(mp)
        add_run(mp, "Mobile Responsive:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(mp, f"{cb(mobile_resp == 'yes')}  Yes      {cb(mobile_resp == 'no')}  No", size_pt=11, color=DARK_GRAY)

    section_box_table(doc, tech_content, HEX_WHITE)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 7: DEVELOPER BRIEFS
    # ═══════════════════════════════════════════
    heading1(doc, "7. INTERNAL DEVELOPER BRIEFS")
    italic_note(doc, "⚠️  INTERNAL DOCUMENT — This section is for the Codesino development team only.")
    spacer(doc, 60, 40)

    heading2(doc, "7.1 Frontend Developer Notes")

    def fe_content(cell):
        add_lv_to_cell(cell, "Assigned Developer", data.get("frontendDev", ""))
        rp = cell.add_paragraph()
        add_paragraph_spacing(rp, before=0, after=60)
        add_run(rp, "Key Responsibilities (Default):", bold=True, size_pt=11, color=BRAND_PURPLE)
        for item in [
            "Build all pages listed in Section 5 according to the UI specs in Section 4",
            "Ensure full mobile responsiveness across all standard breakpoints",
            "Integrate with backend APIs as provided by the backend developer",
            "Cross-browser compatibility (Chrome, Firefox, Safari, Edge)",
            "Implement SEO-friendly markup and performance best practices",
        ]:
            add_bullet(cell, item)
        np = cell.add_paragraph()
        add_paragraph_spacing(np, before=60, after=60)
        add_run(np, "Special Notes / Extra Instructions:", bold=True, size_pt=11, color=BRAND_PURPLE)
        fn = data.get("frontendNotes", "")
        fp = cell.add_paragraph()
        no_space_after(fp)
        add_run(fp, fn if fn else "None", italic=not bool(fn), size_pt=11,
                color=DARK_GRAY if fn else PLACEHOLDER)

    section_box_table(doc, fe_content, HEX_LAVENDER)
    spacer(doc, 80, 60)

    heading2(doc, "7.2 Backend Developer Notes")

    def be_content(cell):
        add_lv_to_cell(cell, "Assigned Developer", data.get("backendDev", ""))
        rp = cell.add_paragraph()
        add_paragraph_spacing(rp, before=0, after=60)
        add_run(rp, "Key Responsibilities (Default):", bold=True, size_pt=11, color=BRAND_PURPLE)
        for item in [
            "Set up and configure the server, database, and hosting environment",
            "Build and document all API endpoints required by the frontend",
            "Handle authentication, data validation, and security hardening",
            "Database schema design and migration management",
            "Implement proper error handling, logging, and monitoring",
        ]:
            add_bullet(cell, item)
        np = cell.add_paragraph()
        add_paragraph_spacing(np, before=60, after=60)
        add_run(np, "Special Notes / Extra Instructions:", bold=True, size_pt=11, color=BRAND_PURPLE)
        bn = data.get("backendNotes", "")
        bp = cell.add_paragraph()
        no_space_after(bp)
        add_run(bp, bn if bn else "None", italic=not bool(bn), size_pt=11,
                color=DARK_GRAY if bn else PLACEHOLDER)

    section_box_table(doc, be_content, HEX_LAVENDER)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 8: TIMELINE
    # ═══════════════════════════════════════════
    heading1(doc, "8. PROJECT TIMELINE")
    italic_note(doc, "Timelines are subject to client feedback response times and payment completion.")
    spacer(doc, 40, 40)

    start_label = (format_date(data.get("startDate")) + " (Subject to payment confirmation)") if data.get("startDate") else "Not agreed"
    end_label = format_date(data.get("endDate")) if data.get("endDate") else "Not agreed"

    def timeline_content(cell):
        add_lv_to_cell(cell, "Estimated Start Date", start_label)
        add_lv_to_cell(cell, "Estimated Completion Date", end_label)
        add_lv_to_cell(cell, "Total Development Duration", data.get("duration", ""))

    section_box_table(doc, timeline_content, HEX_WHITE)
    spacer(doc, 80, 60)

    heading2(doc, "8.1 Project Milestones")
    milestones = [
        ("Project Kickoff (payment confirmed)", data.get("m1Date", "")),
        ("Design Mockups / Wireframes", data.get("m2Date", "")),
        ("Client Design Approval", data.get("m3Date", "")),
        ("Frontend Development Complete", data.get("m4Date", "")),
        ("Backend Development Complete", data.get("m5Date", "")),
        ("Integration & Testing", data.get("m6Date", "")),
        ("Client Review & Feedback", data.get("m7Date", "")),
        ("Final Revisions", data.get("m8Date", "")),
        ("Deployment / Go Live", data.get("m9Date", "")),
    ]
    data_table(doc,
               ["Milestone", "Expected Completion", "Status"],
               [[m, d, "☐  Pending"] for m, d in milestones],
               [4000, 3000, 2360])

    spacer(doc, 80, 60)
    heading2(doc, "8.2 Timeline Policy")
    body_text(doc, "Development will not begin until the required payment has been confirmed. The timeline starts from the date payment is confirmed, NOT the date this agreement is signed. Codesino is not responsible for delays caused by late client feedback, delayed provision of content or assets, change of project scope after sign-off, or unavailability of third-party services.")
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 9: PRICING & PAYMENT
    # ═══════════════════════════════════════════
    heading1(doc, "9. PRICING & PAYMENT TERMS")
    italic_note(doc, "Read carefully before signing. No development work begins until the required payment is received and confirmed.")
    spacer(doc, 60, 40)

    currency = data.get("currency", "ngn")
    currency_sym = "$" if currency == "usd" else "₦"
    total_amount = data.get("totalAmount", "")

    def pricing_box(cell):
        tp = cell.add_paragraph()
        add_paragraph_spacing(tp, before=0, after=60)
        add_run(tp, "Total Project Amount:  ", bold=True, size_pt=14, color=BRAND_PURPLE)
        add_run(tp, f"{currency_sym} {total_amount}", bold=True, size_pt=14, color=DARK_GRAY)
        add_lv_to_cell(cell, "Amount in Words", data.get("amountInWords", ""))

    section_box_table(doc, pricing_box, HEX_LAVENDER)
    spacer(doc, 80, 60)

    heading2(doc, "9.1 Payment Terms")

    payment_plan = data.get("paymentPlan", "")

    # Option A
    def opt_a_box(cell):
        is_sel = payment_plan == "A"
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=60)
        add_run(p, f"{cb(is_sel)}  OPTION A: Full Payment Before Kickoff",
                bold=True, size_pt=12,
                color=hex_to_rgb("1A7A40") if is_sel else hex_to_rgb("888888"))
        d = cell.add_paragraph()
        add_paragraph_spacing(d, before=0, after=40)
        add_run(d, "The client pays 100% of the total project amount before any development work begins.",
                size_pt=11, color=DARK_GRAY)
        if is_sel:
            add_lv_to_cell(cell, "Amount Due Upfront", f"{currency_sym} {total_amount}  (100%)")
            sp = cell.add_paragraph()
            add_paragraph_spacing(sp, before=0, after=40)
            st = data.get("optionAStatus", "")
            add_run(sp, "Payment Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(sp, f"{cb(st == 'received')}  Received & Confirmed        {cb(st == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if st == "received":
                add_lv_to_cell(cell, "Confirmation Date", format_date(data.get("optionADate", "")) if data.get("optionADate") else "")
                add_lv_to_cell(cell, "Confirmed Amount", f"{currency_sym} {data.get('optionAAmount', '')}")
            np = cell.add_paragraph()
            no_space_after(np)
            add_run(np, "Note: Development begins only after this payment is confirmed.",
                    italic=True, size_pt=10, color=hex_to_rgb("666666"))

    section_box_table(doc, opt_a_box, "F0FBF4" if payment_plan == "A" else "FAFAFA")
    spacer(doc, 80, 60)

    # Option B
    def opt_b_box(cell):
        is_sel = payment_plan == "B"
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=60)
        add_run(p, f"{cb(is_sel)}  OPTION B: 60% Upfront / 40% on Completion",
                bold=True, size_pt=12,
                color=BRAND_PURPLE if is_sel else hex_to_rgb("888888"))
        d = cell.add_paragraph()
        add_paragraph_spacing(d, before=0, after=40)
        add_run(d, "The client pays 60% before development begins. The remaining 40% is due upon project completion, before deployment to the live server.",
                size_pt=11, color=DARK_GRAY)
        if is_sel:
            h1 = cell.add_paragraph()
            add_paragraph_spacing(h1, before=40, after=40)
            add_run(h1, "First Payment (60%) — Due Before Kickoff", bold=True, size_pt=11, color=BRAND_PURPLE)
            add_lv_to_cell(cell, "Amount", f"{currency_sym} {data.get('optionB1Amount', '')}")
            s1p = cell.add_paragraph()
            add_paragraph_spacing(s1p, before=0, after=40)
            s1 = data.get("optionB1Status", "")
            add_run(s1p, "Payment Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(s1p, f"{cb(s1 == 'received')}  Received & Confirmed        {cb(s1 == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if s1 == "received":
                add_lv_to_cell(cell, "Confirmation Date", format_date(data.get("optionB1Date", "")) if data.get("optionB1Date") else "")
            h2 = cell.add_paragraph()
            add_paragraph_spacing(h2, before=60, after=40)
            add_run(h2, "Second Payment (40%) — Due Before Deployment", bold=True, size_pt=11, color=BRAND_PURPLE)
            add_lv_to_cell(cell, "Amount", f"{currency_sym} {data.get('optionB2Amount', '')}")
            s2p = cell.add_paragraph()
            add_paragraph_spacing(s2p, before=0, after=40)
            s2 = data.get("optionB2Status", "")
            add_run(s2p, "Payment Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(s2p, f"{cb(s2 == 'received')}  Received & Confirmed        {cb(s2 == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if s2 == "received":
                add_lv_to_cell(cell, "Confirmation Date", format_date(data.get("optionB2Date", "")) if data.get("optionB2Date") else "")
            np = cell.add_paragraph()
            no_space_after(np)
            add_run(np, "Note: The final 40% must be received before deployment. No project will go live on an outstanding balance.",
                    italic=True, size_pt=10, color=hex_to_rgb("666666"))

    section_box_table(doc, opt_b_box, HEX_LIGHT_PURPLE if payment_plan == "B" else "FAFAFA")
    spacer(doc, 80, 60)

    heading2(doc, "9.2 Payment Account Details")

    def ngn_account(cell):
        add_lv_to_cell(cell, "Bank Name", "KudaBank")
        add_lv_to_cell(cell, "Account Name", "Codesino Software Development Services")
        add_lv_to_cell(cell, "Account Number", "3003017268")
        add_lv_to_cell(cell, "Currency", "Nigerian Naira (₦)")

    def usd_account(cell):
        add_lv_to_cell(cell, "Bank Name", "— (USD account details pending)")
        add_lv_to_cell(cell, "Account Name", "— (to be added)")
        add_lv_to_cell(cell, "Account Number", "— (to be added)")
        add_lv_to_cell(cell, "Currency", "US Dollar ($)")

    section_box_table(doc, ngn_account if currency == "ngn" else usd_account, HEX_WHITE)
    spacer(doc, 80, 60)

    heading2(doc, "9.3 General Payment Policy")
    payment_policy = [
        "No development work will commence until the required upfront payment has been received and confirmed in writing by Codesino Software Development Services.",
        "All payments made are strictly non-refundable once the development phase has commenced.",
        "If the client cancels the project after development has started, all payments made up to that point are forfeited.",
        "If Codesino fails to deliver the agreed scope within the agreed timeline (excluding client-caused delays), a partial or full refund may be negotiated at the discretion of both parties.",
        "All amounts are in the currency agreed above. Cross-currency payment requests are subject to conversion at the prevailing rate on the date of payment.",
        "Codesino reserves the right to suspend work on any project where payment is overdue by more than 7 (seven) business days.",
        "A receipt/confirmation will be issued for every payment received.",
    ]
    for item in payment_policy:
        add_numbered(doc, item)

    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 10: CONTENT & ASSETS
    # ═══════════════════════════════════════════
    heading1(doc, "10. CONTENT & ASSETS")
    italic_note(doc, "Codesino builds the structure and functionality. Content delays by the client will result in corresponding timeline delays.")
    spacer(doc, 40, 40)

    asset_data = data.get("assets") or {}
    asset_items = [
        ("Website copy / written content", asset_data.get("websiteCopy")),
        ("Product / service images", asset_data.get("productImages")),
        ("Company logo (PNG/SVG)", asset_data.get("companyLogo")),
        ("Video content", asset_data.get("videoContent")),
        ("Staff / team photos", asset_data.get("teamPhotos")),
        ("Social media links", asset_data.get("socialLinks")),
    ]
    extra_assets = [a for a in (data.get("extraAssets") or []) if a.get("name")]
    for a in extra_assets:
        asset_items.append((a["name"], a.get("providedBy", "")))

    def asset_label(by):
        if by == "client":
            return "☑  Client  ☐  Codesino"
        if by == "codesino":
            return "☐  Client  ☑  Codesino"
        return "☐  Client  ☐  Codesino"

    data_table(doc,
               ["Asset / Content", "Provided By", "Notes"],
               [[name, asset_label(by), ""] for name, by in asset_items],
               [4000, 3000, 2360])
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 11: POST-DELIVERY
    # ═══════════════════════════════════════════
    heading1(doc, "11. POST-DELIVERY & SUPPORT")
    spacer(doc, 40, 40)

    sc_owner = data.get("sourceCodeOwner", "")
    hosting_after = data.get("hostingAfter", "")
    maintenance = data.get("maintenance", "")
    maint_fee = data.get("maintenanceFee", "")

    def support_content(cell):
        sp = data.get("supportPeriod", "")
        add_lv_to_cell(cell, "Free Support Period After Delivery",
                       f"{sp} days (bug fixes only — no new features)" if sp else "")
        scp = cell.add_paragraph()
        add_paragraph_spacing(scp, before=0, after=60)
        add_run(scp, "Source Code Handover:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(scp, f"{cb(sc_owner == 'included')}  Included      {cb(sc_owner == 'not-included')}  Not Included      {cb(sc_owner == 'extra')}  Available at Extra Cost",
                size_pt=11, color=DARK_GRAY)
        hp = cell.add_paragraph()
        add_paragraph_spacing(hp, before=0, after=60)
        add_run(hp, "Hosting Management After Delivery:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        add_run(hp, f"{cb(hosting_after == 'codesino')}  Managed by Codesino      {cb(hosting_after == 'client')}  Transferred to Client      {cb(hosting_after == 'tbd')}  TBD",
                size_pt=11, color=DARK_GRAY)
        mp = cell.add_paragraph()
        no_space_after(mp)
        add_run(mp, "Ongoing Maintenance Retainer:  ", bold=True, size_pt=11, color=BRAND_PURPLE)
        if maintenance == "agreed":
            add_run(mp, f"☑  Agreed — {currency_sym} {maint_fee} / month      ☐  Not Agreed",
                    size_pt=11, color=DARK_GRAY)
        else:
            add_run(mp, "☐  Agreed      ☑  Not Agreed", size_pt=11, color=DARK_GRAY)

    section_box_table(doc, support_content, HEX_WHITE)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 12: INTELLECTUAL PROPERTY
    # ═══════════════════════════════════════════
    heading1(doc, "12. INTELLECTUAL PROPERTY")
    ip_items = [
        "Upon receipt of full and final payment, all intellectual property rights for the final deliverables transfer entirely to the client.",
        "Codesino retains the right to display the completed project in its portfolio and marketing materials unless the client requests otherwise in writing prior to project completion.",
        "Any third-party libraries, plugins, frameworks, or open-source software used remain subject to their own respective licenses.",
        "Codesino retains full ownership of all custom code, internal tools, and reusable components used in the project until final payment is received and confirmed.",
        "The client warrants that all content, images, logos, and materials provided to Codesino are legally owned or licensed by the client and do not infringe any third-party intellectual property rights.",
    ]
    for item in ip_items:
        add_numbered(doc, item)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 13: CONFIDENTIALITY
    # ═══════════════════════════════════════════
    heading1(doc, "13. CONFIDENTIALITY")
    body_text(doc, "Both parties agree to keep strictly confidential any sensitive business information, trade secrets, proprietary data, client lists, pricing, or technical specifications shared during the course of this project. This obligation remains in full effect for two (2) years after the completion or termination of this agreement.")
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 14: TERMINATION
    # ═══════════════════════════════════════════
    heading1(doc, "14. TERMINATION")
    for item in [
        "Either party may terminate this agreement with 7 (seven) calendar days' written notice.",
        "If the client terminates the agreement after development has commenced, all payments made are non-refundable.",
        "If Codesino terminates the agreement without cause, a pro-rated refund will be issued for work not yet started.",
        "Termination must be communicated via email or documented written message to be legally valid.",
        "Upon termination, each party shall immediately return or destroy any confidential materials belonging to the other party.",
    ]:
        add_numbered(doc, item)
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 15: DISPUTE RESOLUTION
    # ═══════════════════════════════════════════
    heading1(doc, "15. DISPUTE RESOLUTION")
    body_text(doc, "In the event of a dispute, both parties agree to first attempt resolution through good-faith negotiation within 14 (fourteen) calendar days of the dispute being raised in writing. If resolution cannot be reached, both parties agree to submit the matter to mediation before resorting to formal legal proceedings. This agreement shall be governed by and construed in accordance with the laws of the Federal Republic of Nigeria.")
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 16: ADDITIONAL NOTES
    # ═══════════════════════════════════════════
    heading1(doc, "16. ADDITIONAL NOTES & SPECIAL CONDITIONS")
    notes = data.get("additionalNotes", "")
    if notes:
        body_text(doc, notes)
    else:
        italic_note(doc, "No additional notes for this agreement.")
    divider(doc)

    # ═══════════════════════════════════════════
    # SECTION 17: SIGNATURES
    # ═══════════════════════════════════════════
    heading1(doc, "17. AGREEMENT & SIGNATURES")
    body_text(doc, "By signing below, both parties confirm that they have read, understood, and agreed to all terms and conditions outlined in this Client Project Agreement. This document constitutes a legally binding agreement between Codesino Software Development Services and the client named herein.")
    spacer(doc, 80, 60)

    sig_table = doc.add_table(rows=2, cols=3)
    set_table_width(sig_table, 9360)
    set_column_widths(sig_table, [4400, 560, 4400])
    set_table_no_borders(sig_table)

    # Header row
    provider_hdr = sig_table.cell(0, 0)
    spacer_hdr   = sig_table.cell(0, 1)
    client_hdr   = sig_table.cell(0, 2)

    for cell in [provider_hdr, client_hdr]:
        set_cell_shading(cell, HEX_BRAND_PURPLE)
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        set_cell_borders(cell, top=None, bottom=None, left=None, right=None)
        for p in list(cell.paragraphs):
            p._element.getparent().remove(p._element)
    set_cell_borders(spacer_hdr, top=None, bottom=None, left=None, right=None)
    for p in list(spacer_hdr.paragraphs):
        p._element.getparent().remove(p._element)
    spacer_hdr.add_paragraph()

    ph = provider_hdr.add_paragraph()
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    no_space_after(ph)
    add_run(ph, "SERVICE PROVIDER", bold=True, size_pt=11, color=WHITE_COLOR)

    ch = client_hdr.add_paragraph()
    ch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    no_space_after(ch)
    add_run(ch, "CLIENT", bold=True, size_pt=11, color=WHITE_COLOR)

    # Sig row
    provider_sig = sig_table.cell(1, 0)
    spacer_sig   = sig_table.cell(1, 1)
    client_sig   = sig_table.cell(1, 2)

    thin_border = {"val": "single", "sz": 2, "color": HEX_C5B8E0}
    for sc in [provider_sig, client_sig]:
        set_cell_borders(sc, top=None, bottom=None, left=thin_border, right=thin_border)
        set_cell_margins(sc, top=100, bottom=120, left=150, right=150)
        for p in list(sc.paragraphs):
            p._element.getparent().remove(p._element)
    set_cell_borders(spacer_sig, top=None, bottom=None, left=None, right=None)
    for p in list(spacer_sig.paragraphs):
        p._element.getparent().remove(p._element)
    spacer_sig.add_paragraph()

    # Provider sig content
    pp1 = provider_sig.add_paragraph()
    add_paragraph_spacing(pp1, before=0, after=60)
    add_run(pp1, "Company: Codesino Software Development Services", size_pt=10, color=DARK_GRAY)

    pp2 = provider_sig.add_paragraph()
    add_paragraph_spacing(pp2, before=0, after=80)
    add_run(pp2, "Authorised Signatory", size_pt=10, color=DARK_GRAY)

    pp3 = provider_sig.add_paragraph()
    add_paragraph_spacing(pp3, before=0, after=20)
    add_run(pp3, "Signature:", size_pt=10, color=DARK_GRAY)
    sig_lines(provider_sig)

    pds = data.get("providerSignDate", "")
    pp4 = provider_sig.add_paragraph()
    no_space_after(pp4)
    add_run(pp4, f"Date: {format_date(pds) if pds else '___ / ___ / ______'}",
            italic=True, size_pt=10, color=hex_to_rgb("999999"))

    # Client sig content
    cc = data.get("clientCompany", "")
    cn = data.get("clientName", "")
    cp1 = client_sig.add_paragraph()
    add_paragraph_spacing(cp1, before=0, after=60)
    add_run(cp1, f"Company: {cc if cc else '____________________'}",
            italic=not bool(cc), size_pt=10,
            color=DARK_GRAY if cc else hex_to_rgb("999999"))

    cp2 = client_sig.add_paragraph()
    add_paragraph_spacing(cp2, before=0, after=80)
    add_run(cp2, f"Name: {cn if cn else '______________________'}",
            italic=not bool(cn), size_pt=10,
            color=DARK_GRAY if cn else hex_to_rgb("999999"))

    cp3 = client_sig.add_paragraph()
    add_paragraph_spacing(cp3, before=0, after=20)
    add_run(cp3, "Signature:", size_pt=10, color=DARK_GRAY)
    sig_lines(client_sig)

    cds = data.get("clientSignDate", "")
    cp4 = client_sig.add_paragraph()
    no_space_after(cp4)
    add_run(cp4, f"Date: {format_date(cds) if cds else '___ / ___ / ______'}",
            italic=True, size_pt=10, color=hex_to_rgb("999999"))

    spacer(doc, 120, 60)

    # Footer text
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    no_space_after(fp)
    add_run(fp, "This agreement was prepared by Codesino Software Development Services  |  www.codesinodev.com  |  contact@codesinodev.com",
            italic=True, size_pt=9, color=hex_to_rgb("9E9E9E"))

    # ── Page footer ──
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(footer_para, "Codesino Software Development Services  |  Client Project Agreement",
            size_pt=9, color=hex_to_rgb("9E9E9E"))

    # ── Serialize ──
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()