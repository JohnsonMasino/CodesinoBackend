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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.style import WD_STYLE_TYPE

# ─────────────────────────────────────────────────────────────────────────────
# BRAND COLORS  –  Professional Blue Palette
# ─────────────────────────────────────────────────────────────────────────────
BRAND_BLUE      = RGBColor(0x1A, 0x4A, 0x8A)   # #1A4A8A  deep navy-blue (primary)
MID_BLUE        = RGBColor(0x25, 0x6B, 0xC4)   # #256BC4  medium blue   (accent)
LIGHT_BLUE      = RGBColor(0xE8, 0xF1, 0xFB)   # #E8F1FB  very light blue (cell fill)
PALE_BLUE       = RGBColor(0xD0, 0xE4, 0xF7)   # #D0E4F7  lavender-blue  (alt fill)
DARK_GRAY       = RGBColor(0x22, 0x22, 0x22)   # #222222  near-black body text
MID_GRAY        = RGBColor(0x55, 0x55, 0x55)   # #555555  secondary text
WHITE_COLOR     = RGBColor(0xFF, 0xFF, 0xFF)   # #FFFFFF
GREEN_ACCENT    = RGBColor(0x19, 0x7A, 0x4A)   # #197A4A  payment confirmed
PLACEHOLDER     = RGBColor(0xAA, 0xAA, 0xAA)   # #AAAAAA  empty-field hint
MUTED_BLUE      = RGBColor(0x5A, 0x7E, 0xB5)   # #5A7EB5  italic notes

# Hex strings for XML shading  (no leading #)
HEX_BRAND_BLUE   = "1A4A8A"
HEX_MID_BLUE     = "256BC4"
HEX_LIGHT_BLUE   = "E8F1FB"
HEX_PALE_BLUE    = "D0E4F7"
HEX_DARK_GRAY    = "222222"
HEX_WHITE        = "FFFFFF"
HEX_RULE         = "B8CDE8"   # divider / border line colour
HEX_GREEN        = "197A4A"
HEX_FAFAFA       = "FAFAFA"
HEX_F0F8FF       = "F0F8FF"   # faint blue for selected option box


# ─────────────────────────────────────────────────────────────────────────────
# IMAGE PATHS
# ─────────────────────────────────────────────────────────────────────────────
_HERE       = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH   = os.path.join(_HERE, "images", "codesino_logo.png")
STAMP_PATH  = os.path.join(_HERE, "images", "codesino_stamp.png")


# ─────────────────────────────────────────────────────────────────────────────
# LOW-LEVEL XML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_cell_shading(cell, fill_hex):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_hex.upper())
    for ex in tcPr.findall(qn("w:shd")):
        tcPr.remove(ex)
    tcPr.append(shd)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tbl_borders = OxmlElement("w:tcBorders")
    for side, cfg in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        if cfg is None:
            el.set(qn("w:val"), "none")
        else:
            el.set(qn("w:val"),   cfg.get("val",   "single"))
            el.set(qn("w:sz"),    str(cfg.get("sz",  4)))
            el.set(qn("w:color"), cfg.get("color",  "000000"))
        tbl_borders.append(el)
    for ex in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(ex)
    tcPr.append(tbl_borders)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar  = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    for ex in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(ex)
    tcPr.append(mar)


def set_cell_vertical_align(cell, align="center"):
    tc     = cell._tc
    tcPr   = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), align)
    for ex in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(ex)
    tcPr.append(vAlign)


def set_cell_width(cell, width_dxa):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW  = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    for ex in tcPr.findall(qn("w:tcW")):
        tcPr.remove(ex)
    tcPr.append(tcW)


def add_paragraph_border_bottom(para, color=HEX_RULE, sz=6, space=4):
    pPr    = para._p.get_or_add_pPr()
    pBdr   = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    for ex in pPr.findall(qn("w:pBdr")):
        pPr.remove(ex)
    pPr.append(pBdr)


def add_paragraph_spacing(para, before=0, after=0, line=None):
    pPr     = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    if before is not None:
        spacing.set(qn("w:before"), str(before))
    if after is not None:
        spacing.set(qn("w:after"),  str(after))
    if line is not None:
        spacing.set(qn("w:line"),     str(line))
        spacing.set(qn("w:lineRule"), "auto")
    for ex in pPr.findall(qn("w:spacing")):
        pPr.remove(ex)
    pPr.append(spacing)


def set_run_font(run, font_name="Calibri"):
    rPr    = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    for ex in rPr.findall(qn("w:rFonts")):
        rPr.remove(ex)
    rPr.insert(0, rFonts)


def set_table_width(table, width_dxa):
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(width_dxa))
    tblW.set(qn("w:type"), "dxa")
    for ex in tblPr.findall(qn("w:tblW")):
        tblPr.remove(ex)
    tblPr.append(tblW)


def set_column_widths(table, widths):
    tbl    = table._tbl
    for ex in tbl.findall(qn("w:tblGrid")):
        tbl.remove(ex)
    tblGrid = OxmlElement("w:tblGrid")
    for w in widths:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        tblGrid.append(gridCol)
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        tblPr.addnext(tblGrid)
    else:
        tbl.insert(0, tblGrid)


def set_table_no_borders(table):
    tbl   = table._tbl
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


def no_space(para):
    add_paragraph_spacing(para, before=0, after=0)


def clear_cell_paras(cell):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)


# ─────────────────────────────────────────────────────────────────────────────
# MISC HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def generate_ref():
    now  = datetime.now()
    y    = str(now.year)[-2:]
    m    = str(now.month).zfill(2)
    d    = str(now.day).zfill(2)
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


def val_or(v, fallback="Not Provided"):
    return v if v else fallback


# ─────────────────────────────────────────────────────────────────────────────
# TYPOGRAPHY PRIMITIVES
# ─────────────────────────────────────────────────────────────────────────────

BODY_FONT = "Calibri"


def add_run(para, text, bold=False, italic=False, size_pt=11,
            color=None, font=BODY_FONT):
    run = para.add_run(text)
    run.bold        = bold
    run.italic      = italic
    run.font.size   = Pt(size_pt)
    run.font.name   = font
    if color:
        run.font.color.rgb = color if isinstance(color, RGBColor) else hex_to_rgb(color)
    set_run_font(run, font)
    return run


# ─────────────────────────────────────────────────────────────────────────────
# DOCUMENT-LEVEL BLOCK HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def heading1(doc, text):
    """Bold section heading with a solid blue underline rule."""
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=300, after=80)
    add_paragraph_border_bottom(para, color=HEX_BRAND_BLUE, sz=10, space=4)
    r = add_run(para, text, bold=True, size_pt=13, color=BRAND_BLUE)
    return para


def heading2(doc, text):
    """Sub-section heading in mid-blue."""
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=160, after=60)
    add_run(para, text, bold=True, size_pt=11, color=MID_BLUE)
    return para


def body_text(doc, text, italic=False):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=40, after=60, line=276)
    add_run(para, text, italic=italic, size_pt=11, color=DARK_GRAY)
    return para


def italic_note(doc, text):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=30, after=60)
    add_run(para, text, italic=True, size_pt=10, color=MUTED_BLUE)
    return para


def spacer(doc, before=80, after=80):
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=before, after=after)
    return para


def divider(doc):
    """Horizontal rule between major sections."""
    para = doc.add_paragraph()
    add_paragraph_spacing(para, before=180, after=180)
    add_paragraph_border_bottom(para, color=HEX_RULE, sz=2, space=1)
    return para


# ─────────────────────────────────────────────────────────────────────────────
# LABEL : VALUE PARAGRAPH
# ─────────────────────────────────────────────────────────────────────────────

def lv_para(container, label, value, label_color=None):
    """
    Renders   [Label]:  value
    container can be doc or a table cell.
    """
    if label_color is None:
        label_color = BRAND_BLUE
    para = container.add_paragraph()
    add_paragraph_spacing(para, before=0, after=56)
    r1 = para.add_run(f"{label}:  ")
    r1.bold             = True
    r1.font.size        = Pt(11)
    r1.font.color.rgb   = label_color
    r1.font.name        = BODY_FONT
    set_run_font(r1)
    is_empty = not value or value.strip() == ""
    r2 = para.add_run("Not Provided" if is_empty else value)
    r2.font.size      = Pt(11)
    r2.italic         = is_empty
    r2.font.color.rgb = PLACEHOLDER if is_empty else DARK_GRAY
    r2.font.name      = BODY_FONT
    set_run_font(r2)
    return para


# ─────────────────────────────────────────────────────────────────────────────
# INFO BOX  (1-column table with left accent bar)
# ─────────────────────────────────────────────────────────────────────────────

def section_box(doc, fill_fn, fill_hex=HEX_WHITE, accent_color=HEX_BRAND_BLUE):
    """
    Creates a full-width single-cell table styled as an info card.
    fill_fn(cell) → adds paragraphs into the cell.
    """
    table = doc.add_table(rows=1, cols=1)
    set_table_width(table,    9360)
    set_column_widths(table, [9360])
    set_table_no_borders(table)

    cell = table.cell(0, 0)
    set_cell_width(cell,    9360)
    set_cell_shading(cell,  fill_hex)
    set_cell_margins(cell,  top=140, bottom=140, left=220, right=160)
    set_cell_borders(
        cell,
        top    = {"val": "single", "sz": 2,  "color": HEX_RULE},
        bottom = {"val": "single", "sz": 2,  "color": HEX_RULE},
        left   = {"val": "single", "sz": 18, "color": accent_color},
        right  = None,
    )
    clear_cell_paras(cell)
    fill_fn(cell)
    return table


# ─────────────────────────────────────────────────────────────────────────────
# DATA TABLE  (header row + alternating body rows)
# ─────────────────────────────────────────────────────────────────────────────

def data_table(doc, headers, rows, col_widths):
    n_cols = len(headers)
    table  = doc.add_table(rows=1 + len(rows), cols=n_cols)
    set_table_width(table,    sum(col_widths))
    set_column_widths(table,  col_widths)
    set_table_no_borders(table)

    thin = {"val": "single", "sz": 2, "color": HEX_RULE}

    # ── Header row ───────────────────────────────────────────────────────────
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_width(cell,   col_widths[i])
        set_cell_shading(cell, HEX_BRAND_BLUE)
        set_cell_margins(cell, top=100, bottom=100, left=140, right=120)
        set_cell_borders(cell, top=thin, bottom=thin, left=thin, right=thin)
        clear_cell_paras(cell)
        p = cell.add_paragraph()
        no_space(p)
        add_run(p, h, bold=True, size_pt=10, color=WHITE_COLOR)

    # ── Data rows ────────────────────────────────────────────────────────────
    for ri, row_data in enumerate(rows):
        row  = table.rows[ri + 1]
        fill = HEX_WHITE if ri % 2 == 0 else HEX_LIGHT_BLUE
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            set_cell_width(cell,   col_widths[ci])
            set_cell_shading(cell, fill)
            set_cell_margins(cell, top=90, bottom=90, left=140, right=120)
            set_cell_borders(cell, top=thin, bottom=thin, left=thin, right=thin)
            clear_cell_paras(cell)
            p = cell.add_paragraph()
            no_space(p)
            add_run(p, str(val) if val is not None else "", size_pt=10, color=DARK_GRAY)

    return table


# ─────────────────────────────────────────────────────────────────────────────
# LIST HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def add_bullet(container, text):
    try:
        para = container.add_paragraph(style="List Bullet")
    except Exception:
        para = container.add_paragraph()
        add_run(para, "•  ", bold=False, size_pt=11, color=BRAND_BLUE)
    add_paragraph_spacing(para, before=30, after=30)
    for r in list(para.runs):
        r._r.getparent().remove(r._r)
    add_run(para, text, size_pt=11, color=DARK_GRAY)
    return para


def add_numbered(container, text):
    try:
        para = container.add_paragraph(style="List Number")
    except Exception:
        para = container.add_paragraph()
    add_paragraph_spacing(para, before=30, after=30)
    for r in list(para.runs):
        r._r.getparent().remove(r._r)
    add_run(para, text, size_pt=11, color=DARK_GRAY)
    return para


# ─────────────────────────────────────────────────────────────────────────────
# SIGNATURE LINE HELPER
# ─────────────────────────────────────────────────────────────────────────────

def sig_line(cell, label="Signature:"):
    p1 = cell.add_paragraph()
    add_paragraph_spacing(p1, before=0, after=20)
    add_run(p1, label, size_pt=10, color=MID_GRAY)
    # blank spacer rows
    for _ in range(2):
        pb = cell.add_paragraph()
        no_space(pb)
    # underline rule
    pu = cell.add_paragraph()
    add_paragraph_spacing(pu, before=0, after=60)
    add_paragraph_border_bottom(pu, color=HEX_BRAND_BLUE, sz=4, space=1)


# ─────────────────────────────────────────────────────────────────────────────
# INLINE IMAGE HELPER
# ─────────────────────────────────────────────────────────────────────────────

def add_image_para(container, path, width_inches, align=WD_ALIGN_PARAGRAPH.LEFT):
    """Add an image paragraph; silently skips if file is missing."""
    if not os.path.exists(path):
        return None
    para = container.add_paragraph()
    no_space(para)
    para.alignment = align
    run  = para.add_run()
    run.add_picture(path, width=Inches(width_inches))
    return para


# ─────────────────────────────────────────────────────────────────────────────
# MAIN GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

def generate_agreement_buffer(data: dict) -> bytes:
    doc = Document()

    # ── Page setup  (US Letter, 1-inch margins) ──────────────────────────────
    for section in doc.sections:
        section.page_width    = Twips(12240)
        section.page_height   = Twips(15840)
        section.left_margin   = Twips(1440)
        section.right_margin  = Twips(1440)
        section.top_margin    = Twips(1080)
        section.bottom_margin = Twips(1080)

    # ── Ensure list styles exist ─────────────────────────────────────────────
    for style_name, style_type in [
        ("List Bullet", WD_STYLE_TYPE.PARAGRAPH),
        ("List Number", WD_STYLE_TYPE.PARAGRAPH),
    ]:
        try:
            doc.styles[style_name]
        except KeyError:
            doc.styles.add_style(style_name, style_type)

    agreement_ref  = generate_ref()
    agreement_date = format_datetime()
    currency       = data.get("currency", "ngn")
    currency_sym   = "$" if currency == "usd" else "₦"

    # ═════════════════════════════════════════════════════════════════════════
    # COVER HEADER  –  logo + title block + ref/date bar
    # ═════════════════════════════════════════════════════════════════════════

    # ── Top bar: logo (left) | title text (right) ────────────────────────────
    hdr_table = doc.add_table(rows=1, cols=2)
    set_table_width(hdr_table,    9360)
    set_column_widths(hdr_table, [2200, 7160])
    set_table_no_borders(hdr_table)

    logo_cell  = hdr_table.cell(0, 0)
    title_cell = hdr_table.cell(0, 1)

    # Logo cell — white background, logo image centred
    set_cell_width(logo_cell,    2200)
    set_cell_shading(logo_cell,  HEX_WHITE)
    set_cell_margins(logo_cell,  top=100, bottom=100, left=100, right=100)
    set_cell_borders(logo_cell,  top=None, bottom=None, left=None, right=None)
    set_cell_vertical_align(logo_cell, "center")
    clear_cell_paras(logo_cell)
    lp = logo_cell.add_paragraph()
    no_space(lp)
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(LOGO_PATH):
        lp.add_run().add_picture(LOGO_PATH, width=Inches(1.35))
    else:
        add_run(lp, "CODESINO", bold=True, size_pt=14, color=BRAND_BLUE)

    # Title cell — deep blue background
    set_cell_width(title_cell,    7160)
    set_cell_shading(title_cell,  HEX_BRAND_BLUE)
    set_cell_margins(title_cell,  top=240, bottom=240, left=260, right=300)
    set_cell_borders(title_cell,  top=None, bottom=None, left=None, right=None)
    set_cell_vertical_align(title_cell, "center")
    clear_cell_paras(title_cell)

    tp1 = title_cell.add_paragraph()
    add_paragraph_spacing(tp1, before=0, after=40)
    tp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(tp1, "CODESINO SOFTWARE DEVELOPMENT SERVICES",
            bold=True, size_pt=15, color=WHITE_COLOR)

    tp2 = title_cell.add_paragraph()
    add_paragraph_spacing(tp2, before=0, after=40)
    tp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(tp2, "Client Project Agreement",
            size_pt=12, color=hex_to_rgb("C8DCFA"))

    tp3 = title_cell.add_paragraph()
    no_space(tp3)
    tp3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(tp3, "www.codesinodev.com",
            italic=True, size_pt=9, color=hex_to_rgb("92B8E8"))

    spacer(doc, 60, 30)

    # ── Ref / Date bar ───────────────────────────────────────────────────────
    ref_table = doc.add_table(rows=1, cols=2)
    set_table_width(ref_table,    9360)
    set_column_widths(ref_table, [4680, 4680])
    set_table_no_borders(ref_table)

    for idx, (label, value, align) in enumerate([
        ("Agreement Reference", agreement_ref,  WD_ALIGN_PARAGRAPH.LEFT),
        ("Date & Time Generated", agreement_date, WD_ALIGN_PARAGRAPH.RIGHT),
    ]):
        cell = ref_table.cell(0, idx)
        set_cell_shading(cell,  HEX_LIGHT_BLUE)
        set_cell_margins(cell,  top=90, bottom=90, left=160, right=160)
        set_cell_borders(cell,
                         top    = {"val": "single", "sz": 4,  "color": HEX_MID_BLUE},
                         bottom = {"val": "single", "sz": 4,  "color": HEX_MID_BLUE},
                         left   = None, right = None)
        clear_cell_paras(cell)
        p = cell.add_paragraph()
        no_space(p)
        p.alignment = align
        add_run(p, f"{label}:  ", bold=True, size_pt=9.5, color=BRAND_BLUE)
        add_run(p, value,         bold=False, size_pt=9.5, color=DARK_GRAY)

    spacer(doc, 160, 60)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 1 — PARTIES
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "1.  PARTIES TO THIS AGREEMENT")
    spacer(doc, 40, 40)
    heading2(doc, "1.1  Service Provider")

    def provider_fn(cell):
        lv_para(cell, "Company Name",       "Codesino Software Development Services")
        lv_para(cell, "Email",              "contact@codesinodev.com")
        lv_para(cell, "Phone / WhatsApp",   "+2349036206457")
        lv_para(cell, "Website",            "https://www.codesinodev.com")

    section_box(doc, provider_fn, HEX_WHITE)
    spacer(doc, 80, 60)
    heading2(doc, "1.2  Client")

    def client_fn(cell):
        lv_para(cell, "Client Full Name",         data.get("clientName",    ""))
        lv_para(cell, "Company / Organisation",   data.get("clientCompany", ""))
        lv_para(cell, "Email Address",            data.get("clientEmail",   ""))
        lv_para(cell, "Phone / WhatsApp",         data.get("clientPhone",   ""))

    section_box(doc, client_fn, HEX_WHITE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 2 — PROJECT OVERVIEW
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "2.  PROJECT OVERVIEW")
    spacer(doc, 40, 40)

    proj_types    = ["Website", "Web Application", "E-Commerce", "Web Portal", "Other"]
    sel_type      = data.get("projectType", "")
    type_cbs      = "      ".join(f"{cb(sel_type == t)}  {t}" for t in proj_types)

    def overview_fn(cell):
        lv_para(cell, "Project Name", data.get("projectName", ""))
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=56)
        add_run(p, "Project Type:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(p, type_cbs, size_pt=11, color=DARK_GRAY)
        if sel_type == "Other" and data.get("projectTypeOther"):
            lv_para(cell, "Other (specified)", data.get("projectTypeOther", ""))
        dp = cell.add_paragraph()
        add_paragraph_spacing(dp, before=0, after=30)
        add_run(dp, "Project Description:", bold=True, size_pt=11, color=BRAND_BLUE)
        desc = data.get("projectDescription", "")
        ddp  = cell.add_paragraph()
        add_paragraph_spacing(ddp, before=0, after=56, line=276)
        add_run(ddp, desc if desc else "Not Provided",
                italic=not bool(desc), size_pt=11,
                color=DARK_GRAY if desc else PLACEHOLDER)
        lv_para(cell, "Target Audience / End Users", data.get("targetAudience", ""))

    section_box(doc, overview_fn, HEX_WHITE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 3 — SCOPE OF WORK
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "3.  SCOPE OF WORK")
    italic_note(doc, "All features listed below have been agreed upon by both parties. Any feature not listed is considered out of scope and may attract additional charges.")
    spacer(doc, 40, 40)

    heading2(doc, "3.1  Core / Main Features")
    core = [f for f in (data.get("coreFeatures") or []) if f.get("feature")]
    if core:
        data_table(doc,
                   ["#", "Feature / Deliverable", "Description"],
                   [[str(i + 1), f["feature"], f.get("description", "—")] for i, f in enumerate(core)],
                   [480, 3840, 5040])
    else:
        italic_note(doc, "No core features specified.")

    spacer(doc, 120, 60)
    heading2(doc, "3.2  Extra / Add-on Features")
    add_ons = [f for f in (data.get("addOnFeatures") or []) if f.get("feature")]
    if add_ons and data.get("hasAddOns"):
        data_table(doc,
                   ["#", "Extra Feature", "Notes / Condition"],
                   [[str(i + 1), f["feature"], f.get("notes", "—")] for i, f in enumerate(add_ons)],
                   [480, 3840, 5040])
    else:
        italic_note(doc, "No add-on features specified for this agreement.")

    spacer(doc, 120, 60)
    heading2(doc, "3.3  Explicitly Out of Scope")
    italic_note(doc, "The following items are NOT included in this agreement. Any additional requests will be quoted separately.")
    oos = [x for x in (data.get("outOfScope") or []) if x]
    if oos:
        for item in oos:
            add_numbered(doc, item)
    else:
        italic_note(doc, "No explicit out-of-scope items listed.")

    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 4 — DESIGN SPECIFICATIONS
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "4.  UI / DESIGN SPECIFICATIONS")
    spacer(doc, 40, 40)

    des_styles  = ["Minimal", "Corporate", "Bold/Vibrant", "Elegant", "Playful", "Other"]
    sel_style   = data.get("designStyle", "")
    style_cbs   = "    ".join(f"{cb(sel_style == s)}  {s}" for s in des_styles)
    logo_prov   = data.get("logoProvided", "")
    brand_assets= data.get("brandAssetsProvided", "")

    def design_fn(cell):
        lv_para(cell, "Primary Brand Colour",    data.get("primaryColor",   ""))
        lv_para(cell, "Secondary Colour",         data.get("secondaryColor", ""))
        lv_para(cell, "Accent Colour",            data.get("accentColor",    ""))
        lv_para(cell, "Preferred Font(s)",        data.get("preferredFonts", ""))
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=56)
        add_run(p, "Design Style / Mood:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(p, style_cbs, size_pt=11, color=DARK_GRAY)
        if sel_style == "Other" and data.get("designStyleOther"):
            lv_para(cell, "Other (specified)", data.get("designStyleOther", ""))
        lv_para(cell, "Inspiration Websites", data.get("inspirationSites", ""))
        lp = cell.add_paragraph()
        add_paragraph_spacing(lp, before=0, after=56)
        add_run(lp, "Logo Provided by Client:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(lp, f"{cb(logo_prov == 'yes')}  Yes      {cb(logo_prov == 'no')}  No (Codesino will create basic logo — quoted separately)",
                size_pt=11, color=DARK_GRAY)
        bp = cell.add_paragraph()
        add_paragraph_spacing(bp, before=0, after=56)
        add_run(bp, "Brand Assets / Style Guide Provided:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(bp, f"{cb(brand_assets == 'yes')}  Yes      {cb(brand_assets == 'no')}  No",
                size_pt=11, color=DARK_GRAY)
        rev = data.get("revisionRounds", "")
        lv_para(cell, "Design Revision Rounds",
                f"{rev} rounds (additional revisions billed separately)" if rev else "")

    section_box(doc, design_fn, HEX_WHITE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 5 — PAGES & SCREENS
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "5.  PAGES & SCREENS INCLUDED")
    italic_note(doc, "All pages listed below are included in this agreement. Any page not listed is out of scope.")
    spacer(doc, 40, 40)

    default_pages = [
        {"name": "Home / Landing Page",       "notes": "Default"},
        {"name": "About Page",                "notes": "Default"},
        {"name": "Services / Products Page",  "notes": "Default"},
        {"name": "Contact Page",              "notes": "Default"},
    ]
    extra_pages = [p for p in (data.get("extraPages") or []) if p.get("name")]
    all_pages   = default_pages + [{"name": p["name"], "notes": p.get("notes", "Additional")}
                                    for p in extra_pages]
    data_table(doc,
               ["#", "Page / Screen Name", "Notes"],
               [[str(i + 1), p["name"], p["notes"]] for i, p in enumerate(all_pages)],
               [480, 4440, 4440])
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 6 — TECHNICAL SPECIFICATIONS
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "6.  TECHNICAL SPECIFICATIONS")
    spacer(doc, 40, 40)

    dom_by   = data.get("domainProvidedBy", "")
    ssl_stat = data.get("sslCertificate",   "")
    mob_resp = data.get("mobileResponsive", "")

    def tech_fn(cell):
        lv_para(cell, "Frontend Technology",  data.get("frontendTech",    ""))
        lv_para(cell, "Backend Technology",   data.get("backendTech",     ""))
        lv_para(cell, "Database",             data.get("database",        ""))
        lv_para(cell, "Hosting Platform",     data.get("hostingPlatform", ""))
        lv_para(cell, "Domain Name",          data.get("domainName",      ""))
        dp = cell.add_paragraph()
        add_paragraph_spacing(dp, before=0, after=56)
        add_run(dp, "Domain Provided By:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(dp, f"{cb(dom_by == 'client')}  Client      {cb(dom_by == 'codesino')}  Codesino Development",
                size_pt=11, color=DARK_GRAY)
        sp = cell.add_paragraph()
        add_paragraph_spacing(sp, before=0, after=56)
        add_run(sp, "SSL Certificate:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(sp, f"{cb(ssl_stat == 'included')}  Included      {cb(ssl_stat == 'not-included')}  Not Included      {cb(ssl_stat == 'client')}  Client to Arrange",
                size_pt=11, color=DARK_GRAY)
        mp = cell.add_paragraph()
        no_space(mp)
        add_run(mp, "Mobile Responsive:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(mp, f"{cb(mob_resp == 'yes')}  Yes      {cb(mob_resp == 'no')}  No",
                size_pt=11, color=DARK_GRAY)

    section_box(doc, tech_fn, HEX_WHITE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 7 — DEVELOPER BRIEFS
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "7.  INTERNAL DEVELOPER BRIEFS")
    italic_note(doc, "⚠  INTERNAL USE ONLY — This section is for the Codesino development team and is not intended for the client.")
    spacer(doc, 60, 40)

    heading2(doc, "7.1  Frontend Developer Notes")

    def fe_fn(cell):
        lv_para(cell, "Assigned Developer", data.get("frontendDev", ""))
        hp = cell.add_paragraph()
        add_paragraph_spacing(hp, before=40, after=40)
        add_run(hp, "Default Responsibilities:", bold=True, size_pt=11, color=BRAND_BLUE)
        for item in [
            "Build all pages per the UI specs in Section 4",
            "Ensure full mobile responsiveness across all standard breakpoints",
            "Integrate with backend APIs provided by the backend developer",
            "Cross-browser compatibility (Chrome, Firefox, Safari, Edge)",
            "Implement SEO-friendly markup and performance best practices",
        ]:
            add_bullet(cell, item)
        np = cell.add_paragraph()
        add_paragraph_spacing(np, before=60, after=30)
        add_run(np, "Special Notes / Extra Instructions:", bold=True, size_pt=11, color=BRAND_BLUE)
        fn  = data.get("frontendNotes", "")
        fnp = cell.add_paragraph()
        no_space(fnp)
        add_run(fnp, fn if fn else "None", italic=not bool(fn), size_pt=11,
                color=DARK_GRAY if fn else PLACEHOLDER)

    section_box(doc, fe_fn, HEX_LIGHT_BLUE, accent_color=HEX_MID_BLUE)
    spacer(doc, 80, 60)

    heading2(doc, "7.2  Backend Developer Notes")

    def be_fn(cell):
        lv_para(cell, "Assigned Developer", data.get("backendDev", ""))
        hp = cell.add_paragraph()
        add_paragraph_spacing(hp, before=40, after=40)
        add_run(hp, "Default Responsibilities:", bold=True, size_pt=11, color=BRAND_BLUE)
        for item in [
            "Set up and configure server, database, and hosting environment",
            "Build and document all API endpoints required by the frontend",
            "Handle authentication, data validation, and security hardening",
            "Database schema design and migration management",
            "Implement error handling, logging, and monitoring",
        ]:
            add_bullet(cell, item)
        np = cell.add_paragraph()
        add_paragraph_spacing(np, before=60, after=30)
        add_run(np, "Special Notes / Extra Instructions:", bold=True, size_pt=11, color=BRAND_BLUE)
        bn  = data.get("backendNotes", "")
        bnp = cell.add_paragraph()
        no_space(bnp)
        add_run(bnp, bn if bn else "None", italic=not bool(bn), size_pt=11,
                color=DARK_GRAY if bn else PLACEHOLDER)

    section_box(doc, be_fn, HEX_LIGHT_BLUE, accent_color=HEX_MID_BLUE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 8 — TIMELINE
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "8.  PROJECT TIMELINE")
    italic_note(doc, "Timelines are subject to client feedback response times and payment completion.")
    spacer(doc, 40, 40)

    start_label = (format_date(data.get("startDate")) + " (subject to payment confirmation)") \
                  if data.get("startDate") else "Not agreed"
    end_label   = format_date(data.get("endDate")) if data.get("endDate") else "Not agreed"

    def timeline_fn(cell):
        lv_para(cell, "Estimated Start Date",       start_label)
        lv_para(cell, "Estimated Completion Date",  end_label)
        lv_para(cell, "Total Development Duration", data.get("duration", ""))

    section_box(doc, timeline_fn, HEX_WHITE)
    spacer(doc, 80, 60)

    heading2(doc, "8.1  Project Milestones")
    milestones = [
        ("Project Kickoff (payment confirmed)",   data.get("m1Date", "")),
        ("Design Mockups / Wireframes",            data.get("m2Date", "")),
        ("Client Design Approval",                 data.get("m3Date", "")),
        ("Frontend Development Complete",          data.get("m4Date", "")),
        ("Backend Development Complete",           data.get("m5Date", "")),
        ("Integration & Testing",                  data.get("m6Date", "")),
        ("Client Review & Feedback",               data.get("m7Date", "")),
        ("Final Revisions",                        data.get("m8Date", "")),
        ("Deployment / Go Live",                   data.get("m9Date", "")),
    ]
    data_table(doc,
               ["Milestone", "Expected Completion", "Status"],
               [[m, format_date(d) if d else "TBC", "☐  Pending"] for m, d in milestones],
               [4000, 3000, 2360])

    spacer(doc, 80, 60)
    heading2(doc, "8.2  Timeline Policy")
    body_text(doc,
              "Development does not begin until the required upfront payment is confirmed. "
              "The timeline starts from the date payment is received, not the date this "
              "agreement is signed. Codesino is not responsible for delays caused by late "
              "client feedback, delayed provision of content or assets, change of scope "
              "after sign-off, or unavailability of third-party services.")
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 9 — PRICING & PAYMENT
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "9.  PRICING & PAYMENT TERMS")
    italic_note(doc, "No development work begins until the required payment is received and confirmed in writing.")
    spacer(doc, 60, 40)

    total_amount  = data.get("totalAmount", "")
    payment_plan  = data.get("paymentPlan", "")

    def pricing_fn(cell):
        tp = cell.add_paragraph()
        add_paragraph_spacing(tp, before=0, after=56)
        add_run(tp, "Total Project Amount:  ", bold=True, size_pt=14, color=BRAND_BLUE)
        add_run(tp, f"{currency_sym}{total_amount}", bold=True, size_pt=14, color=DARK_GRAY)
        lv_para(cell, "Amount in Words", data.get("amountInWords", ""))

    section_box(doc, pricing_fn, HEX_PALE_BLUE, accent_color=HEX_MID_BLUE)
    spacer(doc, 80, 60)
    heading2(doc, "9.1  Payment Terms")

    # Option A
    def opt_a_fn(cell):
        is_sel = payment_plan == "A"
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=56)
        add_run(p, f"{cb(is_sel)}  OPTION A: Full Payment Before Kickoff",
                bold=True, size_pt=12,
                color=GREEN_ACCENT if is_sel else hex_to_rgb("888888"))
        d = cell.add_paragraph()
        add_paragraph_spacing(d, before=0, after=40)
        add_run(d, "The client pays 100% of the total project amount before any development begins.",
                size_pt=11, color=DARK_GRAY)
        if is_sel:
            lv_para(cell, "Amount Due Upfront", f"{currency_sym}{total_amount}  (100%)")
            sp = cell.add_paragraph()
            add_paragraph_spacing(sp, before=0, after=40)
            st = data.get("optionAStatus", "")
            add_run(sp, "Payment Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(sp, f"{cb(st == 'received')}  Received & Confirmed        {cb(st == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if st == "received":
                lv_para(cell, "Confirmation Date",
                        format_date(data.get("optionADate", "")) if data.get("optionADate") else "")
                lv_para(cell, "Confirmed Amount", f"{currency_sym}{data.get('optionAAmount', '')}")
            np = cell.add_paragraph()
            no_space(np)
            add_run(np, "Note: Development begins only after this payment is fully confirmed.",
                    italic=True, size_pt=10, color=MID_GRAY)

    section_box(doc, opt_a_fn, HEX_F0F8FF if payment_plan == "A" else HEX_FAFAFA,
                accent_color=HEX_GREEN if payment_plan == "A" else HEX_RULE)
    spacer(doc, 80, 60)

    # Option B
    def opt_b_fn(cell):
        is_sel = payment_plan == "B"
        p = cell.add_paragraph()
        add_paragraph_spacing(p, before=0, after=56)
        add_run(p, f"{cb(is_sel)}  OPTION B: 60% Upfront / 40% on Completion",
                bold=True, size_pt=12,
                color=BRAND_BLUE if is_sel else hex_to_rgb("888888"))
        d = cell.add_paragraph()
        add_paragraph_spacing(d, before=0, after=40)
        add_run(d, "The client pays 60% before development begins. The remaining 40% is due upon completion, before deployment.",
                size_pt=11, color=DARK_GRAY)
        if is_sel:
            h1 = cell.add_paragraph()
            add_paragraph_spacing(h1, before=40, after=30)
            add_run(h1, "First Payment — 60%  (Due Before Kickoff)", bold=True, size_pt=11, color=BRAND_BLUE)
            lv_para(cell, "Amount", f"{currency_sym}{data.get('optionB1Amount', '')}")
            s1p = cell.add_paragraph()
            add_paragraph_spacing(s1p, before=0, after=40)
            s1 = data.get("optionB1Status", "")
            add_run(s1p, "Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(s1p, f"{cb(s1 == 'received')}  Received & Confirmed        {cb(s1 == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if s1 == "received":
                lv_para(cell, "Confirmation Date",
                        format_date(data.get("optionB1Date", "")) if data.get("optionB1Date") else "")

            h2 = cell.add_paragraph()
            add_paragraph_spacing(h2, before=60, after=30)
            add_run(h2, "Second Payment — 40%  (Due Before Deployment)", bold=True, size_pt=11, color=BRAND_BLUE)
            lv_para(cell, "Amount", f"{currency_sym}{data.get('optionB2Amount', '')}")
            s2p = cell.add_paragraph()
            add_paragraph_spacing(s2p, before=0, after=40)
            s2 = data.get("optionB2Status", "")
            add_run(s2p, "Status:  ", bold=True, size_pt=11, color=DARK_GRAY)
            add_run(s2p, f"{cb(s2 == 'received')}  Received & Confirmed        {cb(s2 == 'pending')}  Pending",
                    size_pt=11, color=DARK_GRAY)
            if s2 == "received":
                lv_para(cell, "Confirmation Date",
                        format_date(data.get("optionB2Date", "")) if data.get("optionB2Date") else "")

            np = cell.add_paragraph()
            no_space(np)
            add_run(np, "Note: The final 40% must be received before deployment. No project goes live on an outstanding balance.",
                    italic=True, size_pt=10, color=MID_GRAY)

    section_box(doc, opt_b_fn, HEX_LIGHT_BLUE if payment_plan == "B" else HEX_FAFAFA,
                accent_color=HEX_BRAND_BLUE if payment_plan == "B" else HEX_RULE)
    spacer(doc, 80, 60)

    heading2(doc, "9.2  Payment Account Details")

    def ngn_fn(cell):
        lv_para(cell, "Bank Name",       "KudaBank")
        lv_para(cell, "Account Name",    "Codesino Software Development Services")
        lv_para(cell, "Account Number",  "3003017268")
        lv_para(cell, "Currency",        "Nigerian Naira (₦)")

    def usd_fn(cell):
        lv_para(cell, "Bank Name",      "— (USD account details pending)")
        lv_para(cell, "Account Name",   "— (to be added)")
        lv_para(cell, "Account Number", "— (to be added)")
        lv_para(cell, "Currency",       "US Dollar ($)")

    section_box(doc, ngn_fn if currency == "ngn" else usd_fn, HEX_WHITE)
    spacer(doc, 80, 60)

    heading2(doc, "9.3  General Payment Policy")
    payment_policy = [
        "No development work commences until the required upfront payment has been received and confirmed in writing by Codesino Software Development Services.",
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

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 10 — CONTENT & ASSETS
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "10.  CONTENT & ASSETS")
    italic_note(doc, "Codesino builds the structure and functionality. Content delays by the client will result in corresponding timeline delays.")
    spacer(doc, 40, 40)

    asset_data  = data.get("assets") or {}
    asset_items = [
        ("Website copy / written content",  asset_data.get("websiteCopy")),
        ("Product / service images",        asset_data.get("productImages")),
        ("Company logo (PNG / SVG)",        asset_data.get("companyLogo")),
        ("Video content",                   asset_data.get("videoContent")),
        ("Staff / team photos",             asset_data.get("teamPhotos")),
        ("Social media links",              asset_data.get("socialLinks")),
    ]
    extra_assets = [a for a in (data.get("extraAssets") or []) if a.get("name")]
    for a in extra_assets:
        asset_items.append((a["name"], a.get("providedBy", "")))

    def asset_label(by):
        if by == "client":   return "☑  Client   ☐  Codesino"
        if by == "codesino": return "☐  Client   ☑  Codesino"
        return "☐  Client   ☐  Codesino"

    data_table(doc,
               ["Asset / Content", "Provided By", "Notes"],
               [[name, asset_label(by), ""] for name, by in asset_items],
               [4000, 3000, 2360])
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 11 — POST-DELIVERY & SUPPORT
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "11.  POST-DELIVERY & SUPPORT")
    spacer(doc, 40, 40)

    sc_owner     = data.get("sourceCodeOwner", "")
    hosting_aft  = data.get("hostingAfter",    "")
    maintenance  = data.get("maintenance",     "")
    maint_fee    = data.get("maintenanceFee",  "")

    def support_fn(cell):
        sp = data.get("supportPeriod", "")
        lv_para(cell, "Free Support Period After Delivery",
                f"{sp} days (bug fixes only — no new features)" if sp else "")
        scp = cell.add_paragraph()
        add_paragraph_spacing(scp, before=0, after=56)
        add_run(scp, "Source Code Handover:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(scp, f"{cb(sc_owner == 'included')}  Included      "
                     f"{cb(sc_owner == 'not-included')}  Not Included      "
                     f"{cb(sc_owner == 'extra')}  Available at Extra Cost",
                size_pt=11, color=DARK_GRAY)
        hp = cell.add_paragraph()
        add_paragraph_spacing(hp, before=0, after=56)
        add_run(hp, "Hosting Management After Delivery:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        add_run(hp, f"{cb(hosting_aft == 'codesino')}  Managed by Codesino      "
                    f"{cb(hosting_aft == 'client')}  Transferred to Client      "
                    f"{cb(hosting_aft == 'tbd')}  TBD",
                size_pt=11, color=DARK_GRAY)
        mp = cell.add_paragraph()
        no_space(mp)
        add_run(mp, "Ongoing Maintenance Retainer:  ", bold=True, size_pt=11, color=BRAND_BLUE)
        if maintenance == "agreed":
            add_run(mp, f"☑  Agreed — {currency_sym}{maint_fee} / month      ☐  Not Agreed",
                    size_pt=11, color=DARK_GRAY)
        else:
            add_run(mp, "☐  Agreed      ☑  Not Agreed", size_pt=11, color=DARK_GRAY)

    section_box(doc, support_fn, HEX_WHITE)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 12 — INTELLECTUAL PROPERTY
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "12.  INTELLECTUAL PROPERTY")
    for item in [
        "Upon receipt of full and final payment, all intellectual property rights for the final deliverables transfer entirely to the client.",
        "Codesino retains the right to display the completed project in its portfolio and marketing materials unless the client requests otherwise in writing prior to project completion.",
        "Any third-party libraries, plugins, frameworks, or open-source software used remain subject to their own respective licences.",
        "Codesino retains full ownership of all custom code, internal tools, and reusable components used in the project until final payment is received and confirmed.",
        "The client warrants that all content, images, logos, and materials provided to Codesino are legally owned or licenced by the client and do not infringe any third-party intellectual property rights.",
    ]:
        add_numbered(doc, item)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 13 — CONFIDENTIALITY
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "13.  CONFIDENTIALITY")
    body_text(doc,
              "Both parties agree to keep strictly confidential any sensitive business information, "
              "trade secrets, proprietary data, client lists, pricing, or technical specifications "
              "shared during the course of this project. This obligation remains in full effect for "
              "two (2) years after the completion or termination of this agreement.")
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 14 — TERMINATION
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "14.  TERMINATION")
    for item in [
        "Either party may terminate this agreement with 7 (seven) calendar days' written notice.",
        "If the client terminates the agreement after development has commenced, all payments made are non-refundable.",
        "If Codesino terminates the agreement without cause, a pro-rated refund will be issued for work not yet started.",
        "Termination must be communicated via email or documented written message to be legally valid.",
        "Upon termination, each party shall immediately return or destroy any confidential materials belonging to the other party.",
    ]:
        add_numbered(doc, item)
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 15 — DISPUTE RESOLUTION
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "15.  DISPUTE RESOLUTION")
    body_text(doc,
              "In the event of a dispute, both parties agree to first attempt resolution through "
              "good-faith negotiation within 14 (fourteen) calendar days of the dispute being raised "
              "in writing. If resolution cannot be reached, both parties agree to submit the matter "
              "to mediation before resorting to formal legal proceedings. This agreement shall be "
              "governed by and construed in accordance with the laws of the Federal Republic of Nigeria.")
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 16 — ADDITIONAL NOTES
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "16.  ADDITIONAL NOTES & SPECIAL CONDITIONS")
    notes = data.get("additionalNotes", "")
    if notes:
        body_text(doc, notes)
    else:
        italic_note(doc, "No additional notes for this agreement.")
    divider(doc)

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION 17 — SIGNATURES
    # ═════════════════════════════════════════════════════════════════════════
    heading1(doc, "17.  AGREEMENT & SIGNATURES")
    body_text(doc,
              "By signing below, both parties confirm that they have read, understood, and agreed "
              "to all terms and conditions outlined in this Client Project Agreement. This document "
              "constitutes a legally binding agreement between Codesino Software Development "
              "Services and the client named herein.")
    spacer(doc, 80, 60)

    # ── Signature table: [Provider block] [gap] [Client block] ──────────────
    sig_table = doc.add_table(rows=2, cols=3)
    set_table_width(sig_table,    9360)
    set_column_widths(sig_table, [4400, 560, 4400])
    set_table_no_borders(sig_table)

    # Header row — coloured title bands
    for col_idx, label, align in [
        (0, "SERVICE PROVIDER",  WD_ALIGN_PARAGRAPH.CENTER),
        (2, "CLIENT",            WD_ALIGN_PARAGRAPH.CENTER),
    ]:
        cell = sig_table.cell(0, col_idx)
        set_cell_shading(cell,  HEX_BRAND_BLUE)
        set_cell_margins(cell,  top=100, bottom=100, left=150, right=150)
        set_cell_borders(cell,  top=None, bottom=None, left=None, right=None)
        clear_cell_paras(cell)
        p = cell.add_paragraph()
        no_space(p)
        p.alignment = align
        add_run(p, label, bold=True, size_pt=11, color=WHITE_COLOR)

    # Spacer column header
    gap_hdr = sig_table.cell(0, 1)
    set_cell_borders(gap_hdr, top=None, bottom=None, left=None, right=None)
    clear_cell_paras(gap_hdr)
    gap_hdr.add_paragraph()

    # Body row — signature content cells
    thin = {"val": "single", "sz": 2, "color": HEX_RULE}

    # Provider signature cell
    prov_cell = sig_table.cell(1, 0)
    set_cell_borders(prov_cell, top=None, bottom=thin, left=thin, right=thin)
    set_cell_margins(prov_cell, top=120, bottom=140, left=150, right=150)
    clear_cell_paras(prov_cell)

    pp1 = prov_cell.add_paragraph()
    add_paragraph_spacing(pp1, before=0, after=50)
    add_run(pp1, "Codesino Software Development Services",
            size_pt=10, color=DARK_GRAY)

    pp2 = prov_cell.add_paragraph()
    add_paragraph_spacing(pp2, before=0, after=80)
    add_run(pp2, "Authorised Signatory", size_pt=10, color=MID_GRAY)

    sig_line(prov_cell, "Digital Signature:")

    # ── Stamp image: placed below the sig line inside the provider cell ──────
    if os.path.exists(STAMP_PATH):
        stamp_p = prov_cell.add_paragraph()
        add_paragraph_spacing(stamp_p, before=20, after=20)
        stamp_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        stamp_p.add_run().add_picture(STAMP_PATH, width=Inches(1.3))

    pds  = data.get("providerSignDate", "")
    pp4  = prov_cell.add_paragraph()
    no_space(pp4)
    add_run(pp4, f"Date: {format_date(pds) if pds else '___ / ___ / ______'}",
            italic=True, size_pt=10, color=hex_to_rgb("999999"))

    # Spacer body cell
    gap_body = sig_table.cell(1, 1)
    set_cell_borders(gap_body, top=None, bottom=None, left=None, right=None)
    clear_cell_paras(gap_body)
    gap_body.add_paragraph()

    # Client signature cell
    cli_cell = sig_table.cell(1, 2)
    set_cell_borders(cli_cell, top=None, bottom=thin, left=thin, right=thin)
    set_cell_margins(cli_cell, top=120, bottom=140, left=150, right=150)
    clear_cell_paras(cli_cell)

    cc = data.get("clientCompany", "")
    cn = data.get("clientName",    "")

    cp1 = cli_cell.add_paragraph()
    add_paragraph_spacing(cp1, before=0, after=50)
    add_run(cp1, cc if cc else "____________________________",
            italic=not bool(cc), size_pt=10,
            color=DARK_GRAY if cc else hex_to_rgb("999999"))

    cp2 = cli_cell.add_paragraph()
    add_paragraph_spacing(cp2, before=0, after=80)
    add_run(cp2, f"Name: {cn}" if cn else "Name: ______________________",
            italic=not bool(cn), size_pt=10,
            color=DARK_GRAY if cn else hex_to_rgb("999999"))

    sig_line(cli_cell, "Signature:")

    # Blank space where client stamps/signs (no stamp image for client)
    blank_p = cli_cell.add_paragraph()
    add_paragraph_spacing(blank_p, before=20, after=20)
    no_space(blank_p)

    cds  = data.get("clientSignDate", "")
    cp4  = cli_cell.add_paragraph()
    no_space(cp4)
    add_run(cp4, f"Date: {format_date(cds) if cds else '___ / ___ / ______'}",
            italic=True, size_pt=10, color=hex_to_rgb("999999"))

    spacer(doc, 100, 60)

    # ── Document footer text ─────────────────────────────────────────────────
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    no_space(fp)
    add_run(fp,
            "This agreement was prepared by Codesino Software Development Services  |  "
            "www.codesinodev.com  |  contact@codesinodev.com",
            italic=True, size_pt=9, color=hex_to_rgb("9E9E9E"))

    # ── Page-level running footer ────────────────────────────────────────────
    section = doc.sections[0]
    footer  = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(footer_para,
            "Codesino Software Development Services  |  Client Project Agreement  |  CONFIDENTIAL",
            size_pt=8, color=hex_to_rgb("999999"))

    # ── Serialize to bytes ───────────────────────────────────────────────────
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()