#!/usr/bin/env python3
"""Generate a Word CV that reproduces the attached two-column template."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# ---- Palette ----------------------------------------------------------------
NAVY = RGBColor(0x14, 0x32, 0x55)        # dark sidebar / headings
NAVY_HEX = "143255"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE = RGBColor(0x2F, 0x5C, 0x99)  # accent text on white
BAR_FILL = "FFFFFF"                       # filled portion of language bars
BAR_EMPTY = "3C5C86"                      # empty portion of language bars

# Heading font echoes the geometric look of the template; body uses Calibri.
# If the heading font is not installed on the reader's machine, Word will
# substitute a close sans-serif automatically.
HEAD_FONT = "Montserrat"
FONT = "Calibri"


# ---- Low-level helpers ------------------------------------------------------
def set_cell_bg(cell, hex_color):
    """Set table-cell background shading."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    """Margins in twips (1/1440 inch)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_width(cell, inches):
    cell.width = Inches(inches)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(inches * 1440)))
    tcW.set(qn("w:type"), "dxa")


def add_run(paragraph, text, *, size=11, bold=False, italic=False,
            color=NAVY, font=FONT, caps=False):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    if caps:
        rPr = run._element.get_or_add_rPr()
        caps_el = OxmlElement("w:caps")
        caps_el.set(qn("w:val"), "true")
        rPr.append(caps_el)
    # ensure east-asian/complex script also use the font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    return run


def space(paragraph, before=0, after=0, line=None):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line


def progress_bar(cell, filled_ratio):
    """Draw a simple language-level bar using a 1-row nested table."""
    tbl = cell.add_table(rows=1, cols=2)
    tbl.autofit = False
    total = 1.9  # inches
    filled = round(total * filled_ratio, 2)
    empty = round(total - filled, 2)
    left, right = tbl.rows[0].cells
    set_cell_width(left, max(filled, 0.05))
    set_cell_width(right, max(empty, 0.05))
    set_cell_bg(left, BAR_FILL)
    set_cell_bg(right, BAR_EMPTY)
    for c in (left, right):
        set_cell_margins(c, top=10, bottom=10)
        c.paragraphs[0].text = ""
        c.paragraphs[0].add_run(" ").font.size = Pt(3)
    return tbl


# ---- Document ---------------------------------------------------------------
doc = Document()

# Page setup: A4, minimal margins so the sidebar reaches the edges.
section = doc.sections[0]
section.page_height = Inches(11.69)
section.page_width = Inches(8.27)
for attr in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
    setattr(section, attr, Inches(0))

# Outer two-column table -----------------------------------------------------
outer = doc.add_table(rows=1, cols=2)
outer.alignment = WD_TABLE_ALIGNMENT.LEFT
outer.allow_autofit = False

sidebar = outer.rows[0].cells[0]
main = outer.rows[0].cells[1]

set_cell_width(sidebar, 2.85)
set_cell_width(main, 5.42)
set_cell_bg(sidebar, NAVY_HEX)
set_cell_bg(main, "FFFFFF")
set_cell_margins(sidebar, top=360, start=360, bottom=360, end=300)
set_cell_margins(main, top=520, start=420, bottom=360, end=420)


# ===== SIDEBAR ===============================================================
def sidebar_heading(text):
    p = sidebar.add_paragraph()
    space(p, before=16, after=8)
    add_run(p, text, size=15, bold=True, color=WHITE, caps=True, font=HEAD_FONT)


def sidebar_text(text, *, size=11, after=2, bold=False):
    p = sidebar.add_paragraph()
    space(p, before=0, after=after, line=1.05)
    add_run(p, text, size=size, color=WHITE, bold=bold)
    return p


# Photo: a clean circular avatar generated by make_avatar.py. A real
# photograph is intentionally not used (none was provided). ------------------
photo_p = sidebar.paragraphs[0]
photo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
space(photo_p, before=0, after=10)
import os
for candidate in ("photo.png", "avatar.png"):
    if os.path.exists(candidate):
        photo_p.add_run().add_picture(candidate, width=Inches(1.85))
        break
else:
    add_run(photo_p, "\n\n[ PHOTO ]\n\n", size=11, bold=True, color=WHITE)

# COORDONNÉES
sidebar_heading("Coordonnées")
sidebar_text("Tél : 123-456-7890", after=6)
sidebar_text("hello@reallygreatsite.com", after=6)
sidebar_text("123 Anywhere St., Any City", after=6)

# LANGUES
sidebar_heading("Langues")
p = sidebar.add_paragraph()
space(p, before=2, after=1)
add_run(p, "Anglais", size=11, color=WHITE)
progress_bar(sidebar, 0.85)
p = sidebar.add_paragraph()
space(p, before=4, after=1)
add_run(p, "Allemand", size=11, color=WHITE)
progress_bar(sidebar, 0.6)

# COMPÉTENCES
sidebar_heading("Compétences")
for comp in [
    "Gestion du temps",
    "Capacités d'organisation",
    "Communication",
    "Leadership",
    "Logiciels de gestion de projet",
    "Gestion de budget",
]:
    sidebar_text(comp, after=4)

# CENTRES D'INTÉRÊT
sidebar_heading("Centres d'intérêt")
for hobby in ["Lecture", "Randonnée", "Gymnastique"]:
    sidebar_text(hobby, after=4)


# ===== MAIN COLUMN ===========================================================
# Name + title
p = main.paragraphs[0]
space(p, before=0, after=0)
add_run(p, "SACHA DUBOIS", size=34, bold=True, color=NAVY, font=HEAD_FONT)

p = main.add_paragraph()
space(p, before=2, after=18)
add_run(p, "CHARGÉE DE PROJET", size=16, color=LIGHT_BLUE, caps=True, font=HEAD_FONT)


def section_title(text):
    p = main.add_paragraph()
    space(p, before=10, after=10)
    add_run(p, text, size=18, bold=True, color=NAVY, caps=True, font=HEAD_FONT)


def entry(title, org, date, *, org_italic=True):
    # Title + date on one line (date right-aligned via tab)
    p = main.add_paragraph()
    space(p, before=8, after=0, line=1.0)
    # right tab stop
    p.paragraph_format.tab_stops.add_tab_stop(Inches(5.0), WD_TAB_ALIGNMENT.RIGHT)
    add_run(p, title, size=12.5, bold=True, color=NAVY, font=HEAD_FONT)
    add_run(p, "\t" + date, size=10, italic=True, color=LIGHT_BLUE)
    # Org line
    p2 = main.add_paragraph()
    space(p2, before=0, after=2, line=1.0)
    add_run(p2, org, size=11.5, italic=org_italic, color=LIGHT_BLUE)


def bullet(text):
    p = main.add_paragraph()
    space(p, before=0, after=2, line=1.0)
    pf = p.paragraph_format
    pf.left_indent = Inches(0.28)
    pf.first_line_indent = Inches(-0.18)
    add_run(p, "•  ", size=11, color=LIGHT_BLUE)
    add_run(p, text, size=11, color=LIGHT_BLUE)


# FORMATION
section_title("Formation")
entry("Master en Management", "Really Great School - Any City", "2016 - 2018",
      org_italic=False)
entry("Licence en Gestion Administrative", "Really Great School - Any City",
      "2013 - 2016", org_italic=False)

# EXPÉRIENCE PROFESSIONNELLE
section_title("Expérience Professionnelle")

entry("Chargée de Projet", "Really Great Company - Any City",
      "Depuis janvier 2020")
bullet("Elaboration du plan de projet et gestion des calendriers")
bullet("Gestion des ressources et du budget")
bullet("Gestion des risques (identification des risques, réaction aux imprévus)")

entry("Assistante Chargée de Projet", "Really Great Company - Any City",
      "Septembre 2018 - Janvier 2020")
bullet("Rédaction des rapports de progression")
bullet("Participation aux réunions de suivi de projet")
bullet("Suivi de l'avancement des projets en collaboration avec les équipes concernées")

entry("Stagiaire Chargée de Projet", "Really Great Company - Any City",
      "Février 2018 - Juin 2018")
bullet("Aide pour définir les objectifs et les périmètres du projet")
bullet("Participation à l'élaboration des plans de projet (échéanciers, jalons et livrables)")


# Remove default table borders ------------------------------------------------
def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tblPr.append(borders)


no_borders(outer)
for t in main.tables:
    no_borders(t)
for t in sidebar.tables:
    no_borders(t)

doc.save("Sacha_Dubois_CV.docx")
print("Saved Sacha_Dubois_CV.docx")
