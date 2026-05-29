"""
Generate TIP_Architecture_Overview.docx
Rolls-Royce Digital Team — Test & Inspection Plan Application
Architecture & Technical Overview
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Colour palette ──────────────────────────────────────────────────────────
NAVY       = RGBColor(0x1F, 0x38, 0x64)   # #1F3864
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
AMBER      = RGBColor(0xFF, 0xD9, 0x66)   # #FFD966
LIGHT_GREY = RGBColor(0xF2, 0xF2, 0xF2)   # #F2F2F2  (code block bg)
MID_GREY   = RGBColor(0xD9, 0xD9, 0xD9)   # #D9D9D9  (table border / alt row)


# ── Low-level XML helpers ────────────────────────────────────────────────────

def rgb_hex(rgb: RGBColor) -> str:
    """Return 6-char uppercase hex string from an RGBColor (tuple-like)."""
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def set_cell_bg(cell, rgb: RGBColor):
    """Fill a table cell background with a solid colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    hex_col = rgb_hex(rgb)
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_col)
    tcPr.append(shd)


def set_para_bg(paragraph, rgb: RGBColor):
    """Fill a paragraph background (shading) — used for code blocks."""
    pPr  = paragraph._p.get_or_add_pPr()
    shd  = OxmlElement("w:shd")
    hex_col = rgb_hex(rgb)
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_col)
    pPr.append(shd)


def set_para_border(paragraph, rgb: RGBColor, side="left", sz=24, space=4):
    """Add a coloured border to a paragraph (left bar for code blocks)."""
    pPr  = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bdr  = OxmlElement(f"w:{side}")
    hex_col = rgb_hex(rgb)
    bdr.set(qn("w:val"),   "single")
    bdr.set(qn("w:sz"),    str(sz))
    bdr.set(qn("w:space"), str(space))
    bdr.set(qn("w:color"), hex_col)
    pBdr.append(bdr)
    pPr.append(pBdr)


def set_table_borders(table, color: str = "D9D9D9"):
    """Apply thin borders to every cell of a table."""
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tblBorders.append(el)
    tblPr.append(tblBorders)


def set_col_widths(table, widths_cm):
    """Set fixed column widths (in cm) for a table."""
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths_cm):
                cell.width = Cm(widths_cm[idx])


def page_break(doc):
    doc.add_page_break()


# ── Font / style helpers ────────────────────────────────────────────────────

def style_run(run, bold=False, italic=False, size=None, color=None, font="Calibri"):
    run.font.name  = font
    run.font.bold  = bold
    run.font.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def add_heading(doc, text, level, space_before=12, space_after=4):
    """Add a heading paragraph using built-in Heading styles."""
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(space_before)
    h.paragraph_format.space_after  = Pt(space_after)
    return h


def add_body(doc, text, space_after=4):
    p = doc.add_paragraph(text, style="Normal")
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)
    return p


def add_bullet(doc, text, level=0, bold_prefix=None):
    """Add a List Bullet paragraph; optional bold prefix text."""
    style = "List Bullet" if level == 0 else "List Bullet 2"
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    if bold_prefix:
        r = p.add_run(bold_prefix)
        style_run(r, bold=True, size=11)
        r2 = p.add_run(text)
        style_run(r2, size=11)
    else:
        r = p.add_run(text)
        style_run(r, size=11)
    return p


def add_code_block(doc, lines, space_before=4, space_after=4):
    """Add a shaded, monospace code block (one paragraph per line)."""
    GREY = RGBColor(0xF4, 0xF4, 0xF4)
    ACCENT = RGBColor(0x1F, 0x38, 0x64)  # navy left bar

    for i, line in enumerate(lines):
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.space_before = Pt(space_before if i == 0 else 0)
        p.paragraph_format.space_after  = Pt(space_after  if i == len(lines) - 1 else 0)
        p.paragraph_format.left_indent  = Cm(0.3)
        p.paragraph_format.right_indent = Cm(0.3)
        set_para_bg(p, GREY)
        if i == 0:
            set_para_border(p, ACCENT, side="left", sz=20, space=4)
        run = p.add_run(line if line else " ")
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    return p


# ── Table builder ───────────────────────────────────────────────────────────

def add_data_table(doc, headers, rows, col_widths_cm=None):
    """
    Create a formatted table with:
      - Navy header row with white bold text
      - Alternating white / very-light-grey data rows
      - Thin grey borders throughout
    """
    n_cols = len(headers)
    table  = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table, color="BFBFBF")

    # ── Header row
    hdr_row = table.rows[0]
    for idx, hdr_text in enumerate(headers):
        cell = hdr_row.cells[idx]
        set_cell_bg(cell, NAVY)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        run = p.add_run(hdr_text)
        style_run(run, bold=True, size=10, color=WHITE)

    # ── Data rows
    ALT_ROW = RGBColor(0xF7, 0xF9, 0xFC)   # very pale blue-grey
    for r_idx, row_data in enumerate(rows):
        row_obj = table.rows[r_idx + 1]
        bg = ALT_ROW if r_idx % 2 == 1 else WHITE
        for c_idx, cell_text in enumerate(row_data):
            cell = row_obj.cells[c_idx]
            set_cell_bg(cell, bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            run = p.add_run(str(cell_text))
            # Bold the first column for key–value style tables
            style_run(run, bold=(c_idx == 0 and len(headers) <= 2), size=10)

    if col_widths_cm:
        set_col_widths(table, col_widths_cm)

    return table


# ── Document metadata ────────────────────────────────────────────────────────

def set_doc_properties(doc):
    core = doc.core_properties
    core.title   = "Test & Inspection Plan Application — Architecture & Technical Overview"
    core.author  = "Rolls-Royce Digital Team"
    core.subject = "TIP Application — Technical Documentation"
    core.keywords = "TIP, Test, Inspection, Plan, Architecture, Python"
    core.revision = 7


# ── Page setup ───────────────────────────────────────────────────────────────

def set_page_margins(doc):
    for section in doc.sections:
        section.page_width   = Cm(21.0)
        section.page_height  = Cm(29.7)
        section.left_margin  = Cm(2.54)
        section.right_margin = Cm(2.54)
        section.top_margin   = Cm(2.54)
        section.bottom_margin = Cm(2.0)


# ── Style tweaks ─────────────────────────────────────────────────────────────

def configure_styles(doc):
    styles = doc.styles

    # Normal body text
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)

    # Heading 1 → navy, 16pt
    h1 = styles["Heading 1"]
    h1.font.name  = "Calibri"
    h1.font.size  = Pt(14)
    h1.font.bold  = True
    h1.font.color.rgb = NAVY
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after  = Pt(4)
    h1.paragraph_format.keep_with_next = True

    # Heading 2 → slightly smaller navy
    h2 = styles["Heading 2"]
    h2.font.name  = "Calibri"
    h2.font.size  = Pt(12)
    h2.font.bold  = True
    h2.font.color.rgb = NAVY
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after  = Pt(3)
    h2.paragraph_format.keep_with_next = True

    # Heading 3 → same navy, 11pt
    h3 = styles["Heading 3"]
    h3.font.name  = "Calibri"
    h3.font.size  = Pt(11)
    h3.font.bold  = True
    h3.font.color.rgb = NAVY
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after  = Pt(2)
    h3.paragraph_format.keep_with_next = True


# ── Horizontal rule ──────────────────────────────────────────────────────────

def add_hr(doc, color: str = "1F3864", sz: int = 6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    str(sz))
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pBdr.append(bot)
    pPr.append(pBdr)
    return p


# ════════════════════════════════════════════════════════════════════════════
#  MAIN DOCUMENT BUILD
# ════════════════════════════════════════════════════════════════════════════

def build_document():
    doc = Document()
    set_page_margins(doc)
    configure_styles(doc)
    set_doc_properties(doc)

    # ── COVER / TITLE BLOCK ─────────────────────────────────────────────────
    # Spacer
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(18)

    # Main title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after  = Pt(6)
    title_p.paragraph_format.space_before = Pt(0)
    r = title_p.add_run("Test & Inspection Plan Application")
    r.font.name  = "Calibri"
    r.font.size  = Pt(22)
    r.font.bold  = True
    r.font.color.rgb = NAVY

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_p.paragraph_format.space_after = Pt(20)
    r2 = subtitle_p.add_run("Architecture & Technical Overview")
    r2.font.name  = "Calibri"
    r2.font.size  = Pt(16)
    r2.font.bold  = False
    r2.font.color.rgb = RGBColor(0x44, 0x47, 0x4F)

    add_hr(doc, color="1F3864", sz=12)

    # Document metadata table
    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, color="D9D9D9")

    meta_data = [
        ("Prepared for", "Rolls-Royce Digital Team"),
        ("Version",      "Rev 7"),
        ("Date",         "29 May 2026"),
    ]
    for r_idx, (label, value) in enumerate(meta_data):
        row = meta_table.rows[r_idx]
        # Label cell
        c0 = row.cells[0]
        set_cell_bg(c0, RGBColor(0xEE, 0xF2, 0xF8))
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(4)
        p0.paragraph_format.space_after  = Pt(4)
        run0 = p0.add_run(label)
        style_run(run0, bold=True, size=11, color=NAVY)
        # Value cell
        c1 = row.cells[1]
        set_cell_bg(c1, WHITE)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(4)
        p1.paragraph_format.space_after  = Pt(4)
        run1 = p1.add_run(value)
        style_run(run1, size=11)

    set_col_widths(meta_table, [4.5, 11.0])

    add_hr(doc, color="1F3864", sz=6)

    spacer2 = doc.add_paragraph()
    spacer2.paragraph_format.space_after = Pt(12)

    page_break(doc)

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 1 — Purpose & Intent
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "1.  Purpose & Intent", 1)

    add_body(doc,
        "The Test & Inspection Plan (TIP) application is a standalone Windows desktop tool designed "
        "to digitalise and streamline the creation, review, and approval of Test & Inspection Plans "
        "within a manufacturing quality assurance workflow.")

    add_body(doc,
        "The application replaces a legacy Excel/VBA-based process, providing a structured two-role "
        "workflow that separates the responsibilities of the Plan Author from the TAR (Technical "
        "Authority Representative) Holder, eliminating the risk of unintended data crossover between roles.")

    add_heading(doc, "Key Objectives", 2)

    add_bullet(doc, "Guided, step-by-step plan creation for the Author")
    add_bullet(doc,
        "Clear separation: the Author never sees TAR assessment fields; the TAR Holder never sees "
        "Author editing controls")
    add_bullet(doc, "Structured handoff via a portable .tip file (JSON format)")
    add_bullet(doc,
        "PDF export with TAR Holder identity details and a placeholder for a digital certificate signature")
    add_bullet(doc, "Outlook email integration for both the Author-to-TAR and TAR-to-Author handoff steps")

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 2 — Technology Stack
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "2.  Technology Stack", 1)

    tech_headers = ["Component", "Technology", "Notes"]
    tech_rows = [
        ("Language",           "Python 3.12",              ""),
        ("GUI Framework",      "customtkinter 5.2+",       "Modern themed tkinter wrapper"),
        ("Excel Import",       "openpyxl 3.1+",            "NetInspect file parsing"),
        ("PDF Generation",     "ReportLab 4.0+",           "A4 layout, colour-matched to Excel original"),
        ("Email Integration",  "pywin32 (win32com)",       "Windows Outlook COM automation"),
        ("Distribution",       "PyInstaller",              "Single-file .exe, no Python install required"),
        ("File Format",        "JSON (.tip)",              "Human-readable, version-tagged plan files"),
    ]
    t = add_data_table(doc, tech_headers, tech_rows, col_widths_cm=[4.5, 5.0, 6.0])

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 3 — Application Architecture
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "3.  Application Architecture", 1)

    add_body(doc,
        "The application uses a single-window shell (TIPApp) that swaps between three role-based "
        "views. There are no tabs visible across roles — each role sees only the interface relevant to them.")

    # 3.1 View Hierarchy
    add_heading(doc, "3.1  View Hierarchy", 2)

    add_body(doc,
        "TIPApp is the main window shell. It manages the header, status bar, and content area; "
        "routes between views based on user action; and houses the Home button and role mode badge.")

    for line in [
        "LaunchView — the home screen",
        "  Entry point on application start",
        "  Two paths: Create New Plan (→ AuthorView) or Open Existing Plan (→ role selection dialog → appropriate view)",
    ]:
        stripped = line.lstrip()
        indent = len(line) - len(stripped)
        lvl = 1 if indent >= 2 else 0
        add_bullet(doc, stripped, level=lvl)

    doc.add_paragraph()

    for line in [
        "AuthorView — 3-step guided flow",
        "  Step 1: Plan Details",
        "  Step 2: Inspection Cards",
        "  Step 3: Send to TAR Holder",
    ]:
        stripped = line.lstrip()
        lvl = 1 if line.startswith("  ") else 0
        add_bullet(doc, stripped, level=lvl)

    doc.add_paragraph()

    for line in [
        "TARHolderView — TAR review and approval interface",
        "  Instruction banner",
        "  Read-only plan info",
        "  TAR Holder identity details (name, date, employer number)",
        "  Card list + TAR assessment fields panel",
        "  Save / Export / Email actions",
    ]:
        stripped = line.lstrip()
        lvl = 1 if line.startswith("  ") else 0
        add_bullet(doc, stripped, level=lvl)

    # 3.2 Role Routing
    add_heading(doc, "3.2  Role Routing", 2)

    add_body(doc,
        "When a user clicks \"Open Existing Plan\", a RoleDialog asks whether they are opening as "
        "Author or TAR Holder. This determines which view is loaded. There is no authentication — "
        "the separation is enforced by the UI itself and the distinct views each role sees.")

    # 3.3 Data Flow
    add_heading(doc, "3.3  Data Flow", 2)

    data_flow_lines = [
        "Author creates plan (AuthorView Steps 1–2)",
        "        ↓",
        "Author sends .tip file + Draft PDF via Outlook (Step 3)",
        "        ↓",
        "TAR Holder opens .tip → selects TAR Holder role → TARHolderView",
        "        ↓",
        "TAR Holder enters identity details + completes TAR fields per card",
        "        ↓",
        "TAR Holder exports Approved PDF (with identity in signature block)",
        "        ↓",
        "TAR Holder applies digital certificate externally (Adobe / equivalent)",
        "        ↓",
        "TAR Holder emails signed PDF back to Author via app (Outlook draft)",
    ]
    add_code_block(doc, data_flow_lines)

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 4 — File Format
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "4.  File Format — .tip", 1)

    add_body(doc,
        "Plans are stored as JSON files with the .tip extension. The format is human-readable and "
        "version-tagged for future compatibility.")

    add_heading(doc, "Schema (version 1.1)", 2)

    schema_lines = [
        '{',
        '  "version": "1.1",',
        '  "plan": {',
        '    "mva_author":  "string",',
        '    "mva_number":  "string",',
        '    "date":        "DD/MM/YYYY",',
        '    "mode":        "manual | automated",',
        '    "tar_email":   "string  (stored when Author sends)"',
        '  },',
        '  "tar_holder": {',
        '    "name":        "string",',
        '    "date":        "DD/MM/YYYY",',
        '    "employer_no": "string"',
        '  },',
        '  "cards": [',
        '    {',
        '      "item_num":            1,',
        '      "op_no":               "string",',
        '      "sampling":            "string",',
        '      "meas_type":           "string",',
        '      "spec":                "string",',
        '      "units":               "string",',
        '      "bubble":              "string",',
        '      "equipment_name":      "string",',
        '      "equipment_no":        "string",',
        '      "description":         "string",',
        '      "ref_loc":             "string",',
        '      "process_ref":         "string",',
        '      "verification_method": "string",',
        '      "equipment_type":      "string",',
        '      "msa_method":          "string",',
        '      "msa_ref":             "string",',
        '      "responsible_person":  "string",',
        '      "control_method":      "string",',
        '      "reaction_plan":       "string",',
        '      "comments":            "string"',
        '    }',
        '  ]',
        '}',
    ]
    add_code_block(doc, schema_lines)

    add_body(doc,
        "Fields item_num through process_ref are Author fields. Fields verification_method through "
        "comments are TAR Holder assessment fields. The tar_holder block is optional — absent in "
        "draft plans not yet reviewed.")

    add_body(doc,
        "Backwards compatibility: version 1.0 files (no tar_holder block) load correctly; missing "
        "fields default gracefully.")

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 5 — Author Workflow
    # ════════════════════════════════════════════════════════════════════════
    page_break(doc)
    add_heading(doc, "5.  Author Workflow — Detailed", 1)

    add_heading(doc, "Step 1: Plan Details", 2)

    add_body(doc, "The Author enters:")
    for item in [
        "MVA Author name (Prepared By)",
        "MVA Number",
        "Date (free text DD/MM/YYYY, with Today shortcut)",
        "Input Method: Manual or Automated (NetInspect import)",
    ]:
        add_bullet(doc, item)

    add_body(doc,
        "TAR Holder email is deliberately absent from this step — it is only requested at the "
        "point of sending (Step 3), keeping setup focused.")

    add_heading(doc, "Step 2: Inspection Cards", 2)

    add_bullet(doc,
        "Manual mode: ",
        bold_prefix="Manual mode: ")
    # Override last bullet with proper text
    doc.paragraphs[-1].clear()
    p_manual = doc.add_paragraph(style="List Bullet")
    p_manual.paragraph_format.space_after  = Pt(4)
    p_manual.paragraph_format.space_before = Pt(2)
    r_b = p_manual.add_run("Manual mode:  ")
    style_run(r_b, bold=True, size=11)
    r_t = p_manual.add_run(
        "The Author adds cards one at a time via a dialog. Each card captures: Op No., Sampling, "
        "Measurement Type, Spec, Units, Bubble #, Equipment Name, Equipment No., Description, "
        "Reference Location, Process Reference. TAR assessment fields are completely absent from this dialog.")
    style_run(r_t, size=11)

    p_auto = doc.add_paragraph(style="List Bullet")
    p_auto.paragraph_format.space_after  = Pt(4)
    p_auto.paragraph_format.space_before = Pt(2)
    r_b2 = p_auto.add_run("Automated mode:  ")
    style_run(r_b2, bold=True, size=11)
    r_t2 = p_auto.add_run(
        "The Author imports a NetInspect export (.xlsx/.xls/.xlsm). The importer maps standard "
        "NetInspect columns (Op, Sampling, Description, Type, Spec, Units, Bubble, Ref, Process Ref) "
        "and auto-numbers items. Header validation confirms the file is a valid NetInspect export "
        "before importing.")
    style_run(r_t2, size=11)

    add_body(doc,
        "Cards are displayed in a sortable treeview. Double-click to edit. "
        "Delete and Clear All include confirmation dialogs.")

    add_heading(doc, "Step 3: Send to TAR Holder", 2)

    for item in [
        "Plan summary displayed (read-only confirmation)",
        "TAR Holder email entered here (stored to plan on send)",
        "\"Send Email via Outlook\" — creates a draft Outlook email with the .tip file and Draft PDF "
        "attached, with structured ACTION REQUIRED instructions for the TAR Holder",
        "\"Export Draft PDF\" — saves a PDF for reference; TAR assessment fields shown as blank in the PDF",
        "\"Save Plan (.tip)\" — saves current state to file",
    ]:
        add_bullet(doc, item)

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 6 — TAR Holder Workflow
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "6.  TAR Holder Workflow — Detailed", 1)

    add_heading(doc, "Entry", 2)
    add_body(doc,
        "The TAR Holder opens the .tip file in the application, selects \"TAR Holder\" in the role "
        "dialog, and is routed directly to TARHolderView. A green instruction banner at the top of "
        "the screen explains exactly what to do.")

    add_heading(doc, "Identity Details", 2)
    add_body(doc, "Before any export is permitted, the TAR Holder must complete:")
    for item in ["TAR Holder Name", "Date (with Today shortcut)", "Employer Number"]:
        add_bullet(doc, item)
    add_body(doc,
        "Validation fires on export/email attempts — missing fields block the action with a clear "
        "message listing what is incomplete.")

    add_heading(doc, "Card Review", 2)
    add_body(doc,
        "A treeview lists all inspection items. Selecting a card loads its TAR assessment panel "
        "(green background) showing 8 fields:")
    for item in [
        "Verification Method (dropdown — 8 options)",
        "Equipment Type (dropdown — 79 options)",
        "MSA Verification Method (dropdown — 11 options)",
        "MSA Ref / Justification (free text)",
        "Responsible Person (free text)",
        "Control Method (free text)",
        "Reaction Plan (free text)",
        "Comments (free text)",
    ]:
        add_bullet(doc, item)
    add_body(doc,
        "Fields auto-flush to the card data on selection change. No manual Save per card is required.")

    add_heading(doc, "Save Progress", 2)
    add_body(doc,
        '"Save Progress" writes the current state (including all completed TAR fields and identity '
        "details) back to the original .tip file. This allows the TAR Holder to complete the review "
        "across multiple sessions.")

    add_heading(doc, "Export Approved PDF", 2)
    add_body(doc, "Generates a PDF including:")
    for item in [
        "Plan header (MVA Number, Author, Date, Input Method)",
        "All inspection cards with full Author data and completed TAR assessment fields",
        "Authorisation & Approval block populated with TAR Holder name, date, employer number, and a "
        "note indicating a digital certificate is attached",
    ]:
        add_bullet(doc, item)
    add_body(doc,
        "The TAR Holder then applies their digital certificate signature externally "
        "(e.g. Adobe Acrobat) and saves the signed PDF.")

    add_heading(doc, "Email Back to Author", 2)
    add_body(doc,
        "The TAR Holder selects the externally-signed PDF via a file picker. The application creates "
        "an Outlook draft addressed to the Author (email stored in the plan from when Author sent; "
        "prompted if absent), with the signed PDF attached and approval details in the body.")

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 7 — PDF Generation
    # ════════════════════════════════════════════════════════════════════════
    page_break(doc)
    add_heading(doc, "7.  PDF Generation", 1)

    add_body(doc,
        "PDFs are generated using ReportLab. The layout uses nested Table elements to reproduce "
        "the visual style of the original Excel document.")

    add_heading(doc, "Colour Coding", 2)

    colour_headers = ["Colour", "Hex", "Usage"]
    colour_rows = [
        ("Navy",         "#1F3864", "Headers and primary elements"),
        ("Amber",        "#FFD966", "Card field label rows"),
        ("Sky blue",     "#DEEAF6", "Reference rows and plan info"),
        ("Light green",  "#EDF7EE", "TAR assessment sections"),
        ("Dark green",   "#1E4620", "TAR section headers"),
        ("Cream",        "#FFFDE7", "Description text areas"),
    ]
    add_data_table(doc, colour_headers, colour_rows, col_widths_cm=[3.5, 3.0, 9.0])

    add_body(doc,
        "Two card layouts are supported: Manual (full field grid) and Automated (header + description "
        "only). Both modes include the TAR assessment block.")

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 8 — Key Classes & Functions
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "8.  Key Classes & Functions", 1)

    class_headers = ["Name", "Type", "Description"]
    class_rows = [
        ("TIPApp",             "Class",    "Main window shell; manages view switching, header, status bar"),
        ("LaunchView",         "Class",    "Home screen with Create New / Open Existing paths"),
        ("AuthorView",         "Class",    "3-step Author flow; owns plan/cards state during authoring"),
        ("TARHolderView",      "Class",    "TAR review interface; owns cards and tar_holder state"),
        ("CardDialog",         "Class",    "Author card add/edit dialog (no TAR fields)"),
        ("RoleDialog",         "Class",    "Role selection when opening an existing plan"),
        ("export_pdf()",       "Function", "Generates A4 PDF from plan, cards, and optional tar_holder data"),
        ("save_plan()",        "Function", "Serialises plan, cards, tar_holder to JSON .tip file"),
        ("load_plan()",        "Function", "Deserialises .tip file; returns plan, cards, tar_holder tuple"),
        ("import_netinspect()","Function", "Parses NetInspect .xlsx export into card list"),
        ("send_outlook()",     "Function", "Creates Outlook draft via COM automation (Windows only)"),
    ]
    add_data_table(doc, class_headers, class_rows, col_widths_cm=[4.5, 2.8, 8.2])

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 9 — Distribution
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "9.  Distribution", 1)

    add_body(doc,
        "The application is distributed as a single Windows executable compiled with PyInstaller. "
        "The TestInspectionPlan.spec file defines the build configuration, bundling the customtkinter "
        "theme assets.")

    add_body(doc,
        "No Python installation is required on end-user machines. The executable is self-contained.")

    add_heading(doc, "Build Command", 2)
    add_code_block(doc, ["pyinstaller TestInspectionPlan.spec"])

    add_heading(doc, "Output", 2)
    add_code_block(doc, [r"dist\TestInspectionPlan.exe"])

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 10 — Future Considerations
    # ════════════════════════════════════════════════════════════════════════
    add_heading(doc, "10.  Future Considerations", 1)

    add_body(doc,
        "The following areas are noted for potential future development:")

    future_items = [
        ("Authentication / role enforcement",
         " — currently role selection is on trust; a lightweight login or file-based role lock could be added"),
        ("In-app digital signature",
         " — currently delegated to external tooling; a signature pad widget or certificate integration could be embedded"),
        ("Plan status tracking",
         " — a status field (draft / sent / approved) could be added to the .tip schema to enable workflow state visibility"),
        ("Network / shared drive support",
         " — .tip files could be stored on a shared location with concurrent access controls"),
        ("Database backend",
         " — for larger deployments, migrating from JSON files to a database would enable search, audit trail, and reporting"),
    ]
    for bold_part, rest in future_items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after  = Pt(4)
        p.paragraph_format.space_before = Pt(2)
        r_b = p.add_run(bold_part)
        style_run(r_b, bold=True, size=11)
        r_t = p.add_run(rest)
        style_run(r_t, size=11)

    # ── Footer rule ──────────────────────────────────────────────────────────
    add_hr(doc, color="1F3864", sz=6)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.paragraph_format.space_before = Pt(4)
    r_f = footer_p.add_run("Rolls-Royce Digital Team  ·  Test & Inspection Plan Application  ·  Rev 7  ·  29 May 2026")
    r_f.font.name  = "Calibri"
    r_f.font.size  = Pt(9)
    r_f.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

    # ── Save ─────────────────────────────────────────────────────────────────
    out_path = "/home/user/test-inspection-plan/TIP_Architecture_Overview.docx"
    doc.save(out_path)
    print(f"Document saved: {out_path}")
    return out_path


if __name__ == "__main__":
    build_document()
