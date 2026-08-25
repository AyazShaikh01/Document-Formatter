#!/usr/bin/env python3
"""
docformat.py — consistent technical/academic Word formatting, no content rewriting.

Usage:
    python3 docformat.py input.docx output.docx [config.json]

config.json is optional — omit it and you get the built-in default look
(the one that's already been dialed in against real docs). Pass a config
file to override fonts, colors, sizes, heading rules, margins, image width,
and header/footer text. See config.example.json for every available key —
copy it, tweak what you want, leave the rest.

What it does (purely structural — it never changes your words):
  - First line becomes a styled Title.
  - "Part A:" / "Section 1:" style lines become Heading 1.
  - "1) Do the thing" style lines become Heading 2.
  - "a) sub-step" or "1.1 sub-step" style lines become Heading 3.
  - Existing bullet lists get consistent font/spacing (your bullets stay bullets).
  - Body paragraphs get one consistent font, color, size, spacing, justification —
    but any bold/italic you already put on specific words is preserved, not erased.
  - Blank "spacer" paragraphs are removed — spacing comes from styles, not manual returns.
  - Every image is resized to one consistent width and centered (floating/anchored
    screenshots are converted to inline first, splitting them out of shared
    paragraphs so they don't collide with nearby text).
  - ASCII/box-drawing diagrams (arrows, │─┌┐ etc.) are collected into a bordered,
    monospaced callout box instead of floating as plain text lines.
  - Tables get a shaded header row, banded body rows, and clean borders.
  - A running header (doc title) and footer (Page X of Y) are added.

What it deliberately does NOT do:
  - It does not fix grammar/spelling, generate image captions, or invent a subtitle/
    tagline. Those require reading and understanding the content, not just formatting
    it — this tool only touches structure and style, never wording.
"""

import json
import re
import sys
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE

# ---------------------------------------------------------------------------
# Default config — every value here can be overridden by config.json
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "fonts": {"body": "Calibri", "mono": "Consolas"},
    "colors": {
        "body": "3B3B3B",
        "heading1": "1F4E79",
        "heading2": "2E75B6",
        "heading3": "3B3B3B",
        "muted": "595959",
        "table_header_fill": "1F4E79",
        "table_header_text": "FFFFFF",
        "table_band_fill": "F2F2F2",
        "diagram_fill": "F7F9FC",
        "diagram_border": "2E75B6",
        "note_bg": "E7F1FA", "note_border": "2E75B6", "note_text": "1F4E79",
        "warning_bg": "FFF4E0", "warning_border": "C97A0C", "warning_text": "8A5000",
        "danger_bg": "FCEAE8", "danger_border": "B3261E", "danger_text": "8C1D14",
    },
    "sizes": {"title": 24, "heading1": 16, "heading2": 13, "heading3": 12, "body": 11},
    "bold": {"title": True, "heading1": True, "heading2": True, "heading3": False},
    "italic": {"title": False, "heading1": False, "heading2": False, "heading3": True},
    "spacing": {"body_after_pt": 8, "line_spacing": 1.15},
    "page": {"margin_in": 1.0, "image_max_width_in": 6.5},
    "header": {"enabled": True},
    "footer": {"enabled": True, "text_format": "Page {PAGE} of {NUMPAGES}"},
    # Each pattern is a Python regex tested against the paragraph's stripped
    # text with re.match (i.e. anchored at the start). First match wins, in
    # this order: heading1, heading2, heading3.
    "heading_patterns": {
        "heading1": [r"^\s*(Part|Section|Phase|Step)\s+[A-Z0-9]+\s*[:.\u2013\u2014-]"],
        "heading2": [r"^\s*\d+[\).](?!\d)\s*\S"],
        "heading3": [r"^\s*[a-z]\)\s*\S", r"^\s*\d+\.\d+\s*\S"],
    },
    "max_heading_length": 100,
    "bold_line_fallback": True,
    "bold_line_max_words": 12,
    "toc": {"enabled": True},
    # A paragraph starting with one of these (case-insensitive) becomes a
    # callout box instead of a normal paragraph. The matched prefix is
    # stripped; the rest of the line becomes the box's message.
    "callout_patterns": {
        "note": [r"^\s*note\s*:\s*"],
        "warning": [r"^\s*(warning|caution)\s*:\s*"],
        "danger": [r"^\s*(danger|critical)\s*:\s*"],
    },
}


def load_config(path):
    config = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
    if path:
        with open(path) as f:
            override = json.load(f)
        for section, values in override.items():
            if isinstance(values, dict) and isinstance(config.get(section), dict):
                config[section].update(values)
            else:
                config[section] = values
    return config


def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


# ---------------------------------------------------------------------------
# Style setup
# ---------------------------------------------------------------------------
def get_or_add_style(styles, name, style_type=WD_STYLE_TYPE.PARAGRAPH, base=None):
    try:
        return styles[name]
    except KeyError:
        style = styles.add_style(name, style_type)
        if base is not None:
            style.base_style = styles[base]
        return style


def configure_styles(doc, cfg):
    styles = doc.styles
    font = cfg["fonts"]["body"]
    body_color = hex_to_rgb(cfg["colors"]["body"])

    normal = styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(cfg["sizes"]["body"])
    normal.font.color.rgb = body_color
    pf = normal.paragraph_format
    pf.space_after = Pt(cfg["spacing"]["body_after_pt"])
    pf.line_spacing = cfg["spacing"]["line_spacing"]

    title = get_or_add_style(styles, "Title", base="Normal")
    title.font.name = font
    title.font.size = Pt(cfg["sizes"]["title"])
    title.font.bold = cfg["bold"]["title"]
    title.font.color.rgb = hex_to_rgb(cfg["colors"]["heading1"])

    h1 = get_or_add_style(styles, "Heading 1", base="Normal")
    h1.font.name = font
    h1.font.size = Pt(cfg["sizes"]["heading1"])
    h1.font.bold = cfg["bold"]["heading1"]
    h1.font.color.rgb = hex_to_rgb(cfg["colors"]["heading1"])
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after = Pt(12)

    h2 = get_or_add_style(styles, "Heading 2", base="Normal")
    h2.font.name = font
    h2.font.size = Pt(cfg["sizes"]["heading2"])
    h2.font.bold = cfg["bold"]["heading2"]
    h2.font.color.rgb = hex_to_rgb(cfg["colors"]["heading2"])
    h2.paragraph_format.space_before = Pt(16)
    h2.paragraph_format.space_after = Pt(8)

    h3 = get_or_add_style(styles, "Heading 3", base="Normal")
    h3.font.name = font
    h3.font.size = Pt(cfg["sizes"]["heading3"])
    h3.font.bold = cfg["bold"]["heading3"]
    h3.font.italic = cfg["italic"]["heading3"]
    h3.font.color.rgb = hex_to_rgb(cfg["colors"]["heading3"])
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(6)

    lp = get_or_add_style(styles, "List Paragraph", base="Normal")
    lp.font.name = font
    lp.font.size = Pt(cfg["sizes"]["body"])
    lp.font.color.rgb = body_color
    lp.paragraph_format.space_after = Pt(5)
    lp.paragraph_format.line_spacing = cfg["spacing"]["line_spacing"]

    margin = Inches(cfg["page"]["margin_in"])
    for section in doc.sections:
        section.top_margin = margin
        section.bottom_margin = margin
        section.left_margin = margin
        section.right_margin = margin


# ---------------------------------------------------------------------------
# Header / footer
# ---------------------------------------------------------------------------
def add_field_run(paragraph, field_code, font, size, color):
    def new_run():
        r = OxmlElement("w:r")
        paragraph._p.append(r)
        return r

    r1 = new_run()
    r1.append(_fld_char("begin"))

    r2 = new_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field_code
    r2.append(instr)

    r3 = new_run()
    r3.append(_fld_char("separate"))

    r4 = new_run()
    t = OxmlElement("w:t")
    t.text = "1"
    r4.append(t)

    r5 = new_run()
    r5.append(_fld_char("end"))

    for r_el in (r1, r2, r3, r4, r5):
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font)
        rFonts.set(qn("w:hAnsi"), font)
        rPr.append(rFonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(size * 2))
        rPr.append(sz)
        colorEl = OxmlElement("w:color")
        colorEl.set(qn("w:val"), str(color))
        rPr.append(colorEl)
        r_el.insert(0, rPr)


def _fld_char(kind):
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), kind)
    return el


def add_header_footer(doc, cfg, title_text):
    font = cfg["fonts"]["body"]
    muted = cfg["colors"]["muted"]

    if cfg["header"]["enabled"] and title_text:
        header = doc.sections[0].header
        p = header.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(title_text.upper())
        run.font.name = font
        run.font.size = Pt(9)
        run.font.color.rgb = hex_to_rgb(muted)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), cfg["colors"]["heading2"])
        pBdr.append(bottom)
        pPr.append(pBdr)

    if cfg["footer"]["enabled"]:
        footer = doc.sections[0].footer
        p = footer.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        template = cfg["footer"]["text_format"]
        before, _, rest = template.partition("{PAGE}")
        middle, _, after = rest.partition("{NUMPAGES}")
        if before:
            r = p.add_run(before)
            r.font.name = font
            r.font.size = Pt(9)
            r.font.color.rgb = hex_to_rgb(muted)
        add_field_run(p, "PAGE", font, 9, muted)
        if middle:
            r = p.add_run(middle)
            r.font.name = font
            r.font.size = Pt(9)
            r.font.color.rgb = hex_to_rgb(muted)
        if "{NUMPAGES}" in template:
            add_field_run(p, "NUMPAGES", font, 9, muted)
        if after:
            r = p.add_run(after)
            r.font.name = font
            r.font.size = Pt(9)
            r.font.color.rgb = hex_to_rgb(muted)


# ---------------------------------------------------------------------------
# Floating images -> inline
# ---------------------------------------------------------------------------
def convert_anchored_images_to_inline(doc):
    """Some editors save screenshots as floating (wp:anchor) images with an
    absolute position offset instead of inline images. Floating images ignore
    paragraph alignment entirely (they sit at a fixed x/y), so they can't be
    centered or reliably resized by the normal inline pass. Convert them to
    plain inline images so they behave like the rest of the document.

    If a floating image shares a paragraph with real text (e.g. a bullet's
    own sentence plus a screenshot the original author positioned to float
    beside/under it), the two must not end up sharing one inline text line —
    that produces overlapping, garbled layout. Split the image out into its
    own new paragraph first, then convert it.
    """
    body = doc.element.body
    drawings = body.findall(".//" + qn("w:drawing"))
    for drawing in drawings:
        anchor = drawing.find(qn("wp:anchor"))
        if anchor is None:
            continue

        run = drawing.getparent()
        para = run.getparent()
        if para.tag != qn("w:p"):
            para = para.getparent()

        other_text = "".join(
            t.text or ""
            for r in para.findall(qn("w:r"))
            if r is not run
            for t in r.findall(qn("w:t"))
        )

        if other_text.strip():
            new_p = OxmlElement("w:p")
            new_pPr = OxmlElement("w:pPr")
            new_jc = OxmlElement("w:jc")
            new_jc.set(qn("w:val"), "center")
            new_pPr.append(new_jc)
            new_p.append(new_pPr)
            run.getparent().remove(run)
            new_p.append(run)
            para.addnext(new_p)

        inline = OxmlElement("wp:inline")
        for attr in ("distT", "distB", "distL", "distR"):
            inline.set(attr, anchor.get(attr) or "0")
        for tag in ("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr"):
            child = anchor.find(qn(tag))
            if child is not None:
                inline.append(child)
        graphic = anchor.find(qn("a:graphic"))
        if graphic is not None:
            inline.append(graphic)
        anchor.getparent().replace(anchor, inline)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def is_blank(paragraph):
    return not paragraph.text.strip() and not paragraph._p.findall(".//" + qn("w:drawing"))


def has_numbering(paragraph):
    pPr = paragraph._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def has_drawing(paragraph):
    return bool(paragraph._p.findall(".//" + qn("w:drawing")))


def is_fully_bold(paragraph):
    text_runs = [r for r in paragraph.runs if r.text.strip()]
    if not text_runs:
        return False
    return all(_run_is_bold(r) for r in text_runs)


def _run_is_bold(run):
    """True bold, or bold via a character style (e.g. rStyle="Strong") —
    python-docx's run.font.bold only sees direct <w:b/>, not style-inherited
    bold, so check both."""
    if run.font.bold:
        return True
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return False
    rStyle = rPr.find(qn("w:rStyle"))
    return rStyle is not None and rStyle.get(qn("w:val")) in ("Strong", "Bold")


def is_diagram_signal(text):
    """A line that is clearly diagram content (arrows, box-drawing, ascii borders)."""
    if DIAGRAM_CHARS.search(text):
        return True
    return text.count("|") >= 1 or text.count("+") >= 2 or text.count("-") >= 3


DIAGRAM_CHARS = re.compile(r"[─│┌┐└┘├┤┬┴┼►▼▲◄→←↑↓]")


def is_diagram_blocker(paragraph, heading_patterns):
    """A line that can never be part of a diagram run — ends the run if seen."""
    if has_numbering(paragraph) or has_drawing(paragraph):
        return True
    text = paragraph.text.strip()
    if not text:
        return False  # blanks are neutral, not blockers
    for patterns in heading_patterns.values():
        for pat in patterns:
            if re.match(pat, text, re.IGNORECASE):
                return True
    words = text.split()
    if len(words) > 8 and text.endswith((".", ":", "!", "?")):
        return True  # a genuine prose sentence
    return False


def set_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="BFBFBF", sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def set_cell_borders_accent_left(cell, color, sz="4", left_sz="36"):
    """Thin border all around, thick accent stripe on the left — the callout-box look."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge, s in (("top", sz), ("bottom", sz), ("right", sz), ("left", left_sz)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), s)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def force_run_style(run, *, bold, italic, color, size, font):
    """Used for headings/titles/table headers — every run is forced uniform."""
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def restyle_body_run(run, *, color, size, font):
    """Used for ordinary body/list text — font/size/color become consistent,
    but any bold/italic emphasis you already had on this specific run is
    preserved rather than wiped out."""
    existing_bold = run.font.bold
    existing_italic = run.font.italic
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = existing_bold
    run.font.italic = existing_italic


# ---------------------------------------------------------------------------
# Body pass: classify paragraphs, apply styles, resize images, drop blanks
# ---------------------------------------------------------------------------
def format_body(doc, cfg):
    font = cfg["fonts"]["body"]
    body_color = hex_to_rgb(cfg["colors"]["body"])
    h1_color = hex_to_rgb(cfg["colors"]["heading1"])
    h2_color = hex_to_rgb(cfg["colors"]["heading2"])
    h3_color = hex_to_rgb(cfg["colors"]["heading3"])
    max_len = cfg["max_heading_length"]
    patterns = cfg["heading_patterns"]

    paragraphs = doc.paragraphs
    first_text_seen = False
    title_text = None

    for p in list(paragraphs):
        if p._p.getparent() is None:
            continue

        if is_blank(p):
            p._p.getparent().remove(p._p)
            continue

        if has_drawing(p):
            resize_images(p, cfg)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        if not first_text_seen:
            title_text = p.text.strip()
            p.style = doc.styles["Title"]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                force_run_style(
                    r, bold=cfg["bold"]["title"], italic=cfg["italic"]["title"],
                    color=h1_color, size=cfg["sizes"]["title"], font=font,
                )
            first_text_seen = True
            continue

        if has_numbering(p):
            p.style = doc.styles["List Paragraph"]
            for r in p.runs:
                restyle_body_run(r, color=body_color, size=cfg["sizes"]["body"], font=font)
            continue

        text = p.text.strip()

        if len(text) < max_len and any(re.match(pat, text, re.IGNORECASE) for pat in patterns["heading1"]):
            p.style = doc.styles["Heading 1"]
            for r in p.runs:
                force_run_style(
                    r, bold=cfg["bold"]["heading1"], italic=cfg["italic"]["heading1"],
                    color=h1_color, size=cfg["sizes"]["heading1"], font=font,
                )
            continue

        if len(text) < max_len and any(re.match(pat, text, re.IGNORECASE) for pat in patterns["heading2"]):
            p.style = doc.styles["Heading 2"]
            for r in p.runs:
                force_run_style(
                    r, bold=cfg["bold"]["heading2"], italic=cfg["italic"]["heading2"],
                    color=h2_color, size=cfg["sizes"]["heading2"], font=font,
                )
            continue

        if len(text) < max_len and any(re.match(pat, text, re.IGNORECASE) for pat in patterns["heading3"]):
            p.style = doc.styles["Heading 3"]
            for r in p.runs:
                force_run_style(
                    r, bold=cfg["bold"]["heading3"], italic=cfg["italic"]["heading3"],
                    color=h3_color, size=cfg["sizes"]["heading3"], font=font,
                )
            continue

        # Safety net: a short, standalone line the author already made fully
        # bold is almost always intended as a heading, even if it doesn't
        # match any of the specific patterns above (different phrasing per
        # doc is exactly what patterns can't chase forever). ALL CAPS reads
        # as a top-level section; everything else lands at Heading 2.
        if (
            cfg.get("bold_line_fallback", True)
            and len(text) < max_len
            and 0 < len(text.split()) <= cfg.get("bold_line_max_words", 12)
            and not text.endswith((".", "!", "?"))
            and is_fully_bold(p)
        ):
            level_style = "Heading 1" if text.isupper() else "Heading 2"
            level_key = "heading1" if text.isupper() else "heading2"
            color = h1_color if text.isupper() else h2_color
            p.style = doc.styles[level_style]
            for r in p.runs:
                force_run_style(
                    r, bold=cfg["bold"][level_key], italic=cfg["italic"][level_key],
                    color=color, size=cfg["sizes"][level_key], font=font,
                )
            continue

        # Plain body paragraph
        p.style = doc.styles["Normal"]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for r in p.runs:
            restyle_body_run(r, color=body_color, size=cfg["sizes"]["body"], font=font)

    return title_text


def resize_images(paragraph, cfg):
    max_width = Inches(cfg["page"]["image_max_width_in"])
    for inline in paragraph._p.findall(".//" + qn("wp:inline")):
        extent = inline.find(qn("wp:extent"))
        if extent is None:
            continue
        cx = int(extent.get("cx"))
        cy = int(extent.get("cy"))
        if cx > max_width:
            ratio = max_width / cx
            new_cx = max_width
            new_cy = Emu(int(cy * ratio))
            extent.set("cx", str(int(new_cx)))
            extent.set("cy", str(int(new_cy)))
            ext = inline.find(".//" + qn("a:ext"))
            if ext is not None:
                ext.set("cx", str(int(new_cx)))
                ext.set("cy", str(int(new_cy)))


# ---------------------------------------------------------------------------
# Diagram pass: turn runs of ASCII/box-drawing lines into a bordered callout
# ---------------------------------------------------------------------------
def convert_diagrams(doc, cfg):
    paragraphs = doc.paragraphs
    patterns = cfg["heading_patterns"]
    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        if p._p.getparent() is None or is_diagram_blocker(p, patterns):
            i += 1
            continue

        j = i
        candidates = []
        while j < len(paragraphs):
            pj = paragraphs[j]
            if pj._p.getparent() is None:
                j += 1
                continue
            if is_diagram_blocker(pj, patterns):
                break
            candidates.append(pj)
            j += 1

        has_signal = any(is_diagram_signal(c.text) for c in candidates)
        while candidates and not candidates[0].text.strip():
            candidates.pop(0)
        while candidates and not candidates[-1].text.strip():
            candidates.pop()

        if has_signal and candidates:
            lines = [c.text for c in candidates]
            anchor = candidates[-1]
            insert_diagram_box(anchor, lines, cfg)
            for c in candidates:
                if c._p.getparent() is not None:
                    c._p.getparent().remove(c._p)

        i = j


def insert_diagram_box(anchor_paragraph, lines, cfg):
    max_width = Inches(cfg["page"]["image_max_width_in"])
    body_color = hex_to_rgb(cfg["colors"]["body"])
    mono_font = cfg["fonts"]["mono"]

    doc_body = anchor_paragraph._parent
    table = doc_body.add_table(rows=1, cols=1, width=max_width)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_shading(cell, cfg["colors"]["diagram_fill"])
    set_cell_borders(cell, color=cfg["colors"]["diagram_border"], sz="4")

    cell.paragraphs[0].text = ""
    for idx, line in enumerate(lines):
        target = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        target.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = target.add_run(line if line.strip() else " ")
        force_run_style(run, bold=False, italic=False, color=body_color, size=10, font=mono_font)
        target.paragraph_format.space_after = Pt(0)

    tbl_element = table._tbl
    tbl_element.getparent().remove(tbl_element)
    anchor_paragraph._p.addprevious(tbl_element)


# ---------------------------------------------------------------------------
# Callout pass: "Note: ..." / "Warning: ..." / "Danger: ..." -> colored box
# ---------------------------------------------------------------------------
def match_callout(text, patterns):
    for kind, pats in patterns.items():
        for pat in pats:
            m = re.match(pat, text, re.IGNORECASE)
            if m:
                return kind, text[m.end():].strip()
    return None, None


def convert_callouts(doc, cfg):
    patterns = cfg["callout_patterns"]
    for p in list(doc.paragraphs):
        if p._p.getparent() is None:
            continue
        if has_drawing(p) or has_numbering(p):
            continue
        text = p.text.strip()
        if not text:
            continue
        kind, message = match_callout(text, patterns)
        if kind and message:
            insert_callout_box(p, kind, message, cfg)
            p._p.getparent().remove(p._p)


def insert_callout_box(anchor_paragraph, kind, message, cfg):
    font = cfg["fonts"]["body"]
    body_color = hex_to_rgb(cfg["colors"]["body"])
    bg = cfg["colors"][f"{kind}_bg"]
    border = cfg["colors"][f"{kind}_border"]
    text_color = hex_to_rgb(cfg["colors"][f"{kind}_text"])
    max_width = Inches(cfg["page"]["image_max_width_in"])

    doc_body = anchor_paragraph._parent
    table = doc_body.add_table(rows=1, cols=1, width=max_width)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_shading(cell, bg)
    set_cell_borders_accent_left(cell, border)

    label_p = cell.paragraphs[0]
    label_run = label_p.add_run(kind.upper())
    force_run_style(label_run, bold=True, italic=False, color=text_color, size=10, font=font)
    label_p.paragraph_format.space_after = Pt(2)

    msg_p = cell.add_paragraph()
    msg_run = msg_p.add_run(message)
    force_run_style(msg_run, bold=False, italic=False, color=body_color, size=11, font=font)
    msg_p.paragraph_format.space_after = Pt(0)

    tbl_element = table._tbl
    tbl_element.getparent().remove(tbl_element)
    anchor_paragraph._p.addprevious(tbl_element)


# ---------------------------------------------------------------------------
# Table of contents (real Word field — updates to actual page numbers the
# first time the reader right-clicks it and chooses "Update Field", same as
# any Word TOC; python-docx can't run Word's layout engine to know page
# numbers up front, so this is the standard way every docx-generation tool
# handles it)
# ---------------------------------------------------------------------------
def insert_toc(doc, cfg, title_paragraph):
    if not cfg["toc"]["enabled"]:
        return
    font = cfg["fonts"]["body"]
    muted_color = hex_to_rgb(cfg["colors"]["muted"])

    toc_heading = doc.add_paragraph()
    toc_heading.style = doc.styles["Heading 2"]
    r = toc_heading.add_run("Table of Contents")
    force_run_style(
        r, bold=cfg["bold"]["heading2"], italic=cfg["italic"]["heading2"],
        color=hex_to_rgb(cfg["colors"]["heading2"]), size=cfg["sizes"]["heading2"], font=font,
    )

    toc_p = doc.add_paragraph()
    hint_run = toc_p.add_run(
        "Right-click here and choose \u201cUpdate Field\u201d (Word) or press F9 "
        "to generate the table of contents."
    )
    force_run_style(hint_run, bold=False, italic=True, color=muted_color, size=10, font=font)

    _p = toc_p._p
    r1 = OxmlElement("w:r")
    r1.append(_fld_char("begin"))
    _p.append(r1)

    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    _p.append(r2)

    r3 = OxmlElement("w:r")
    r3.append(_fld_char("separate"))
    _p.append(r3)

    r4 = OxmlElement("w:r")
    r4.append(_fld_char("end"))
    _p.append(r4)

    # Move the two paragraphs (currently appended at the end of the body via
    # doc.add_paragraph) to sit right after the title.
    toc_heading_el = toc_heading._p
    toc_p_el = toc_p._p
    toc_heading_el.getparent().remove(toc_heading_el)
    toc_p_el.getparent().remove(toc_p_el)
    title_paragraph._p.addnext(toc_p_el)
    title_paragraph._p.addnext(toc_heading_el)

    # Page break after the TOC so the document body starts clean
    break_p = OxmlElement("w:p")
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    break_p.append(run)
    toc_p_el.addnext(break_p)

    # Ask Word to auto-recalculate all fields (the TOC and the page-number
    # fields in the footer) the moment the document is opened, so the
    # reader doesn't have to manually hit F9.
    settings = doc.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)


# ---------------------------------------------------------------------------
# Table pass: shaded header, banded rows, clean borders
# ---------------------------------------------------------------------------
def format_tables(doc, cfg):
    body_color = hex_to_rgb(cfg["colors"]["body"])
    header_text_color = hex_to_rgb(cfg["colors"]["table_header_text"])
    font = cfg["fonts"]["body"]

    for table in doc.tables:
        if len(table.rows) < 2 or len(table.columns) < 2:
            continue  # skip the 1-cell diagram callouts made above
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            for cell in row.cells:
                set_cell_borders(cell)
                for para in cell.paragraphs:
                    para.paragraph_format.space_after = Pt(4)
                    for run in para.runs:
                        if r_idx == 0:
                            force_run_style(
                                run, bold=True, italic=False,
                                color=header_text_color, size=cfg["sizes"]["body"], font=font,
                            )
                        else:
                            restyle_body_run(run, color=body_color, size=cfg["sizes"]["body"], font=font)
                if r_idx == 0:
                    set_shading(cell, cfg["colors"]["table_header_fill"])
                elif r_idx % 2 == 0:
                    set_shading(cell, cfg["colors"]["table_band_fill"])


# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) not in (3, 4):
        print("Usage: python3 docformat.py input.docx output.docx [config.json]")
        sys.exit(1)

    config_path = sys.argv[3] if len(sys.argv) == 4 else None
    cfg = load_config(config_path)

    doc = Document(sys.argv[1])
    configure_styles(doc, cfg)
    convert_anchored_images_to_inline(doc)
    convert_diagrams(doc, cfg)
    convert_callouts(doc, cfg)
    title_text = format_body(doc, cfg)
    format_tables(doc, cfg)
    add_header_footer(doc, cfg, title_text)
    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    if title_paragraph is not None:
        insert_toc(doc, cfg, title_paragraph)
    doc.save(sys.argv[2])
    print(f"Formatted document written to {sys.argv[2]}")


if __name__ == "__main__":
    main()
