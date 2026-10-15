#!/usr/bin/env python3
"""Build the two day-specific decks from the original one-day deck.

############################################################################
# WARNING — BUILD OUTPUT IS NOT THE FINAL VERSION OF THE DECKS
#
# The checked-in slides/UChicago_Day1_Slides.pptx and
# slides/UChicago_Day2_Slides.pptx received additional visual tweaks
# (layout, link styling, polish) made by hand, outside this script.
# Running this script will regenerate decks whose CONTENT is identical
# and fully reproducible, but whose formatting will not match the final
# checked-in files. Do not overwrite the final decks without re-applying
# those tweaks.
############################################################################

Day 1: slides 1-31 + 67-79 of the original, edited for the new
  schedule, plus new slides (VaR worked example, two quick exercises,
  Day-2 preview).
Day 2: slide 1 + 32-66 (minus 49-50) + 76-79, edited, plus new
  slides (welcome/teams, Day-1 recap, arc, tool use & agents, context
  engineering, extension-sprint section).

Approach: copy the original .pptx, edit text runs in place, append new
slides built with python-pptx in the same visual style, then rewrite the
sldIdLst to delete/reorder, and renumber footers.
"""

import copy
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from lxml import etree

SRC = "slides/UChicago_Workshop_Slides.pptx"
DAY1 = "slides/UChicago_Day1_Slides.pptx"
DAY2 = "slides/UChicago_Day2_Slides.pptx"

NAVY = "0B2545"; MID = "13315C"; STEEL = "8DA9C4"; TINT = "EEF2F7"
AMBER = "E09F3E"; CODEBG = "1E2433"; INKDARK = "0B2545"; BODY = "1F2937"
GREY = "6B7280"; CODEFG = "E6EDF3"; WHITE = "FFFFFF"

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

FOOTER = "From Excel to Python and AI · University of Chicago"


# ---------------------------------------------------------------- helpers

def paras_text(tf):
    return "\n".join("".join(r.text for r in p.runs) for p in tf.paragraphs)


def replace_sub(slide, old, new, must=True):
    """Replace substring `old` with `new` inside any single paragraph."""
    hits = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            joined = "".join(r.text for r in p.runs)
            if old in joined:
                newtext = joined.replace(old, new)
                if p.runs:
                    p.runs[0].text = newtext
                    for r in p.runs[1:]:
                        r.text = ""
                hits += 1
    if must and hits == 0:
        raise SystemExit(f"replace_sub: NOT FOUND: {old[:70]!r}")
    return hits


def replace_run(slide, old, new, must=True):
    """Replace substring `old` with `new` inside individual runs only,
    preserving per-run formatting (unlike replace_sub, which flattens
    a paragraph's runs to the first run's format)."""
    hits = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new)
                    hits += 1
    if must and hits == 0:
        raise SystemExit(f"replace_run: NOT FOUND: {old[:70]!r}")
    return hits


def grab_fmt(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            f = r.font
            col = None
            try:
                if f.color and f.color.type is not None:
                    col = f.color.rgb
            except Exception:
                pass
            return (f.name, f.size, f.bold, col)
    return ("Calibri", Pt(14), None, RGBColor.from_string(BODY))


def set_box_text(slide, contains, new_text, must=True):
    """Rebuild an entire text frame whose current text contains `contains`,
    reapplying the first run's formatting to every new run."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        if contains in paras_text(tf):
            name, size, bold, col = grab_fmt(tf)
            tf.text = new_text
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.name = name
                    if size: r.font.size = size
                    r.font.bold = bold
                    if col is not None: r.font.color.rgb = col
            return True
    if must:
        raise SystemExit(f"set_box_text: NOT FOUND: {contains[:70]!r}")
    return False


def set_box_bullets(slide, contains, items, must=True):
    """Rebuild a text frame as proper bulleted paragraphs (hanging indent +
    buChar). `items` are strings or (bold_lead, rest) pairs; base formatting
    is copied from the frame's first run."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        if contains in paras_text(tf):
            name, size, bold, col = grab_fmt(tf)
            tf.text = ""
            for i, it in enumerate(items):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                lead, rest = it if isinstance(it, tuple) else (None, it)
                if lead:
                    r = p.add_run(); r.text = lead
                r = p.add_run(); r.text = rest
                for r in p.runs:
                    r.font.name = name
                    if size: r.font.size = size
                    if col is not None: r.font.color.rgb = col
                if lead:
                    p.runs[0].font.bold = True
                p.space_after = Pt(7)
                p.line_spacing = 1.12
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", "171450"); pPr.set("indent", "-171450")
                bu = etree.SubElement(pPr, f"{A}buChar"); bu.set("char", "•")
            return True
    if must:
        raise SystemExit(f"set_box_bullets: NOT FOUND: {contains[:70]!r}")
    return False


def set_bg(slide, hexcolor):
    cSld = slide._element.find(f"{P}cSld")
    old = cSld.find(f"{P}bg")
    if old is not None:
        cSld.remove(old)
    bg = etree.SubElement(cSld, f"{P}bg")
    bgPr = etree.SubElement(bg, f"{P}bgPr")
    fill = etree.SubElement(bgPr, f"{A}solidFill")
    clr = etree.SubElement(fill, f"{A}srgbClr"); clr.set("val", hexcolor)
    etree.SubElement(bgPr, f"{A}effectLst")
    cSld.insert(0, bg)


def tb(slide, x, y, w, h, text, font="Calibri", size=14, bold=False,
       color=BODY, align=None, line_spacing=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        if align: p.alignment = align
        if line_spacing: p.line_spacing = line_spacing
        for r in p.runs:
            r.font.name = font; r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = RGBColor.from_string(color)
    return box


def rrect(slide, x, y, w, h, fill=TINT, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sp.adjustments[0] = 0.03
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    sp.line.color.rgb = RGBColor.from_string(line or fill)
    sp.line.width = Pt(1)
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.text = ""
    return sp


def card(slide, x, y, w, h, title, body, title_color=INKDARK, body_size=13):
    rrect(slide, x, y, w, h)
    tb(slide, x + 0.2, y + 0.15, w - 0.4, 0.4, title,
       size=14, bold=True, color=title_color)
    tb(slide, x + 0.2, y + 0.62, w - 0.4, h - 0.8, body, size=body_size,
       color=BODY, line_spacing=1.12)


def code_box(slide, x, y, w, h, code, size=13):
    rrect(slide, x, y, w, h, fill=CODEBG)
    tb(slide, x + 0.25, y + 0.2, w - 0.5, h - 0.4, code,
       font="Courier New", size=size, color=CODEFG, line_spacing=1.15)


# ---- notebook-moment badges ------------------------------------------
# Every slide that asks students to open (or run along with) a notebook
# carries a colored strip under the title: a notebook glyph, the mode of
# use, and the exact notebook. One color per functional mode; the three
# labs share a teal family in deepening shades.
NB_MODES = {
    "SETUP":    ("C2410C", "SETUP CHECK · RUN NOW"),
    "EXERCISE": ("C2410C", "EXERCISE · TYPE ALONG"),
    "WATCH":    ("475569", "WORKED EXAMPLE · WATCH"),
    "LAB A":    ("0E7490", "LAB A · WORK SOLO"),
    "LAB B":    ("0F766E", "LAB B · WORK SOLO"),
    "LAB C":    ("047857", "LAB C · WORK SOLO"),
    "DEMO":     ("6D28D9", "LIVE DEMO · RUN ALONG"),
    "CASE":     ("1D4ED8", "CASE · NOTEBOOK OPEN"),
    "SPRINT":   ("BE185D", "TEAM SPRINT · BUILD"),
}


# Colab deep links (same pattern as the README badges)
COLAB = ("https://colab.research.google.com/github/limorigu/"
         "uchicago-python-ai-workshop/blob/main/notebooks/")


def _shape_text(sp, text, color, size=11.5, align=None, margin_left=0.1):
    tf = sp.text_frame
    tf.text = text
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(margin_left); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0); tf.margin_bottom = Inches(0)
    for p in tf.paragraphs:
        if align: p.alignment = align
        for r in p.runs:
            r.font.name = "Calibri"; r.font.size = Pt(size); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string(color)


def _nb_glyph(slide, col, x, y):
    """Small notebook: white page with rule lines in the badge color."""
    rrect(slide, x, y, 0.20, 0.25, fill=WHITE)
    for i in range(3):
        rrect(slide, x + 0.045, y + 0.05, 0.11, 0.018, fill=col)
        y += 0.062


def nb_badge(slide, mode, detail, notebook=None, y=1.38):
    """Notebook-moment badge: tinted base strip · colored mode pill with
    glyph · detail in navy · white Open-in-Colab chip (the whole chip is
    the hyperlink — no link-styled text on a colored fill)."""
    col, label = NB_MODES[mode]
    rrect(slide, 0.6, y, 12.13, 0.46, fill=TINT)
    pill = rrect(slide, 0.68, y + 0.05, 2.78, 0.36, fill=col)
    _nb_glyph(slide, col, 0.80, y + 0.105)
    _shape_text(pill, label, WHITE, margin_left=0.42)
    tb(slide, 3.62, y + 0.105, 6.55, 0.3, detail, size=12, bold=True,
       color=NAVY)
    if notebook:
        chip = rrect(slide, 10.32, y + 0.05, 2.33, 0.36, fill=WHITE, line=col)
        _shape_text(chip, "Open in Colab  ↗", col, align=PP_ALIGN.CENTER,
                    margin_left=0.05)
        chip.click_action.hyperlink.address = COLAB + notebook


def nb_hint(slide, mode, pill_text, text, y=6.5):
    """Small foreshadowing marker (bottom of a talk slide): tiny colored
    pill + one line in the same color — references a notebook moment that
    is coming up, without telling anyone to open it yet."""
    col, _ = NB_MODES[mode]
    pill = rrect(slide, 0.6, y, 1.95, 0.3, fill=col)
    _shape_text(pill, pill_text, WHITE, size=10, align=PP_ALIGN.CENTER,
                margin_left=0.05)
    tb(slide, 2.7, y + 0.025, 10.0, 0.3, text, size=11, bold=True, color=col)


def break_slide(prs, kicker, title, back, note):
    """Full-bleed demarcation slide for breaks and lunch, so the agenda
    lives in the deck itself."""
    s = prs.slides.add_slide(prs.slide_layouts[0])
    set_bg(s, MID)
    tb(s, 0.6, 1.9, 6.0, 0.4, kicker, size=14, bold=True, color=AMBER)
    tb(s, 0.6, 2.4, 12.0, 1.2, title, font="Cambria", size=44, bold=True,
       color=WHITE)
    tb(s, 0.6, 3.7, 11.5, 0.5, back, size=20, bold=True, color=AMBER)
    tb(s, 0.6, 4.5, 11.5, 1.8, note, size=16, color=WHITE, line_spacing=1.25)
    return s


def content_slide(prs, kicker, title):
    layout = prs.slide_layouts[0]
    s = prs.slides.add_slide(layout)
    set_bg(s, WHITE)
    tb(s, 0.6, 0.4, 10.0, 0.3, kicker, size=10, bold=True, color=AMBER)
    tb(s, 0.6, 0.7, 12.1, 0.9, title, font="Cambria", size=30, bold=True,
       color=NAVY)
    tb(s, 0.6, 7.0, 6.0, 0.3, FOOTER, size=9, color=GREY)
    tb(s, 11.9, 7.0, 0.8, 0.3, "0", size=9, color=GREY)
    return s


def divider_slide(prs, section, title, timing, body):
    layout = prs.slide_layouts[0]
    s = prs.slides.add_slide(layout)
    set_bg(s, NAVY)
    tb(s, 0.6, 1.6, 6.0, 0.4, section, size=14, bold=True, color=AMBER)
    tb(s, 0.6, 2.1, 11.5, 1.6, title, font="Cambria", size=40, bold=True,
       color=WHITE)
    tb(s, 0.6, 3.9, 11.0, 0.4, timing, size=16, color=STEEL)
    tb(s, 0.6, 4.52, 8.0, 0.3, "BY THE END OF THIS SECTION", size=11, bold=True,
       color=AMBER)
    tb(s, 0.6, 4.95, 11.2, 1.8, body, size=18, color=WHITE, line_spacing=1.15)
    return s


def divider_goals(slide, contains, goals):
    """Reformat a SRC divider's bottom line: amber 'BY THE END OF THIS SECTION'
    label above a dot-separated list of section goals."""
    for sp in slide.shapes:
        if sp.has_text_frame and contains in sp.text_frame.text:
            set_box_text(slide, contains, goals)
            sp.top = Inches(4.95)
            break
    else:
        raise AssertionError(f"divider body not found: {contains}")
    tb(slide, 0.6, 4.52, 8.0, 0.3, "BY THE END OF THIS SECTION", size=11,
       bold=True, color=AMBER)


def chain(slide, items, y=2.3, x0=0.7, box_w=1.65, gap=0.42, box_h=1.05):
    """Row of navy chip boxes joined by › arrows (like the arc slide)."""
    x = x0
    for i, item in enumerate(items):
        rrect(slide, x, y, box_w, box_h, fill=TINT)
        tb(slide, x + 0.08, y + 0.12, box_w - 0.16, box_h - 0.2, item,
           size=11.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        x += box_w
        if i < len(items) - 1:
            tb(slide, x + 0.02, y + 0.22, gap, 0.5, "›", size=20, bold=True,
               color=AMBER, align=PP_ALIGN.CENTER)
            x += gap


def rebuild_order(prs, keep_order):
    """keep_order: list of slide objects in desired final order; everything
    else is deleted (rels dropped)."""
    lst = prs.slides._sldIdLst
    id_by_slide = {}
    slides = list(prs.slides)
    for sldId, slide in zip(list(lst), slides):
        id_by_slide[slide.slide_id] = sldId
    wanted_ids = {s.slide_id for s in keep_order}
    for slide in slides:
        if slide.slide_id not in wanted_ids:
            sldId = id_by_slide[slide.slide_id]
            rId = sldId.get(f"{R}id")
            lst.remove(sldId)
            prs.part.drop_rel(rId)
    for s in keep_order:
        sldId = id_by_slide[s.slide_id]
        lst.remove(sldId)
        lst.append(sldId)


def renumber_footers(prs):
    for i, slide in enumerate(prs.slides, start=1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            if abs(sh.left - Inches(11.9)) < Inches(0.2) and \
               abs(sh.top - Inches(7.0)) < Inches(0.2):
                txt = sh.text_frame.text.strip()
                if txt.isdigit() or txt == "0":
                    p = sh.text_frame.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = str(i)
                        for r in p.runs[1:]:
                            r.text = ""


def strip_notes(prs):
    """Remove all speaker-note text — the decks go to students; instructor
    notes live in INSTRUCTOR_NOTES.md, not in distributed files."""
    for slide in prs.slides:
        if slide.has_notes_slide:
            slide.notes_slide.notes_text_frame.clear()


def find_table(slide):
    for sh in slide.shapes:
        if sh.has_table:
            return sh.table
    raise SystemExit("table not found")


# ================================================================ DAY 1

def build_day1():
    shutil.copy(SRC, DAY1)
    prs = Presentation(DAY1)
    S = list(prs.slides)          # original order, 0-indexed

    # ---- slide 1 (title)
    replace_sub(S[0], "A one-day workshop for finance students",
                "Day 1 of 2 · Python bootcamp for finance students")
    # (time line on the title slide stays date-free)

    # ---- slide 2
    replace_sub(S[1],
        "By 5:00 you will have a number, a baseline, and a list of the exact cases where the model was wrong.",
        "By tomorrow at 3:00 you will have a number, a baseline, and a list of the exact cases where the model was wrong. Today builds the Python footing that makes that measurement possible.")

    # ---- slide 3 (arc) : re-label the 8 stops left-to-right
    stops = []
    for sh in S[2].shapes:
        if sh.has_text_frame and abs(sh.top - Inches(2.2)) < Emu(91440):
            txt = sh.text_frame.text.strip()
            if txt and txt != "›":
                stops.append(sh)
    stops.sort(key=lambda sh: sh.left)
    new_stops = ["Welcome\n& setup", "Python I\n+ Lab A", "Python II\n+ Lab B",
                 "Python III\n+ Lab C", "Data\nmanagement",
                 "Case 1\nFCC whitespace", "Git &\nCI/CD", "Wrap & Day 2\npreview"]
    assert len(stops) == 8, f"expected 8 arc stops, got {len(stops)}"
    for sh, txt in zip(stops, new_stops):
        name, size, bold, col = grab_fmt(sh.text_frame)
        sh.text_frame.text = txt
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.name = name
                if size: r.font.size = size
                r.font.bold = bold
                if col is not None: r.font.color.rgb = col
    set_box_text(S[2], "Afternoon: an LLM becomes a component",
        "Morning: three short talks, each followed by its own lab — variables and a VaR calculation, collections and functions on a five-position book, then pandas on 236,549 rows of FCC data.\n \nAfternoon: a live Excel→pandas demo, data management, and Case 1 end-to-end — rank whitespace counties and deliver an Excel workbook a partner can open.\n \nLast hour: version control and CI, the habits that make the work shareable. Tomorrow an LLM becomes a measured component in your analysis — today earns that.")

    # ---- slide 4
    replace_sub(S[3], "Talks are 20 minutes. Labs end when the clock says so",
        "Talks are 25–30 minutes, each followed by its own lab. Labs end when the clock says so")
    replace_sub(S[3], "Groups of 3–4 with one experienced coder each.",
        "You sit at assigned tables of 3–4, fixed all day, one experienced coder each — ask your table before you ask me.")

    # ---- slide 5 (setup check — a notebook moment, badge it)
    nb_badge(S[4], "SETUP", "00_setup_check.ipynb · run every cell · 5 minutes",
             notebook="00_setup_check.ipynb")

    # ---- slide 6 (divider, SECTION 1)
    replace_sub(S[5], "20-minute talk · 40-minute lab (9:20–10:20)",
        "Three talks · three labs · 9:25–2:10 (lunch 12:00–1:00)")
    replace_sub(S[5], "By 10:20 you can read and write the eight constructs",
        "By 2:10 you can read and write the eight constructs")

    # ---- slide 13: reorder the Excel→pandas table, gray rows last
    table = find_table(S[12])
    tbl = table._tbl
    trs = tbl.findall(f"{A}tr")
    assert len(trs) == 11, f"expected 11 rows, got {len(trs)}"
    # data rows (1-indexed into trs): keep header 0
    # old: 1 Filter, 2 Calc, 3 Sum, 4 Countif, 5 Pivot, 6 Lookup, 7 Dedup,
    #      8 Clean, 9 Sort, 10 Save
    new_order = [0, 1, 2, 3, 9, 4, 7, 10, 5, 6, 8]
    rows = [trs[i] for i in new_order]
    for tr in trs:
        tbl.remove(tr)
    for tr in rows:
        tbl.append(tr)
    nb_hint(S[12], "DEMO", "LIVE DEMO AHEAD",
        "the last three rows (gray) — pivot, lookup, text cleaning — each get a run-along demo in a few slides: 01b_excel_to_pandas_demo.ipynb")

    # ---- slide 15: Colab-demo pointer
    nb_hint(S[14], "DEMO", "LIVE DEMO AHEAD",
        "both habits shown live in Colab in a moment — 01b_excel_to_pandas_demo.ipynb, on a trades table small enough to check by eye")

    # ---- slide 17 (original lab slide → Lab A launch, right after the Python I talk)
    replace_sub(S[16], "Lab · 40 minutes · 01_python_bootcamp.ipynb",
        "Lab A · 9:55 · 35 minutes — variables and a VaR")
    set_box_text(S[16], "Run each cell; do every # TODO beneath it.",
        "Open 01_python_bootcamp.ipynb, section Lab A — run each cell; do every # TODO beneath it.\nThe worked example is your starting point: run the four-line VaR, change the position to $40M, add a 95% version (z = 1.65).\nCheck-out: one sentence — what would have been painful in Excel?")
    set_box_text(S[16], "Bar chart of offers per technology.",
        "⚡ cells at the end of the section: scale the VaR to 10 days with √t.\nThen help a neighbour. Hands off their keyboard.")
    set_box_text(S[16], "10:00 — everyone has run avail.info()",
        "10:30 — your VaR runs and you re-ran it with a bigger position. Break.")
    nb_badge(S[16], "LAB A", "01_python_bootcamp.ipynb · section Lab A · 35 min",
             notebook="01_python_bootcamp.ipynb")

    # ---- slide 18 (divider, SECTION 2)
    replace_sub(S[17], "20-minute talk (10:35–10:55)", "20-minute talk (2:25–2:45)")

    # ---- slide 26 provenance
    replace_sub(S[25], "(3:55 today)", "(4:20 today)")

    # ---- slide 28 (divider, SECTION 3)
    replace_sub(S[27],
        "5-minute brief · 30 min individual · 20 min group · 10 min debrief (10:55–12:00)",
        "10-minute brief · 45 min solo · 25 min with your table · 15 min debrief (2:45–4:20)")

    # ---- slide 31 Case 1 deliverables
    replace_sub(S[30], "Individual · 30 min", "Part 1 · solo · 45 min")
    replace_sub(S[30], "Group · 20 min", "Part 2 · your table · 25 min")
    replace_sub(S[30], "groups of 3–4", "hard switch at 3:40", must=False)
    replace_sub(S[30], "read it, close it, type it yourself.",
        "read it, close it, type it yourself. At 3:40 the whole table switches to Part 2 — start from the table's most complete notebook; an unfinished Part 1 is normal.")
    replace_sub(S[30], "defend it at 11:50", "defend it at 4:05")
    nb_badge(S[30], "CASE", "02_fcc_whitespace_STARTER.ipynb · solo, then your table",
             notebook="02_fcc_whitespace_STARTER.ipynb")

    # ---- slide 67 (divider, SECTION 8 → 4)
    replace_sub(S[66], "SECTION 8", "SECTION 4")
    replace_sub(S[66], "12-minute talk + 13-minute exercise (3:55–4:20)",
        "15-minute talk + 15-minute exercise (4:20–4:50)")

    # ---- slide 73 tests — retold with Case 1 examples (no Case 2 jargon).
    # must=False: these no-op once the SRC deck is rebuilt from build_slides.js,
    # which now carries the same content.
    set_box_text(S[72], "def test_lexicon_handles_negation", (
        "def test_geoid_keeps_leading_zeros():\n"
        "    locs = load_locations(\"bdc_il_locations_sample.csv.gz\")\n"
        "    assert locs.block_geoid.str.len().eq(15).all()\n"
        "\n"
        "def test_satellite_never_counts_as_served():\n"
        "    served = served_offers(avail, down_min=100, up_min=20)\n"
        "    assert not served.technology.isin(SATELLITE).any()\n"
        "\n"
        "def test_whitespace_blocks_are_adjacent_only():\n"
        "    ws = blk[blk.whitespace]\n"
        "    assert (ws.relation_to_footprint == \"adjacent\").all()"),
        must=False)
    replace_run(S[72],
        "Part A gave you three lexicon failures. One of them becomes a test in 13 minutes.",
        "This afternoon's case handed you real ones — the GEOID dtype, the satellite rows, the whitespace rule. Each fix could ship with the test that keeps it fixed.",
        must=False)
    replace_sub(S[72],
        "The gold set from Part B can run as a test too: 'accuracy on the labelled set must stay above 0.85'. That is how an evaluation becomes a living asset.",
        "A finished analysis can run as a test too: 'the county ranking must not change when the pipeline re-runs on the same data'. That is how an analysis becomes a living asset.",
        must=False)

    # ---- slides 43–48: keep Git/CI section self-contained within Day 1 ----
    # The git story is retold with Case 1 (whitespace) artifacts instead of
    # Case 2 (news classification) artifacts students only meet tomorrow.
    # must=False so these no-op once the SRC deck is rebuilt from build_slides.js.
    # slide 43 — commit chips + branch
    replace_run(S[68], "initial newslib", "initial case notebook", must=False)
    replace_run(S[68], "add lexicon negation", "fix GEOID leading zeros", must=False)
    replace_run(S[68], "PROMPT_V2 + tests", "whitespace flag + tests", must=False)
    replace_run(S[68], "fix: relevance gate", "fix: exclude satellite", must=False)
    replace_run(S[68], "merge: few-shot prompt", "merge: 25/3 sensitivity", must=False)
    replace_run(S[68], "try few-shot prompt", "try 25/3 threshold", must=False)
    replace_run(S[68], "git switch -c try-few-shot-prompt", "git switch -c try-25-3-threshold", must=False)
    replace_run(S[68], "try the few-shot prompt without touching main",
        "rerun the case at the older 25/3 threshold without touching main", must=False)
    replace_run(S[68], "prompt tests", "25/3 county table", must=False)
    # slide 44 — commit-message example
    replace_run(S[69], "'fix relevance gate for name collisions' beats 'update newslib.py'",
        "'fix GEOID dtype so leading zeros survive' beats 'update notebook'", must=False)
    # slide 45 — API key example is tomorrow's code
    replace_run(S[70], "anthropic.Anthropic() looks for ANTHROPIC_API_KEY on its own",
        "tomorrow's anthropic.Anthropic() will find ANTHROPIC_API_KEY on its own", must=False)
    # slide 46 — ci.yml now installs pandas+pytest and runs the whole suite
    replace_run(S[71], "# 3 dependencies", " ", must=False)  # orphan comment line first
    replace_run(S[71], "pip install -r app/requirements.txt pytest",
        "pip install pandas pytest     # 3 dependencies", must=False)
    replace_run(S[71], "pytest app/tests -q", "pytest -q          ", must=False)
    # slide 47 — the shown tests are Case 1 flavoured and ship in case1/tests/
    replace_run(S[72], "Seven of them live in app/tests/.",
        "These three pin down bugs you met in this afternoon's case. They ship in the repo under case1/tests/ — in the exercise you push them through CI yourself.", must=False)
    # slide 48 — green-check limits without Case 2 jargon
    replace_run(S[73], "the analysis is right, the labels are right, the prompt is good, or the data is current",
        "the analysis is right, the assumptions are sound, or the data is current", must=False)
    replace_run(S[73], "They do not cover 'the model understands hotel economics'.",
        "They do not cover 'the analysis answers the right question'.", must=False)
    replace_run(S[73], "Seven tests cover parsing, cleaning, negation, name collisions and metrics.",
        "The exercise's four tests cover the GEOID dtype, the satellite rule and the whitespace cut.", must=False)

    # ---- slide 75 exercise — rebuilt in the lab-slide style: badge strip,
    # Do/Why cards, checkpoint line, direct links. Old shorthand boxes removed.
    replace_sub(S[74], "Exercise · 13 minutes · push, break, fix",
        "Exercise · 4:35 · 15 minutes — push, break, fix", must=False)
    KEEP = ("CI/CD", "Exercise ·", "From Excel to Python")
    for sh in list(S[74].shapes):
        txt = sh.text_frame.text if sh.has_text_frame else ""
        if (txt.strip() and any(k in txt for k in KEEP)) or txt.strip().isdigit():
            continue                      # keep eyebrow, title, footer, page number
        sh._element.getparent().remove(sh._element)
    # badge strip (exercise orange), with a live link chip
    EX_COL = "C2410C"
    rrect(S[74], 0.6, 1.38, 12.13, 0.46, fill=TINT)
    pill = rrect(S[74], 0.68, 1.43, 2.78, 0.36, fill=EX_COL)
    _nb_glyph(S[74], EX_COL, 0.80, 1.485)
    _shape_text(pill, "EXERCISE · YOUR GROUP", WHITE, margin_left=0.42)
    tb(S[74], 3.62, 1.485, 6.55, 0.3,
       "all in the browser on github.com · nothing to install · 15 min",
       size=12, bold=True, color=NAVY)
    chip = rrect(S[74], 10.32, 1.43, 2.33, 0.36, fill=WHITE, line=EX_COL)
    _shape_text(chip, "github.com/new  ↗", EX_COL, align=PP_ALIGN.CENTER,
                margin_left=0.05)
    chip.click_action.hyperlink.address = "https://github.com/new"
    # left: exact steps
    card(S[74], 0.6, 1.98, 7.35, 4.3, "Do — exact steps (one driver, everyone reads along)",
         "1 · github.com/new → name it ci-exercise → Public → Create repository\n"
         "2 · Get the code: workshop repo → green Code button → Download ZIP → unzip\n"
         "3 · Your new repo → Add file → Upload files → drag in case1/ and .github/ → Commit changes. (.github must sit at the repo root — it is the only place GitHub looks for workflows)\n"
         "4 · Actions tab → ci runs → green ✓ — this afternoon's whitespace tests passing on GitHub's machine, not yours\n"
         "5 · Break it: open case1/whitespace.py → pencil icon (Edit) → in served_offers, delete the line & ~avail[\"technology\"].isin(SATELLITE) → Commit\n"
         "6 · Actions tab → red ✗ → click the run, then the failed step. Read aloud: which test failed, what it expected, what it saw\n"
         "7 · Fix it: paste the satellite line back → Commit → green ✓",
         body_size=11.5)
    # right: why + fast lane
    card(S[74], 8.15, 1.98, 4.58, 2.55, "Why step 6 is the point",
         "The red ✗ is CI doing its job: test_satellite_never_counts_as_served names the exact bug from this afternoon — satellite rows quietly counting as served — and catches it before anyone downstream builds on the wrong number. Reading a failure calmly (test name, expected, got) is the skill.",
         body_size=11.5)
    card(S[74], 8.15, 4.68, 4.58, 1.6, "Done early? ⚡",
         "Add a test of your own: at 25/3 thresholds, served_offers must return at least as many rows as at 100/20. Then try it on a branch and open a pull request — CI runs on the PR before you merge.",
         body_size=11.5)
    tb(S[74], 0.6, 6.42, 12.1, 0.45,
       "Checkpoint, 4:50 — your repo shows a green check, everyone saw the red one, and anyone in the group can name what the failing test checked.",
       size=14, bold=True, color=NAVY)

    # ---- slide 76 (divider, SECTION 9 → 5)
    replace_sub(S[75], "SECTION 9", "SECTION 5")

    # ---- slide 77: retitle as takeaways (no-op once SRC deck is rebuilt from build_slides.js)
    replace_sub(S[75], "What to do on Monday morning, and where to go next.",
        "What to leave today with — and where to go next.", must=False)
    replace_sub(S[76], "Monday morning", "Takeaways — put them to work Monday", must=False)

    # ---- slide 77: 4th card becomes Day-2 teaser
    replace_sub(S[76], "Keep a prompt log", "Come back tomorrow")
    set_box_text(S[76], "Version, change, score. Even for chat use.",
        "Day 2, same room, 9:00–3:00: how LLMs actually work, context engineering, Case 2 — an AI news screen you measure before trusting — ending in a team build sprint with UChicago tech students. Bring this laptop.")

    # ---- slide 79 thank-you
    replace_sub(S[78],
        "Feedback form: two minutes, three questions — please, before you close the laptop.",
        "See you tomorrow at 9:00 — AI day, joint with the tech cohort.")

    # ---------------- new slides
    # VaR worked example (after slide 7)
    var_s = content_slide(prs, "PYTHON BOOTCAMP",
                          "Worked example · a 1-day VaR in four lines")
    code_box(var_s, 0.6, 1.9, 6.4, 3.3,
"""position_value = 25_000_000   # $25M of HLT shares
daily_vol = 0.018             # 1.8% daily volatility
z_99 = 2.33                   # 99% confidence

var_1d = position_value * daily_vol * z_99
print(f"1-day 99% VaR: ${var_1d:,.0f}")

# 1-day 99% VaR: $1,048,500""")
    card(var_s, 7.3, 1.9, 5.4, 3.3, "Why a risk team cares",
         "Value at Risk: “on 99 days out of 100, this position should not lose more than X in a day.” Risk desks compute it every night, for every position, at several confidence levels.\nIn Excel this is one formula — until it is per position, per confidence level, per holding period. Then names beat cell references: change one line, everything downstream updates.")
    tb(var_s, 0.6, 5.5, 12.1, 1.2,
       "Lab A, in a few minutes: run exactly this, change the position to $40M, add a 95% version (z = 1.65), and scale to 10 days with √t. Every variable on this slide reappears on Day 2 inside the news screen for the same hotel company.",
       size=14, color=BODY, line_spacing=1.15)
    nb_badge(var_s, "WATCH", "no notebook yet — this exact code becomes Lab A in a few minutes")

    # Exercise after slide 9 (lists & dicts)
    ex1 = content_slide(prs, "PYTHON BOOTCAMP · EXERCISE",
                        "Exercise · 3 minutes · your first book")
    code_box(ex1, 0.6, 1.9, 6.4, 3.3,
"""book = {
    "HLT": 25_000_000,
    "MAR": 18_000_000,
    "H":    9_500_000,
}

# TODO 1: add CHH at 4_200_000
# TODO 2: print the MAR position
# TODO 3: total = sum(book.values())""")
    card(ex1, 7.3, 1.9, 5.4, 3.3, "Type it — don't read it",
         "In a fresh Colab cell, build the dictionary and do the three TODOs. A dict maps a name to a value, exactly like a two-column lookup range in Excel — ticker in column A, position in column B.\nStuck on syntax? The colon and the commas are the usual suspects. Compare with a neighbour before raising a hand.")
    tb(ex1, 0.6, 5.5, 12.1, 0.8,
       "This exact book — five positions by the time we are done — carries through Lab B, the loops talk, and the limit-breach check before lunch.",
       size=14, color=BODY)
    nb_badge(ex1, "EXERCISE", "a fresh Colab cell · 3 minutes · type it yourself")

    # Exercise after slide 10 (functions)
    ex2 = content_slide(prs, "PYTHON BOOTCAMP · EXERCISE",
                        "Exercise · 3 minutes · P&L as a function")
    code_box(ex2, 0.6, 1.9, 6.4, 3.3,
"""def position_pnl(value_start, value_end):
    return value_end - value_start

print(position_pnl(25_000_000, 25_750_000))
# 750000

# TODO: write pct_return(value_start,
#       value_end) → 0.03 for the inputs
#       above""")
    card(ex2, 7.3, 1.9, 5.4, 3.3, "One verb, named inputs, one output",
         "A function is a named formula: inputs in parentheses, result after return. The Excel analogue is a LAMBDA — except you can read it, test it, and reuse it across every workbook.\nNaming convention: a function does one thing and its name says it. position_pnl, not calc or do_stuff.")
    tb(ex2, 0.6, 5.5, 12.1, 0.8,
       "In Lab B you will write position_var(value, vol, z) and call it in a loop over the whole book — the same pattern Case 1 uses on 458 census blocks.",
       size=14, color=BODY)
    nb_badge(ex2, "EXERCISE", "a fresh Colab cell · 3 minutes · type it yourself")

    # Live demo hand-off (after "Python feeds Excel", right before Lab C)
    demo = content_slide(prs, "PYTHON BOOTCAMP · LIVE DEMO",
                         "Live demo · run along · the pivot and the lookup")
    code_box(demo, 0.6, 1.9, 6.4, 3.3,
"""# notebooks → 01b_excel_to_pandas_demo.ipynb

pivot = trades.pivot_table(
    index="ticker", columns="side",
    values="notional", aggfunc="sum")

looked_up = summary.merge(
    ref, on="ticker", how="left")""")
    card(demo, 7.3, 1.9, 5.4, 3.3, "Open it now — we build it together",
         "Open 01b_excel_to_pandas_demo.ipynb in Colab and run each cell as it comes up on screen. The data is the five-position book from Lab B plus a reference table — small enough to check every number by eye.\nWe break both moves on purpose: a missing key, then a duplicated one. Excel fails these quietly; pandas fails them loudly. Loud is what you want.")
    tb(demo, 0.6, 5.5, 12.1, 0.8,
       "When the last cell has run, Lab C hands you the same moves on 236,549 FCC rows — where nobody can check by eye. That is the whole argument of the morning.",
       size=14, color=BODY)
    nb_badge(demo, "DEMO", "01b_excel_to_pandas_demo.ipynb · run each cell as I do",
             notebook="01b_excel_to_pandas_demo.ipynb")

    # Lab B launch (after the loops slide)
    labB = content_slide(prs, "PYTHON BOOTCAMP · LAB",
                         "Lab B · 11:20 · 40 minutes — the five-position book")
    card(labB, 0.6, 1.9, 5.9, 3.4, "Do (everyone)",
         "Same notebook, section Lab B — run each cell; do every # TODO beneath it.\nA five-position book: write position_var(value, vol, z), loop it over the book, sum a total, and flag any VaR above $500,000.\nCheck-out: one sentence — what would have been painful in Excel?")
    card(labB, 6.8, 1.9, 5.9, 3.4, "Fast lane (optional)",
         "⚡ sort the book by VaR without pandas; rewrite the loop as a dict comprehension; extend position_pnl with a stop_loss threshold.\nThen help a neighbour. Hands off their keyboard.")
    tb(labB, 0.6, 5.5, 12.1, 0.8,
       "Checkpoint, 12:00 — the book loop prints five P&L lines. Lunch.",
       size=14, bold=True, color=NAVY)
    nb_badge(labB, "LAB B", "01_python_bootcamp.ipynb · section Lab B · 40 min",
             notebook="01_python_bootcamp.ipynb")

    # Lab C launch (after the 01b demo)
    labC = content_slide(prs, "PYTHON BOOTCAMP · LAB",
                         "Lab C · 1:30 · 40 minutes — pandas on the FCC data")
    card(labC, 0.6, 1.9, 5.9, 3.4, "Do (everyone)",
         "Same notebook, section Lab C. Load the availability file, filter it, pivot providers by technology, clean the brand names (21 raw spellings → 7 providers), merge the technology codes, export bootcamp_sample.xlsx.\nCheck-out: one sentence — what would have been painful in Excel?")
    card(labC, 6.8, 1.9, 5.9, 3.4, "Fast lane (optional)",
         "⚡ a bar chart of offers per technology; a served(df, down, up) function; the one-expression rebuild of the whole pipeline.\nThen help a neighbour. Hands off their keyboard.")
    tb(labC, 0.6, 5.5, 12.1, 0.8,
       "Checkpoint, 2:10 — the brand-name pivot shows 7 providers, not 21, and bootcamp_sample.xlsx is downloaded. Then a short break.",
       size=14, bold=True, color=NAVY)
    nb_badge(labC, "LAB C", "01_python_bootcamp.ipynb · section Lab C · 40 min",
             notebook="01_python_bootcamp.ipynb")

    # ---------------- break & lunch demarcation slides
    brk1 = break_slide(prs, "DAY 1 · 10:30", "Break",
        "Back at 10:50 — Python II: collections and functions",
        "Before you step out: your VaR runs, and you re-ran it at $40M.\nNot there yet? This is the moment to grab me or a neighbour.")
    lunch1 = break_slide(prs, "DAY 1 · 12:00", "Lunch",
        "Back at 1:00 — the pandas talk, a live demo, then Lab C",
        "Checkpoint before you go: the book loop prints five P&L lines.\nThis afternoon: pandas on 236,549 FCC rows, data management, and Case 1 end-to-end.")
    brk2 = break_slide(prs, "DAY 1 · 2:10", "Break",
        "Back at 2:25 — data management, then straight into Case 1",
        "The bootcamp is done: three talks, three labs, bootcamp_sample.xlsx on your machine.\nAfter the break the afternoon runs uninterrupted — a 20-minute data-management talk, then Case 1, the main event of the day.")

    # Day-2 preview (before resources)
    prev = content_slide(prs, "WRAP-UP", "Tomorrow · Day 2 — the AI day")
    card(prev, 0.6, 1.9, 5.9, 2.2, "Morning · how LLMs work",
         "Next-token prediction, training stages, tool use and agents — then context engineering. Everything you measure in Case 2.")
    card(prev, 6.8, 1.9, 5.9, 2.2, "Midday · Case 2",
         "A news screen for Marriott: lexicon baseline vs an LLM classifier vs your own labels. The deliverable is a number and the list of cases where the model was wrong.")
    card(prev, 0.6, 4.3, 5.9, 2.2, "Afternoon · extension sprint",
         "Teams of five — you plus UChicago tech majors. Extend your Case 2 pipeline in a 45-minute sprint; two-minute demos; prizes. Day ends at 3:00.")
    card(prev, 6.8, 4.3, 5.9, 2.2, "What to bring",
         "This laptop, your Colab account, and your Case 1 notebook — tomorrow's case reuses today's push/CI/test loop on a new problem. Same room, 9:00 sharp.")

    # ---- section dividers: goals line reformat
    divider_goals(S[5], "eight constructs",
        "Read and write the eight constructs the two cases use · Translate any Excel operation you know into one line of pandas")
    divider_goals(S[17], "Data hygiene is analysis",
        "Treat data hygiene as analysis, not overhead · Name the exact places Excel breaks — and where it is still the right tool")
    divider_goals(S[27], "northern-Illinois broadband operator",
        "Answer a real sizing question — where should Lakeshore expand, and how big is it · Do the analysis in Python, deliver it in Excel")
    divider_goals(S[66], "machine that is not yours",
        "Explain repo, commit, branch and pull request · Read a CI file line by line · Push code that ran — failed, then passed — on a machine that is not yours")
    divider_goals(S[75], "What to leave today with",
        "Know what to take home from today · Know what to practise before tomorrow")

    # ---------------- final order
    # Strictly linear run of show: each talk is followed immediately by its
    # own lab-launch slide, and every break/lunch is its own slide — the
    # deck IS the agenda, no table needed alongside.
    order = (S[0:7] + [var_s] + [S[7]] + [S[16]] + [brk1] +
             [S[8]] + [ex1] + [S[9]] + [ex2] +
             [S[10]] + [labB] + [lunch1] + S[11:16] + [demo] + [labC] +
             [brk2] + S[17:31] +
             S[66:77] + [prev] + S[77:79])
    rebuild_order(prs, order)
    renumber_footers(prs)
    strip_notes(prs)
    prs.save(DAY1)
    print(f"Day 1 deck: {len(order)} slides")


# ================================================================ DAY 2

def build_day2():
    shutil.copy(SRC, DAY2)
    prs = Presentation(DAY2)
    S = list(prs.slides)

    # ---- slide 1 (title)
    replace_sub(S[0], "A one-day workshop for finance students",
                "Day 2 of 2 · AI for finance — joint with UChicago tech majors")
    replace_sub(S[0], "9:00 – 5:00 · laptops required",
                "9:00 – 3:00 · laptops required")

    # ---- slide 32 (divider → SECTION 2)
    replace_sub(S[31], "SECTION 4", "SECTION 1")
    replace_sub(S[31], "20-minute talk (1:00–1:20)", "25-minute talk (9:15–9:40)")

    # ---- slide 36: temperature → determinism, harness-aware
    set_box_text(S[35], "Temperature",
        "Determinism — will it answer the same twice?")
    # left box: proper bullets, RAG line lives here
    set_box_bullets(S[35], "Everything the model can 'see' for this answer", [
        "Everything the model can 'see' for this answer: your instructions, the data you pasted, the conversation so far. Hundreds of thousands of tokens — hundreds of pages — but finite, and attention gets thinner as it fills.",
        "Nothing persists between calls unless you send it again. 'It remembered' means the product re-sent the history.",
        "Retrieval (RAG) is just: search your documents first, paste the relevant passages into the prompt, then ask.",
    ])
    # right box: proper bullets, determinism only (no RAG)
    set_box_bullets(S[35], "A dial on randomness when sampling the next token.", [
        "In the chat products you use daily, sampling is tuned for you — and as of 2026 the API is the same: current models manage their own sampling, and the SDK no longer takes a temperature argument.",
        "So determinism is not a dial you turn — it is a property you measure. The same headline can come back 'neutral' now and 'negative' in a minute.",
        "Run the classification twice and report the disagreement rate: that number is the error bar on your accuracy.",
    ])

    # ---- slide 37: model/product/API — SDK 1.0 dropped the temperature argument
    replace_run(S[36], "Your system prompt, your temperature, your data handling",
        "Your system prompt, your data handling")
    replace_sub(S[36], "max_tokens=300, temperature=0,", "max_tokens=300,")

    # ---- slide 38: failure modes, softened + current
    replace_sub(S[37], "Four ways they break — memorise these",
        "Failure modes to keep in mind")
    set_box_text(S[37], "At temperature above 0, the same headline",
        "Run to run, the same headline can be 'neutral' now and 'negative' in a minute — and there is no temperature dial to turn off. Run twice; report the disagreement rate — that number is the error bar on your accuracy.")
    set_box_text(S[37], "Leading questions get led answers.",
        "'Isn't this headline clearly negative?' → 'Yes, clearly negative.' Still measurable in today's frontier models. Ask neutrally — or better, don't reveal your view.")

    # ---- slide 39: strong/weak, model-vs-harness framing
    replace_sub(S[38], "Where they are strong, where they are weak",
        "Strong vs weak — the model, and the harness around it")
    set_box_text(S[38], "Strong — use them here", "Strong — the model itself")
    set_box_text(S[38], "Weak — use a different tool",
        "Weak — know which layer fails")
    set_box_bullets(S[38], "Classification: is this headline about the company?", [
        ("Classification, extraction, reformatting: ",
         "is this headline about the company? covenant terms into a table; 40 pages → one; JSON ↔ prose."),
        ("First drafts: ",
         "code, and explanations of errors — then you read every line."),
        ("In a harness: ",
         "Claude Code, Cowork, a chat app with tools — the model runs Python, searches, reads your files. Arithmetic and fact lookup stop being weaknesses: the tool computes, the model reads the answer."),
    ])
    set_box_bullets(S[38], "Arithmetic over many numbers: let Python add", [
        ("A bare API call has no tools: ",
         "exact arithmetic and fact recall must come from your code and your data. It will answer anyway — that's the trap. Today's classifier is this case."),
        ("A harnessed model can still skip the tool ",
         "and answer from stale memory — ask to see the source before you trust a number."),
        ("Agents compound errors: ",
         "95% per step is ~60% over ten steps. A smooth demo is not evidence."),
        ("Consistency at scale needs a rubric and a gold set ",
         "— in any harness. That is why today builds one."),
    ])

    # ---- slide 42 (end of AI Systems talk: preview of the Case 2 pipeline)
    replace_sub(S[41],
        "Bridge: this afternoon we use one as a classifier — and measure it",
        "The rest of today: use one as a classifier — and measure it")
    replace_run(S[41], "at temperature 0, with a versioned prompt",
        "with a versioned prompt")

    # ---- slide 43 (divider → SECTION 2, retitle; must=False once SRC rebuilt)
    replace_sub(S[42], "SECTION 5", "SECTION 2")
    replace_sub(S[42], "Prompt engineering", "Context engineering", must=False)
    replace_sub(S[42], "20-minute talk (1:20–1:40)", "15-minute talk (9:40–9:55)")

    # ---- slide 52 (secrets — kept in the new Section 2; fix kicker)
    replace_sub(S[51], "PROMPT ENGINEERING", "CONTEXT ENGINEERING", must=False)
    replace_sub(S[51], 'userdata.get("ANTHROPIC_API_KEY")',
        'userdata.get("ANTHROPIC_API_KEY").strip()   # strip paste newlines', must=False)

    # ---- slide 53 (divider → SECTION 4, new times)
    replace_sub(S[52], "SECTION 6", "SECTION 3")
    replace_sub(S[52],
        "5-minute brief · Part A individual 1:45–2:30 · Part B group 3:05–3:55 · Part C 3:55–4:20",
        "10-minute brief · Part A 10:20–11:10 · Part B 11:10–12:00, resumes 1:10–1:40 · debrief 1:40")

    # ---- slide 54
    replace_sub(S[53], "That separation is what makes it testable at 3:55.",
        "That separation is what makes it testable in CI — which the finance cohort set up yesterday.")
    nb_badge(S[53], "CASE", "03_news_dashboard_STARTER.ipynb · Steps 1–5 only before lunch · with your team of five",
             notebook="03_news_dashboard_STARTER.ipynb")
    # step-range labels under the pipeline boxes: this slot covers steps 1–5
    # (boxes 1–4); evaluation + dashboard (boxes 5–6) come after lunch
    tb(S[53], 0.6, 3.56, 7.97, 0.3, "STEPS 1–5 · THIS SLOT · DONE BY 12:00",
       size=10.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    tb(S[53], 8.92, 3.56, 3.81, 0.3, "STEPS 6–7 · AFTER LUNCH · 1:10",
       size=10.5, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

    # ---- slide 55 three parts, two tracks
    replace_sub(S[54], "A · Individual · 1:45–2:30", "A · Individual · 10:20–11:10")
    replace_sub(S[54], "B · Group · 3:05–3:55", "B · Group · 11:10–12:00 · resumes 1:10")
    replace_sub(S[54], "C · Group · 3:55–4:20", "C · Fast lanes · from 11:10")
    set_box_text(S[54], "Run the dashboard (Colab port proxy or local).",
        "Tech students: these start this morning, as soon as Part B labelling starts — don't save them for the afternoon.\nBatch the API calls; swap in a cheaper model and compare accuracy; add a cost-per-headline tracker; deploy the dashboard (share.streamlit.io).\nEvery fast lane ends in a number next to the baseline — same rules as everyone else.")
    replace_sub(S[54], "Track A only: one prompt iteration, one logged score.",
        "Everyone ships PROMPT_V3: one change, logged against V2 on the same gold set — kept even if it scores worse.")
    replace_run(S[54], "Track A — API key available", "Track A — your own API key", must=False)
    replace_run(S[54], "One workshop key with a spend cap serves the class for cents.",
        "Your university's AI access covers it — the whole case costs cents.", must=False)
    replace_run(S[54], "Real model, temperature 0, prompt iteration measurable.",
        "Real model, prompt iteration measurable.")

    # ---- slide 56
    replace_sub(S[55],
        "Show-and-tell at 4:20, two minutes per group: the dashboard, the accuracy table, one disagreement you found interesting, your CI status.",
        "Show-and-tell at 1:40, two minutes per group: the accuracy table, one disagreement you found interesting, and your V2 → V3 score change.")

    # ---- slide 56: a guide to what good looks like, not a grading scheme
    replace_sub(S[55], "What to produce, and how it is assessed",
        "What to produce — and what good looks like", must=False)
    tblf = find_table(S[55])
    if tblf is not None:
        for cell in tblf.rows[0].cells:
            if cell.text == "What a 4 looks like":
                cell.text_frame.paragraphs[0].runs[0].text = "What good looks like"
        for row in tblf.rows:
            for cell in row.cells:
                if "≥20 headlines" in cell.text:
                    runs = [r for p in cell.text_frame.paragraphs for r in p.runs]
                    runs[0].text = ("15 headlines (10 common + 5 team-unique) labelled independently "
                                    "by all five of you; agreement rate reported and interpreted as a ceiling")
                    for r in runs[1:]:
                        r.text = ""
    # Class gold sheet link — stays projected during team work (no board in the room).
    tb(S[55], 0.6, 6.5, 12.1, 0.35,
       "CLASS GOLD SHEET — one teammate enters your 15 agreed labels:  [paste shared-sheet link here before class]",
       size=13, bold=True, color=AMBER)

    # ---- slide 57 (divider → SECTION 4)
    replace_sub(S[56], "SECTION 7", "SECTION 4")
    replace_sub(S[56], "20-minute talk (2:45–3:05) · leads into Part B",
        "25-minute talk (12:45–1:10) · leads into Part B's finish")

    # ---- slide 64 title (softer)
    replace_sub(S[63], "Read the disagreement list — it is the actual deliverable",
        "Read the disagreement list", must=False)

    # ---- slide 65 'Run it twice' card: no temperature dial to set
    replace_run(S[64], "Temperature 0, same prompt, same input",
        "Same prompt, same input")

    # ---- slide 66
    replace_sub(S[65],
        "Part B starts now: 03_news_dashboard — Steps 4 to 6. Labels first, silently.",
        "Part B resumes now — finish Steps 6–7: measure, read every disagreement, ship PROMPT_V3 with both scores logged.")
    nb_badge(S[65], "CASE", "reopen 03_news_dashboard_STARTER.ipynb · Steps 6–7 · done by 1:40",
             notebook="03_news_dashboard_STARTER.ipynb")

    # ---- slide 76 (divider → SECTION 6)
    replace_sub(S[75], "SECTION 9", "SECTION 6")
    replace_sub(S[75], "10 minutes (4:50–5:00)", "Awards & wrap (2:55–3:00)")

    # ---- slide 77 takeaways: clear title + context-eng habit replaces prompt-log lore
    replace_sub(S[76], "Monday morning", "Takeaways — put them to work Monday",
        must=False)
    replace_sub(S[76], "Keep a prompt log", "Start an instruction file")
    set_box_text(S[76], "Version, change, score. Even for chat use.",
        "In your first repo, write a CLAUDE.md / AGENTS.md: what the project is, how to run it, what not to touch. Edit it whenever the agent annoys you — that habit compounds.")

    # ---- slide 79 (survey goes out after the workshop — no slide reference)
    replace_sub(S[78],
        "Everything from today lives in the repository: notebooks, data, the app, this deck, and the syllabus.",
        "Everything from both days lives in the repository: notebooks, data, the app, both decks, and the syllabus.")
    set_box_text(S[78], "Feedback form",
        "Everything from both days lives in the repository: notebooks, data, the app, both decks, and the syllabus.")

    # ---------------- new slides
    w1 = content_slide(prs, "WELCOME", "Two cohorts, one room")
    card(w1, 0.6, 1.9, 5.9, 2.2, "Finance students",
         "You built Case 1 yesterday — you know the data, the case, and the business question. Today you own the judgement calls: labels, materiality, what a partner would believe.")
    card(w1, 6.8, 1.9, 5.9, 2.2, "Tech students",
         "You know code; the finance judgement sits across the table. Prep pack covered Case 1 and how a PE analyst reads news — your fast lanes start at Case 2. Yesterday's details: next slide.")
    card(w1, 0.6, 4.3, 5.9, 2.2, "Teams of five — announced now",
         "Ten teams, 2–3 finance + 2–3 tech each, fixed for the whole day: Case 2 and the extension sprint. Sit with your team before 9:10.")
    card(w1, 6.8, 4.3, 5.9, 2.2, "The contract",
         "Finance drives judgement, tech drives code — and each explains their half out loud. A demo nobody on the team can explain scores zero on teamwork.")

    w2 = content_slide(prs, "WELCOME", "The arc of the day")
    chain(w2, ["Day 1\nrecap", "How AI\nworks", "Context\nengineering",
               "Case 2\nnews screen", "Evaluating\nAI outputs",
               "Case 2\nfinish", "Extension\nsprint", "Demos &\nawards"],
          y=2.2, x0=0.65, box_w=1.28, gap=0.22)
    tb(w2, 0.6, 4.0, 12.1, 2.6,
       "Morning: how LLMs actually work — then Case 2 begins: a news screen for Marriott with a lexicon baseline, an LLM classifier, and your own gold labels (working lunch 12:00–12:45 — pick your sprint direction over food).\n \nEarly afternoon: evaluation as a discipline, then finish Case 2 — every team ships a measured PROMPT_V3 by 1:40.\n \nFrom 1:50: the extension sprint. Extend your own pipeline — 45 minutes to build, two-minute demos, awards at 3:00.",
       size=15, color=BODY, line_spacing=1.2)

    b1 = content_slide(prs, "WELCOME", "Yesterday, in five minutes")
    card(b1, 0.6, 1.9, 5.9, 2.2, "The case",
         "Lakeshore Broadband, a fictional Illinois operator: 236,549 FCC availability rows → 458 whitespace census blocks → 42,220 units → Cook County $5.1M revenue, DuPage the best payback at 4.1 years. Built in pandas, delivered as a 3-sheet Excel workbook (county ranking · whitespace blocks · assumptions).")
    card(b1, 6.8, 1.9, 5.9, 2.2, "The constructs",
         "Variables, lists, dicts, functions, loops, DataFrames, filtering, groupby/merge — plus reading errors from the bottom and assumptions in CAPITALS at the top of the notebook.")
    card(b1, 0.6, 4.3, 5.9, 2.2, "The delivery rule",
         "Python is the engine; Excel is the delivery format. Every analysis ends in .to_excel() with an assumptions sheet written from code — auditability for free. Partners never see a notebook.")
    card(b1, 6.8, 4.3, 5.9, 2.2, "Already done: Git & CI",
         "Repos, commits, PRs — and CI: yesterday every group pushed the case's library and tests (case1/) through GitHub Actions, broke it on purpose, and fixed it. Today we use it, not teach it. The workflow lives at .github/workflows/ci.yml.")

    train4 = content_slide(prs, "AI SYSTEMS",
                           "Four stages of training — and what each one buys you")
    stage_boxes = [
        ("1 · Pre-training", "predict the next token on most of the public internet"),
        ("2 · Instruction tuning", "learn to follow 'Classify this headline' rather than continue it"),
        ("3 · Preference tuning (RLHF)", "prefer answers humans rate as helpful, harmless, honest"),
        ("4 · Reasoning RL", "learn, on checkable problems (math, code, tools), to think in steps before answering"),
    ]
    bx, bw, bgap, by, bh = 0.6, 2.62, 0.42, 2.0, 1.55
    for i, (hd, bd) in enumerate(stage_boxes):
        x = bx + i * (bw + bgap)
        if i == 3:
            rrect(train4, x, by, bw, bh, fill="FDF3E3", line=AMBER)
            tb(train4, x - 0.05, by - 0.38, bw + 0.4, 0.3,
               "newest — the 2025-era reasoning models", size=10, bold=True,
               color=AMBER)
        else:
            rrect(train4, x, by, bw, bh)
        tb(train4, x + 0.15, by + 0.12, bw - 0.3, 0.35, hd, size=12.5,
           bold=True, color=NAVY)
        tb(train4, x + 0.15, by + 0.5, bw - 0.3, bh - 0.62, bd, size=11,
           color=BODY, line_spacing=1.1)
        if i < 3:
            tb(train4, x + bw + 0.03, by + 0.5, bgap, 0.5, "›", size=20,
               bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    tb(train4, 0.6, 3.82, 12.1, 0.3, "WHAT EACH STAGE BUYS YOU",
       size=10.5, bold=True, color=GREY)
    buys = [
        "Knowledge and fluency — and the tendency to produce plausible text whether or not it is true.",
        "The format you want: ask for JSON and you get JSON — trained to honour structure.",
        "Agreeableness — including sycophancy: agreeing with your framing instead of correcting it.",
        "Step-by-step working on hard problems — better at maths, code, logic; slower and pricier per call.",
    ]
    for i, txt in enumerate(buys):
        x = bx + i * (bw + bgap)
        rrect(train4, x, 4.15, bw, 1.42, fill="FFFFFF", line=STEEL)
        tb(train4, x + 0.13, 4.27, bw - 0.26, 1.2, txt, size=10.5,
           color=BODY, line_spacing=1.12)
    rrect(train4, 0.6, 5.8, 12.1, 0.85, fill="FDF3E3", line=AMBER)
    tb(train4, 0.85, 5.96, 11.6, 0.6,
       "None of the stages include your data. The model does not know Lakeshore Broadband, this week's Marriott headlines, or your fund's thesis — unless you put them in the prompt.",
       size=13, bold=True, color=NAVY, line_spacing=1.15)

    tools = content_slide(prs, "AI SYSTEMS", "Tool use and agents — from one call to a system")
    card(tools, 0.6, 1.9, 5.9, 2.2, "Tool use",
         "The model emits a structured call — run_python(...), search(...) — your code executes it and returns the result, and the model continues with that result in context. Arithmetic and fact lookup stop being weaknesses the moment the model can call a tool.")
    card(tools, 6.8, 1.9, 5.9, 2.2, "Agents",
         "A loop: model → tool → result → model, until the task is done. Coding agents, deep-research products and computer-use systems are this loop with different tools. Errors compound across steps — a 95%-per-step agent is a 60% ten-step agent.")
    card(tools, 0.6, 4.3, 12.1, 2.0, "What does not change",
         "Each step is still next-token prediction; the loop still needs a human-set goal, closed outputs where possible, and an evaluation before trust. This afternoon's extension sprint is exactly this pattern — and it is graded on the evaluation, not the demo gloss.")

    # ---------------- SECTION 2 · Context engineering (complete 2026 rebuild;
    # replaces the SRC prompt-engineering slides S[43:48] + S[50], which stay
    # in the SRC deck but are no longer used here)
    c1 = content_slide(prs, "CONTEXT ENGINEERING",
                       "The model sees one thing: its context window")
    card(c1, 0.6, 1.9, 5.9, 2.2, "Everything arrives as one window",
         "Instructions, your data, examples, tool definitions, tool results, conversation history — one token sequence. The model has no other channel into your task. What is not in the window does not exist.")
    card(c1, 6.8, 1.9, 5.9, 2.2, "Attention is finite",
         "More context is not more safety: recall measurably degrades as windows fill with low-signal text ('context rot'). The goal is the smallest set of high-signal tokens that gets the outcome.")
    card(c1, 0.6, 4.3, 5.9, 2.2, "Context engineering, defined",
         "Deciding what earns a place in the window at each step: what goes in, what stays out, what gets summarised, what moves to a file the model can read back later.")
    card(c1, 6.8, 4.3, 5.9, 2.2, "Why 'prompt wording' faded",
         "2023 was magic phrasings and trick templates. Today's models are post-trained to follow plain instructions — what you include now matters far more than how you phrase it.")

    c2 = content_slide(prs, "CONTEXT ENGINEERING",
                       "What goes in the window — and what stays out")
    card(c2, 0.6, 1.9, 5.9, 2.2, "In: role, rules, perspective",
         "Who is answering, for whom — 'from the perspective of an INVESTOR' turns routine hotel-opening news from positive to neutral. State the task in one verb and the rule for edge cases.")
    card(c2, 6.8, 1.9, 5.9, 2.2, "In: this item's data, delimited",
         "The headline, the source, one line that disambiguates Marriott-the-company from the university library that shares the name. Training gave the model no knowledge of your data — the window is where it learns of it.")
    card(c2, 0.6, 4.3, 5.9, 2.2, "Out: what another layer does better",
         "Anything you can compute (let Python), anything you can look up deterministically (merge it in as data), anything secret — keys and MNPI never enter a prompt.")
    card(c2, 6.8, 4.3, 5.9, 2.2, "Out: padding",
         "Ten pages of guidelines dilute the three rules that matter. Short, precise context outperforms sprawling context — in prompts, and in agent instruction files too.")

    c3 = content_slide(prs, "CONTEXT ENGINEERING", "Examples beat instructions")
    card(c3, 0.6, 1.9, 5.9, 4.4, "Rules struggle with edge cases",
         "'Negative means bad for the investor, unless the cost cut signals discipline rather than distress, unless…' — every rule you write breeds exceptions.\n\nA few well-chosen labelled examples calibrate the model faster than any rule: \"'Company cuts 200 corporate roles in efficiency push' → neutral (cost discipline, not distress).\"")
    card(c3, 6.8, 1.9, 5.9, 4.4, "Curate, don't accumulate",
         "A small, diverse set of canonical hard cases beats a long list of near-duplicates.\n\nThe risk: examples transmit your labelling biases, including ones you haven't noticed. Keep them out of your evaluation set, or you are grading the model on questions it has seen.")

    c4 = content_slide(prs, "CONTEXT ENGINEERING",
                       "Structured output: design the schema — the API enforces it")
    card(c4, 0.6, 1.9, 5.9, 2.2, "2023: beg and re-ask",
         "'Return ONLY a JSON object' — then parse, fail on the chatty preamble, retry. A real failure mode then; solved infrastructure now.")
    card(c4, 6.8, 1.9, 5.9, 2.2, "2026: schema-enforced decoding",
         "Modern APIs accept a JSON schema and guarantee output that parses against it. The begging is gone — the design work is what remains, and it is yours.")
    card(c4, 0.6, 4.3, 5.9, 2.2, "Design: closed labels, with an exit",
         "sentiment ∈ positive | negative | neutral; category from a fixed list that includes 'other'. Without an exit, edge cases get forced into the nearest wrong label — silently.")
    card(c4, 6.8, 4.3, 5.9, 2.2, "Design: a rationale field",
         "~20 tokens that buy you a reason to read when a label surprises you — and the raw material for the disagreement list you'll study after lunch. Today's notebook does exactly this.")

    c5 = content_slide(prs, "CONTEXT ENGINEERING",
                       "Guiding an agent: instructions that persist")
    card(c5, 0.6, 1.9, 5.9, 2.2, "The standing instruction file",
         "Agent harnesses read a project file — CLAUDE.md / AGENTS.md, by now a near-universal convention — on every task: what this project is, how to run the tests, what not to touch. Write it like onboarding notes for a sharp new hire.")
    card(c5, 6.8, 1.9, 5.9, 2.2, "Short beats long",
         "A focused few-hundred-token file outperforms a sprawling one that tries to cover everything — the same attention budget applies to instructions as to data.")
    card(c5, 0.6, 4.3, 5.9, 2.2, "It is a living document",
         "When the agent does something annoying, you edit the file so it never does it again. That feedback loop is the 2026 replacement for prompt tinkering.")
    card(c5, 6.8, 4.3, 5.9, 2.2, "Tools are context too",
         "An agent is only as good as the descriptions of the tools it may call. A vague tool description produces vague tool use.")

    c6 = content_slide(prs, "CONTEXT ENGINEERING",
                       "Long tasks: write it down, compress it, isolate it")
    card(c6, 0.6, 1.9, 5.9, 2.2, "Write",
         "Move state out of the window: a plan file, a scratchpad, a task list. The agent re-reads what it needs instead of carrying everything everywhere.")
    card(c6, 6.8, 1.9, 5.9, 2.2, "Compress",
         "Summarise finished sub-tasks; truncate verbose tool output; replace a long history with a short recap ('compaction'). Keep signal, drop bulk.")
    card(c6, 0.6, 4.3, 5.9, 2.2, "Isolate",
         "Give a sub-task its own clean window — a sub-agent — and take back only the result. Ten clean small windows beat one overflowing large one.")
    card(c6, 6.8, 4.3, 5.9, 2.2, "You will feel this in the sprint",
         "When your AI assistant starts 'forgetting', its window is full. Don't push on: ask for a written summary, start fresh, paste the summary.")

    c7 = content_slide(prs, "CONTEXT ENGINEERING",
                       "One change, one measurement — the habit that replaces prompt lore")
    card(c7, 0.6, 1.9, 5.9, 2.2, "Change one thing",
         "An instruction, an example, the schema, the model — one change at a time, re-scored on the same labelled set, logged next to the old score. Keep every version: PROMPT_V2 stays when V3 ships.")
    card(c7, 6.8, 1.9, 5.9, 2.2, "This is how the field works now",
         "Eval-driven development: labs and serious teams improve these systems by measuring against fixed test sets, not by collecting magic phrases.")
    card(c7, 0.6, 4.3, 5.9, 2.2, "Today, concretely",
         "This afternoon every team ships PROMPT_V3: one change against V2, scored on your own gold labels — and the score is logged even if it got worse. Especially if it got worse.")
    card(c7, 6.8, 4.3, 5.9, 2.2, "The one-line summary",
         "A context change without a measurement is a vibe. The whole of Section 4 — and most of what makes AI usable at a fund — is turning vibes into numbers.")

    hack_div = divider_slide(prs, "SECTION 5", "Extension sprint",
        "Brief 1:40 · build 1:50–2:35 · demos 2:35–2:55 · awards 3:00",
        "Extend your Case 2 pipeline in one direction · Build the smallest thing that works, as a team of five · Measure it honestly · Demo it in two minutes")

    h1 = content_slide(prs, "EXTENSION SPRINT",
                       "Four directions — lock in your team's pick")
    card(h1, 0.6, 1.9, 5.9, 2.2, "1 · Watchlist",
         "One company becomes three. A rollup digest ranked by materiality to an investor, a what-moved-today view — and the name-collision problem, now ×3.")
    card(h1, 6.8, 1.9, 5.9, 2.2, "2 · Cost-aware classifier",
         "Batch the API calls, add a cost-per-headline tracker, swap in a cheaper model and compare. The deliverable is accuracy-per-dollar, not just accuracy.")
    card(h1, 0.6, 4.3, 5.9, 2.2, "3 · Analyst digest",
         "Headlines → an analyst-ready digest: extraction (checkable against the source) kept separate from interpretation (labelled as such) — plus three questions for management.")
    card(h1, 6.8, 4.3, 5.9, 2.2, "4 · Dashboard",
         "Deploy the Streamlit dashboard (share.streamlit.io) with your team's PROMPT_V3 and the evaluation table wired in — a URL anyone in the room can open.")

    h2 = content_slide(prs, "EXTENSION SPRINT", "Rules, rubric, and the demo")
    card(h2, 0.6, 1.9, 5.9, 2.2, "Rubric · 100 points",
         "30 — it runs, live, start to finish.\n25 — honest evaluation: ≥10 items checked against a human answer, failures shown.\n25 — finance insight a partner would recognise.\n20 — teamwork: both cohorts can explain the whole thing.")
    card(h2, 6.8, 1.9, 5.9, 2.2, "Rules",
         "Any tool allowed, AI assistants included. You're on your own API key — batch before you loop, prefer cheaper models to iterate. Synthetic/cached data provided for every challenge; the Wi-Fi is not your dependency.")
    card(h2, 0.6, 4.3, 5.9, 2.2, "The 2-minute demo",
         "What it does (one sentence) · live run · the evaluation table · one sentence on what you would NOT trust it for. That last sentence is required — it is where most points are won.")
    card(h2, 6.8, 4.3, 5.9, 2.2, "Checkpoints",
         "2:10 — scope-cut call: we walk the room and cut each build to what can demo.\n2:25 — feature freeze: stop adding, start rehearsing the run.")
    nb_badge(h2, "SPRINT", "your team's Case 2 notebook + anything you need · build starts 1:50")

    # ---- section dividers: goals line reformat
    divider_goals(S[31], "next-token prediction in one sentence",
        "Explain next-token prediction in one sentence · Name the failure modes that matter · Predict what the model — and the harness around it — will do well or badly, before you try")
    divider_goals(S[42], "vague instruction into a constrained",
        "Decide what earns a place in the model's context window · Guide an agent with instructions, examples and tools — not magic words · Change one thing, measure, keep what wins")
    divider_goals(S[52], "repeatable read of the news",
        "Build a repeatable read of the news about Marriott · Establish, with numbers, when that read can be trusted")
    divider_goals(S[56], "read its results honestly",
        "Design a small evaluation and read its results honestly · Report to a partner: n, metric, baseline, three failures — and the decision the number supports")
    divider_goals(S[75], "Monday morning",
        "Know what to do on Monday morning · Know where to go next")

    # ---------------- break & lunch demarcation slides
    brk_d2 = break_slide(prs, "DAY 2 · 9:55", "Break",
        "Back at 10:10 — Case 2 begins",
        "A news screen for an investment target: fetch → baseline → LLM → evaluate → dashboard.\nSit with your team of five before we restart.")
    lunch2 = break_slide(prs, "DAY 2 · 12:00", "Working lunch",
        "Back at 12:45 — evaluating AI outputs, then Case 2 finishes",
        "Over food, with your team: pick your extension-sprint direction — watchlist · cost-aware classifier · analyst digest · dashboard (full briefs at 1:40; you lock in your pick then).\nPart B of Case 2 resumes at 1:10; every team ships a measured PROMPT_V3 by 1:40.")

    # ---------------- final order (drop 49-50, i.e. S[48], S[49])
    order = ([S[0], w1, b1, w2] +
             S[31:34] + [train4] + S[35:37] + [tools] + S[37:42] +  # 32-34, new training slide, 36-37, tools, 38-42
             [S[42], c1, c2, c3, c4, c5, c6, c7, S[51]] +  # Section 2: divider, 7 new context-eng slides, secrets
             [brk_d2] +                                    # break
             S[52:56] + [lunch2] + S[56:66] +             # 53-56, lunch, 57-66
             [hack_div, h1, h2] +
             S[75:79])                                    # 76-79
    rebuild_order(prs, order)
    renumber_footers(prs)
    strip_notes(prs)
    prs.save(DAY2)
    print(f"Day 2 deck: {len(order)} slides")


if __name__ == "__main__":
    build_day1()
    build_day2()
