"""Render the prep guides (HTML -> PDF) with Chromium via playwright.

Usage:  python build/render_prep_guide.py
Needs:  pip install playwright && playwright install chromium
Renders docs/day_1_prep_guide.html -> docs/day_1_prep_guide.pdf  (finance cohort, both days)
        docs/day_2_prep_guide.html -> docs/day_2_prep_guide.pdf  (tech cohort, Day 2 only)
"""
import asyncio, pathlib
from playwright.async_api import async_playwright

def foot(label):
    return ('<div style="width:100%;font-family:Arial,sans-serif;font-size:7.5pt;color:#6B7280;'
            'padding:0 0.55in;display:flex;justify-content:space-between;">'
            f'<span>From Excel to Python and AI · University of Chicago · {label}</span>'
            '<span>page <span class="pageNumber"></span> of <span class="totalPages"></span></span></div>')

def page_count(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            return None
    return len(PdfReader(path).pages)

async def render(pg, src, out, label, expect_pages=None):
    await pg.goto(f"file://{src}")
    await pg.pdf(path=out, format="Letter", print_background=True, prefer_css_page_size=True,
                 display_header_footer=True, header_template="<span></span>", footer_template=foot(label),
                 margin={"top": "0.45in", "bottom": "0.5in", "left": "0.55in", "right": "0.55in"})
    n = page_count(out)
    print(f"wrote {out}" + (f" ({n} pages)" if n else ""))
    if expect_pages and n and n != expect_pages:
        print(f"  WARNING: expected {expect_pages} pages, got {n} — content overflowed; "
              f"trim text or lower the body zoom in {pathlib.Path(src).name}")

async def main():
    D = pathlib.Path(__file__).resolve().parents[1] / "docs"
    async with async_playwright() as p:
        b = await p.chromium.launch(); pg = await b.new_page()
        await render(pg, str(D / "day_1_prep_guide.html"), str(D / "day_1_prep_guide.pdf"), "setup checklist", expect_pages=2)
        await render(pg, str(D / "day_2_prep_guide.html"), str(D / "day_2_prep_guide.pdf"), "Day 2 tech cohort checklist")
        await b.close()

asyncio.run(main())
