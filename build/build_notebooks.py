"""Build the workshop notebooks (.ipynb are JSON; no nbformat dependency needed).

Solution cells live in the private companion repo (checked out or symlinked at
``private/`` in this repo's root — see private/README.md). Without it, this
script builds the STARTER notebooks only.
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NB_DIR = ROOT / "notebooks"
NB_DIR.mkdir(exist_ok=True)

REPO_URL = "https://github.com/limorigu/uchicago-python-ai-workshop"
COLAB_BASE = "https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/"

# Private companion repo (solutions + answer keys) — invite-only for participants.
PRIVATE_DIR = ROOT / "private"
SOL_NB_DIR = PRIVATE_DIR / "notebooks"
SOLUTIONS_COLAB_BASE = "https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop-solutions/blob/main/notebooks/"

sys.path.insert(0, str(PRIVATE_DIR / "build"))
try:
    from solutions_private import SOLUTIONS
except ImportError:
    SOLUTIONS = None


def md(src):
    return {"cell_type": "markdown", "metadata": {}, "source": src.strip("\n")}


def code(src):
    return {"cell_type": "code", "metadata": {}, "execution_count": None, "outputs": [], "source": src.strip("\n")}


def sol(key):
    """Solution cell looked up in the private module; None (= reuse starter cell) without it."""
    return code(SOLUTIONS[key]) if SOLUTIONS else None


def colab_badge(name, base=COLAB_BASE):
    return md(f"[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)]({base}{name})")


def write(name, cells, out_dir=None, colab_base=COLAB_BASE):
    cells = [colab_badge(name, colab_base)] + cells
    nb = {"cells": cells, "metadata": {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
                                       "language_info": {"name": "python"}, "colab": {"provenance": []}},
          "nbformat": 4, "nbformat_minor": 5}
    d = out_dir or NB_DIR
    d.mkdir(parents=True, exist_ok=True)
    (d / name).write_text(json.dumps(nb, indent=1))
    print("wrote", (d / name).relative_to(ROOT), f"({len(cells)} cells)")


SETUP_CELL = code(f'''
# --- Setup: get the workshop files into this Colab session ------------------------------
# Downloads the workshop repository (data + helper code). Takes ~10 seconds the first time.
import os, sys
REPO_URL = "{REPO_URL}"
if not any(os.path.exists(d) for d in ["workshop/data", "data", "../data"]):
    !GIT_TERMINAL_PROMPT=0 git clone -q --depth 1 $REPO_URL workshop
DATA = next((d for d in ["workshop/data", "data", "../data"] if os.path.exists(d)), None)
if DATA is None:
    raise RuntimeError("Could not download the workshop files. Check your internet connection and "
                       "that this link opens in your browser: " + REPO_URL + " . Still stuck? Email the instructor.")
sys.path.insert(0, "workshop/app" if os.path.exists("workshop/app") else "../app")
print("Data folder:", DATA, "| files:", sorted(os.listdir(DATA)))
''')

# =====================================================================================
# 00 — SETUP CHECK (pre-work)
# =====================================================================================
write("00_setup_check.ipynb", [
    md("""
# 00 · Setup check (do this before the workshop — 10 minutes)

**Goal:** confirm you can run Python in the browser and load the workshop data. No installation needed. Both cohorts do this: finance students before Day 1, tech students before Day 2.

1. Open this notebook in **Google Colab** (File → Open notebook → GitHub, or click the Colab badge in the repo).
2. Run each cell with **Shift+Enter**. A cell has finished when the `[ ]` on its left shows a number.
3. If the last cell prints **"You're ready"**, you're done. If not, screenshot the error and email it to the instructor — we will fix it before the day.
"""),
    code("""
import sys, platform
print("Python", sys.version.split()[0], "on", platform.system())
import pandas as pd, numpy as np, matplotlib
print("pandas", pd.__version__, "| numpy", np.__version__, "| matplotlib", matplotlib.__version__)
"""),
    SETUP_CELL,
    code("""
import pandas as pd
avail = pd.read_csv(f"{DATA}/case1/bdc_il_fixed_availability_sample.csv.gz")
print(f"Loaded {len(avail):,} rows x {avail.shape[1]} columns")
avail.head()
"""),
    md("""
### One thing to notice before you leave
Excel stops at **1,048,576 rows**. The sample you just loaded is a small classroom extract; the real FCC availability data is *hundreds of millions* of rows. That single fact is most of the argument for Day 1.
"""),
    code("""
assert len(avail) > 100_000, "unexpected file size — tell the instructor"
print("You're ready ✔  See you at 9:00.")
"""),
])

# =====================================================================================
# 01 — PYTHON BOOTCAMP: LABS A, B, C (Day 1, three lab blocks)
# =====================================================================================
write("01_python_bootcamp.ipynb", [
    md("""
# 01 · Python bootcamp — Labs A, B, C (Day 1)

One notebook, three labs, matched to the three morning/early-afternoon talks:

| Lab | When | After the talk on | Running example |
|---|---|---|---|
| **A** | 9:55–10:30 | values, variables, types | a 1-day VaR calculation |
| **B** | 11:20–12:00 | lists, dicts, functions, loops | a five-position book |
| **C** | 1:30–2:10 | the DataFrame, Excel→pandas | the FCC broadband data |

**How to work:** run each cell, read the output, then do the `# TODO` in the cell below it. Cells marked **⚡ Fast lane** are optional extras for people who have coded before — skip them if you are new.

**Ask for help early.** Raise a hand; an error message is a normal part of programming, not a sign you did it wrong. Read error messages from the **bottom line up**.
"""),
    SETUP_CELL,
    md("""
---
# Lab A · Values, variables, types (9:55–10:30)

Excel cells hold values. Python **variables** hold values *and have names you choose*. The name is the documentation.

In the talk you saw a parametric **Value-at-Risk** built from four variables. Here you rebuild it yourself, then do what the PM always does: change the assumptions.
"""),
    code("""
# A parametric 1-day VaR, from nothing but variables.
position_value = 25_000_000      # a number (int) — $25M position. Underscores are thousands separators.
daily_vol = 0.018                # a decimal (float) — 1.8% daily volatility
z_99 = 2.33                      # 99% confidence z-score
desk = "Consumer & Lodging"      # text (str)
is_liquid = True                 # a boolean

var_1d = position_value * daily_vol * z_99
print(f"{desk}: 1-day 99% VaR = ${var_1d:,.0f}")
print(type(position_value), type(daily_vol), type(desk), type(is_liquid))
"""),
    code("""
# TODO 1: the PM asks for 95% confidence instead of 99% (z = 1.65).
#         Add a variable z_95 and compute var_1d_95. Which assumption did you change? How many places?
# TODO 2: the position grows to $40M. Change ONE line above and re-run both cells.
#         Notice: like Excel, everything downstream updates — but the assumption has a NAME and ONE home.
"""),
    md("""
### Strings and numbers do not mix silently
Excel will happily treat `"5"` and `5` as the same thing until, one day, it doesn't (every analyst has met a ZIP code or CUSIP turned into a number). Python refuses up front:
"""),
    code("""
print("5" + "5")        # text + text = text: '55'
print(5 + 5)            # number + number: 10
try:
    print("5" + 5)      # text + number: an error — on purpose
except TypeError as e:
    print("TypeError:", e)

ticker = "HLT"
shares = 120_000
price = 61.42
print(f"{ticker}: {shares:,} shares × ${price} = ${shares * price:,.0f}")   # f-strings format like a tear sheet
"""),
    code("""
# TODO: compute the simple daily return of a stock that closed at 61.42 yesterday and 62.05 today,
#       and print it as a percentage with 2 decimals (hint: f"{r:.2%}"). Expected: 1.03%
"""),
    md("""
## ⚡ Lab A fast lane (optional)
1. Scale the 1-day VaR to a 10-day horizon using the square-root-of-time rule: `var_10d = var_1d * 10 ** 0.5`.
2. Write the same with `math.sqrt` (`import math`). Which reads better?
3. At what position size does 1-day 99% VaR exceed a $1.5M risk limit? Solve it by algebra, then verify by changing the variable.
"""),
    md("""
## Lab A check-out (2 min)
In the cell below, one sentence: *where do assumptions live in your Excel models today — and who else can find them?*
"""),
    md("""_your answer here_"""),
    md("""
---
# Lab B · Lists, dictionaries, functions, loops (11:20–12:00)

A **list** is an ordered column of values. A **dictionary** is a lookup table (key → value) — the thing you reach for VLOOKUP for. A **function** is a named formula with documentation. A **loop** is fill-down.

Running example: a five-position book.
"""),
    code("""
positions = ["HLT", "MAR", "H", "CHH", "WH"]            # a list: an ordered column
print(positions[0], "|", positions[-1], "|", positions[1:3], "|", len(positions))

book = {                                                  # a dictionary: ticker -> position value ($)
    "HLT": 25_000_000, "MAR": 18_000_000, "H": 9_500_000, "CHH": 4_200_000, "WH": 6_800_000,
}
print("HLT position:", f"${book['HLT']:,}")
book["HLT"] = 27_500_000                                  # update a value
print("Book size:", f"${sum(book.values()):,}")
"""),
    code("""
# TODO 1: add a new position "IHG" of $5_000_000 to the dictionary and reprint the book size.
# TODO 2: positions.append(...) the ticker "IHG" to the list too, and print the new length.
"""),
    md("""
### Functions: define once, reuse, test
"""),
    code("""
def position_pnl(value_start, value_end):
    \"\"\"Dollar P&L and simple return for one position.\"\"\"
    pnl = value_end - value_start
    ret = pnl / value_start
    return pnl, ret

pnl, ret = position_pnl(25_000_000, 25_800_000)
print(f"P&L ${pnl:,.0f} ({ret:.2%})")
"""),
    code("""
# TODO: write a function position_var(value, daily_vol, z=2.33) that returns value * daily_vol * z.
#       Call it on the HLT position with daily_vol=0.018. Expected: about $1,153,575.
"""),
    md("""
### Loops and conditions: fill-down and IF
"""),
    code("""
vols = {"HLT": 0.018, "MAR": 0.016, "H": 0.022, "CHH": 0.019, "WH": 0.025}
POSITION_LIMIT = 20_000_000

for ticker, value in book.items():
    flag = "⚠ LIMIT BREACH" if value > POSITION_LIMIT else ""
    print(f"{ticker:4s} ${value:>12,}  vol {vols.get(ticker, 0):.1%}  {flag}")
"""),
    code("""
# TODO 1: loop over the book and print each position's 1-day 99% VaR, using your position_var function.
# TODO 2: sum the position VaRs into total_var as you loop (start with total_var = 0).
#         (Yes, summing VaRs ignores diversification — deliberately so; we discuss why in the debrief.)
# TODO 3: print only the tickers whose VaR exceeds $500,000 (an if inside the loop).
"""),
    md("""
## ⚡ Lab B fast lane (optional)
1. Sort the book by position VaR, largest first, *without pandas* (hint: `sorted(book, key=..., reverse=True)`).
2. Rewrite the VaR loop as a dictionary comprehension: `var_by_ticker = {t: position_var(v, vols[t]) for t, v in book.items()}`. The goal is a dict of **numbers** you can keep computing with — not printed text. Prove yours works by computing `sum(var_by_ticker.values())` and looking up one ticker, e.g. `var_by_ticker["HLT"]`. (Printing pretty output is fine, but it's a separate second step — if your comprehension produces formatted strings, you've lost the numbers.)
3. Extend the `position_pnl` function (two cells up) with a third parameter `stop_loss=-0.02` (a return threshold with a default value), and make the function also return `True`/`False` for whether the loss breached it.
"""),
    md("""
## Lab B check-out (2 min)
One sentence: *name a calculation you repeat across many rows or tabs at work/school — what would the function be called?*
"""),
    md("""_your answer here_"""),
    md("""
---
# Lab C · pandas: the spreadsheet that fits in a variable (1:30–2:10)

You just watched the live pivot & VLOOKUP demo (`01b_excel_to_pandas_demo.ipynb` — it stays in the repo; re-read it any time). Now the same moves on the real FCC data you'll use in Case 1.

A **DataFrame** is a table. Every Excel operation you know has a one-line equivalent:

| Excel | pandas |
|---|---|
| Open file | `pd.read_csv(...)` |
| Filter rows | `df[df["col"] > 100]` |
| New calculated column | `df["new"] = df["a"] * df["b"]` |
| SUM / AVERAGE / COUNT | `df["col"].sum()`, `.mean()`, `.count()` |
| Pivot table | `df.groupby("key")["col"].agg(["sum","mean"])` |
| VLOOKUP / XLOOKUP | `df.merge(other, on="key", how="left")` |
| Remove duplicates | `df.drop_duplicates()` |
| Sort | `df.sort_values("col", ascending=False)` |
| Save as | `df.to_excel("out.xlsx")` |
"""),
    code("""
import pandas as pd
avail = pd.read_csv(f"{DATA}/case1/bdc_il_fixed_availability_sample.csv.gz", dtype={"frn": str, "block_geoid": str})
print(avail.shape)          # (rows, columns)
avail.head()
"""),
    code("""
avail.info()                # column names, types, missing values — always look at this first
"""),
    code("""
# Filter: which rows are fiber (technology code 50)?
fiber = avail[avail["technology"] == 50]
print(f"{len(fiber):,} fiber offers out of {len(avail):,} rows ({len(fiber)/len(avail):.1%})")
"""),
    code("""
# Pivot: how many offers does each provider have, by technology?
pivot = avail.pivot_table(index="brand_name", columns="technology", values="location_id", aggfunc="count", fill_value=0)
pivot
"""),
    code("""
# Hmm — why do some providers appear twice in the pivot above?
# The table display hides the culprits — print the raw values WITH their quotes to see them:
sorted(avail["brand_name"].unique())
"""),
    code("""
# TODO 1: make the brand names consistent. Hint: avail["brand_name"] = avail["brand_name"].str.strip().str.title()
#         then re-run the unique() above — the duplicates should collapse from 21 variants to 7 brands.
#         But look closely at the survivors: .title() lowercases mid-word capitals, so 'MidState Cable'
#         became 'Midstate Cable' and 'HeritageTel' became 'Heritagetel' — merged, but misbranded.
#         Fix those with .replace({"Midstate Cable": "MidState Cable", ...}) — real-world cleaning
#         is usually a rule (strip/title) plus a small manual exception list. That is normal.

# TODO 2: count how many offers meet the 100/20 Mbps threshold:
#         both max_advertised_download_speed >= 100 AND max_advertised_upload_speed >= 20.
#         Hint: combine two conditions with & and wrap each in parentheses.
"""),
    code("""
# VLOOKUP → merge: attach the technology name to every row
tech = pd.read_csv(f"{DATA}/case1/bdc_technology_codes.csv")
avail = avail.merge(tech, on="technology", how="left")
avail[["brand_name", "technology", "technology_name"]].drop_duplicates().head(10)
"""),
    code("""
# Save your work as an Excel file — Python is not a replacement for Excel as a *delivery* format.
avail.head(1000).to_excel("bootcamp_sample.xlsx", index=False)
print("Saved bootcamp_sample.xlsx — open the file browser on the left to download it.")
"""),
    md("""
## ⚡ Lab C fast lane (optional)
1. Plot the number of offers per technology as a bar chart: `avail["technology_name"].value_counts().plot(kind="bar")`.
2. What share of rows are exact duplicates? (`avail.duplicated().sum()`) Why might a provider file the same location twice?
3. Write a function `served(df, down, up)` that returns the subset meeting any threshold, and call it with (25, 3) and (100, 20).
4. **One-expression rebuild.** Everything you did above happened as separate steps that each
   overwrote `avail`: load the CSV → clean the brand names (strip, title-case, fix the two
   CamelCase brands) → merge on the technology names → filter to offers meeting 100/20 Mbps →
   pivot offers per brand × technology. Now produce that same final pivot as **one chained
   expression** starting from a fresh `pd.read_csv(...)`, with no intermediate variables and
   no re-assigning `avail`. Skeleton to complete:

```python
result = (
    pd.read_csv(f"{DATA}/case1/bdc_il_fixed_availability_sample.csv.gz", dtype={"frn": str, "block_geoid": str})
      .drop_duplicates()
      .assign(brand_name=lambda d: ...)   # the strip/title/replace cleaning, as one expression
      .merge(tech, on="technology", how="left")
      .query("...")                        # the 100/20 threshold
      .pivot_table(...)                    # offers per brand_name × technology_name
)
```

   Why bother? A chain reads top-to-bottom like a recipe, nothing in between can be stale or
   half-updated (a real hazard above — re-run the merge cell twice and you silently get
   duplicate `technology_name_x`/`_y` columns),
   and it's the shape most production pandas code takes.
"""),
    md("""
## Lab C check-out (2 min)
Write one sentence in the cell below: *what is one thing you did here that would have been painful in Excel?*
"""),
    md("""_your answer here_"""),
])

# =====================================================================================
# 01b — EXCEL → PANDAS LIVE DEMO (Day 1, instructor-led, students follow along)
# =====================================================================================
write("01b_excel_to_pandas_demo.ipynb", [
    md("""
# 01b · Excel → pandas, side by side (live demo)

We step through this together in the Day 1 pandas talk — keep it open and run each cell as it appears on screen. It shows the **two Excel moves analysts use most** — the pivot table and the VLOOKUP — built the Excel way (described) and the pandas way (run live), *including the ways each one fails*.

We use a small, readable dataset: the five-position book from Lab B, plus a reference table — small enough to check every number by eye. That is the point: learn the move where you can verify it, then apply it to 236,549 rows in Lab C.
"""),
    code("""
import pandas as pd
import numpy as np

trades = pd.DataFrame({
    "ticker":  ["HLT", "HLT", "MAR", "MAR", "H",  "CHH", "WH", "WH", "WH"],
    "desk":    ["Lodging", "Lodging", "Lodging", "Lodging", "Lodging", "Franchise", "Franchise", "Franchise", "Franchise"],
    "side":    ["BUY", "SELL", "BUY", "BUY", "BUY", "BUY", "BUY", "SELL", "BUY"],
    "shares":  [40_000, 10_000, 25_000, 15_000, 32_000, 28_000, 51_000, 11_000, 20_000],
    "price":   [61.40, 62.10, 155.20, 156.00, 98.75, 118.30, 72.15, 73.00, 72.60],
})
trades["notional"] = trades.shares * trades.price
trades
"""),
    md("""
## 1 · The pivot table

**In Excel:** select the data → Insert → PivotTable → drag `ticker` to Rows, `side` to Columns, `notional` to Values → change "Count" to "Sum" (everyone forgets this once).

**In pandas:** one line, and it is *text* — you can read it in six months, diff it, and re-run it.
"""),
    code("""
pivot = trades.pivot_table(index="ticker", columns="side", values="notional", aggfunc="sum", fill_value=0)
pivot
"""),
    code("""
# groupby is the more general tool — most 'pivots' in practice are groupbys:
summary = trades.groupby("ticker").agg(
    trades_n=("ticker", "count"),
    total_shares=("shares", "sum"),
    gross_notional=("notional", "sum"),
    avg_price=("price", "mean"),
)
summary
"""),
    code("""
# Multi-level pivots — Rows: desk AND ticker. In Excel: drag two fields. Here: a list.
trades.groupby(["desk", "ticker"]).notional.sum()
"""),
    md("""
## 2 · The VLOOKUP

Attach each ticker's sector and daily vol from a reference table.

**In Excel:** `=XLOOKUP(A2, ref!A:A, ref!B:B)` (or the classic `=VLOOKUP(A2, ref!A:C, 2, FALSE)` — exact match flag, column counted by hand).

**In pandas:** `merge` — a left join.
"""),
    code("""
ref = pd.DataFrame({
    "ticker": ["HLT", "MAR", "H", "CHH"],              # note: WH is missing — deliberately
    "sector": ["Lodging", "Lodging", "Lodging", "Franchise"],
    "daily_vol": [0.018, 0.016, 0.022, 0.019],
})
looked_up = summary.reset_index().merge(ref, on="ticker", how="left")
looked_up
"""),
    md("""
### Failure case 1 — the key is missing
Excel shows `#N/A`. pandas shows `NaN`. Same disease, different rash. **Always count your NaNs after a merge:**
"""),
    code("""
missing = looked_up[looked_up.sector.isna()]
print(f"{len(missing)} ticker(s) missing from the reference table: {list(missing.ticker)}")
# Fix: add the row to ref (best), or .fillna() with an explicit default (document why).
"""),
    md("""
### Failure case 2 — the key is duplicated on the right side
The quiet killer. **Excel's VLOOKUP silently returns the *first* match** — no warning, wrong numbers forever. pandas does something louder: it **multiplies the rows**. Louder is better: you can test for it.
"""),
    code("""
ref_dup = pd.concat([ref, pd.DataFrame({"ticker": ["HLT"], "sector": ["Gaming"], "daily_vol": [0.031]})])
oops = summary.reset_index().merge(ref_dup, on="ticker", how="left")
print(f"5 tickers in, {len(oops)} rows out — HLT now appears twice:")
oops[oops.ticker == "HLT"]
"""),
    code("""
# The guard rails, in order of preference:
# 1. validate= makes the assumption explicit — it RAISES if violated:
try:
    summary.reset_index().merge(ref_dup, on="ticker", how="left", validate="one_to_one")
except pd.errors.MergeError as e:
    print("MergeError:", str(e)[:100], "...")
# 2. de-duplicate the reference table first, consciously:
ref_fixed = ref_dup.drop_duplicates(subset="ticker", keep="first")
print("after drop_duplicates:", len(summary.reset_index().merge(ref_fixed, on="ticker", how="left")), "rows — correct")
"""),
    md("""
## 3 · The harder rows of the translation table

These are the "gray rows" from the slide — each worth seeing once, slowly.
"""),
    code("""
# IF / nested IF  →  np.where / np.select
looked_up["risk_bucket"] = np.where(looked_up.daily_vol > 0.02, "high",
                           np.where(looked_up.daily_vol > 0.017, "medium", "low"))
looked_up[["ticker", "daily_vol", "risk_bucket"]]
"""),
    code("""
# TRIM / PROPER / SUBSTITUTE chains  →  .str methods, chained
messy = pd.Series(["  hilton WORLDWIDE ", "marriott intl.", "HYATT hotels  "])
print(messy.str.strip().str.title().str.replace(".", "", regex=False).tolist())
"""),
    code("""
# VLOOKUP on TWO keys (the Excel helper-column trick)  →  merge on a list
px = pd.DataFrame({"ticker": ["HLT", "HLT", "MAR"], "date": ["2026-10-01", "2026-10-02", "2026-10-01"],
                   "close": [61.40, 62.05, 155.20]})
pos = pd.DataFrame({"ticker": ["HLT", "MAR"], "date": ["2026-10-02", "2026-10-01"], "shares": [40_000, 25_000]})
pos.merge(px, on=["ticker", "date"], how="left")        # no helper column, no CONCAT key
"""),
    md("""
## The takeaways (say them back)
1. A pivot is a `groupby`; the spec is *readable text*, not a drag-and-drop you can't audit.
2. A VLOOKUP is a `merge`; count NaNs after every merge; `validate=` turns a silent Excel bug into a loud Python error.
3. The gray rows are the same ideas with more syntax — the cheat sheet (Appendix B of the design doc) has all of them.
"""),
])

# =====================================================================================
# 02 — CASE STUDY 1: FCC WHITESPACE (starter + solution)
# =====================================================================================
CS1_INTRO = md("""
# 02 · Case study 1 — Whitespace analysis for Lakeshore Broadband

**Situation.** Lakeshore Broadband is an imaginary regional cable/fiber operator in northern Illinois (Lake, McHenry, Kane and part of Boone County). A sponsor is considering an investment and wants to know: *where could Lakeshore expand next, and how big is that opportunity?*

**Data.** A classroom extract shaped exactly like the FCC's **Broadband Data Collection (BDC)** fixed-availability download (one row per provider × location × technology), plus a location table, Lakeshore's current footprint, and illustrative county demographics. Everything you need is already loaded for you in Step 1.

> *Side note, not part of today's work:* if you ever want to go deeper after the workshop, the real BDC data — hundreds of millions of rows — is a free download at https://broadbandmap.fcc.gov/data-download, and this notebook's logic runs on it unchanged.

**Deliverable (individual, 45 min):** a ranked table of whitespace counties with units, revenue and payback, exported to Excel, plus one chart.
**Deliverable (group, 25 min):** stress-test the assumptions, choose a top-2 recommendation, and be ready to defend it in 2 minutes.

**It is completely fine to run out of time.** The steps are ordered so that whatever you finish is a coherent piece of analysis, and Part 2 starts from your table's most complete notebook — an unfinished Part 1 costs you nothing. Depth beats speed: better to understand four steps than to paste your way through seven.

**You have done every one of these moves today.** Steps 1–2 are Lab C; step 3 is Lab A (assumptions as named variables); steps 4–5 are the pivot/merge demo; step 6 is Lab B (dictionaries + functions). The case is assembly, not first contact.

**Using AI assistants.** You're welcome to ask an AI assistant (ChatGPT, Claude, Colab's built-in AI) about syntax or what a function does — but **try it yourself first**, and turn off inline code suggestions if you can (in Colab: Settings → AI Assistance), bringing them back only if you're really struggling and a chat on the side isn't enough. If you do use one: read every line it gives you until you can say why it's there, then prove you've retained it — write a *different but similar* line yourself (a new column, another grouping) without the assistant. If you can't, you didn't learn it yet; ask again until you can.

**Definitions this case uses.** These are *choices, not facts* — a different analyst could set different thresholds and get a different answer. In Part 2, your table will stress-test exactly these, so keep them in mind:
- *Served location*: at least one **terrestrial, low-latency, residential** offer at **≥100/20 Mbps**.
- *Competitors in a block*: distinct providers (other than Lakeshore) offering ≥100/20.
- *Whitespace block*: **not** in Lakeshore's footprint, in a county **adjacent** to it, with **≤1** competitor on average.
""")

CS1_CELLS = [  # (starter, solution) — None solution means identical
    (CS1_INTRO, None),
    (SETUP_CELL, None),
    (md("## Step 1 — Load the four tables\nNotice `dtype={...}`: codes like FRN and GEOID must stay text, or leading zeros vanish — the same bug as Excel turning a ZIP code into a number."), None),
    (code("""
import pandas as pd, numpy as np
avail = pd.read_csv(f"{DATA}/case1/bdc_il_fixed_availability_sample.csv.gz", dtype={"frn": str, "block_geoid": str})
locs = pd.read_csv(f"{DATA}/case1/bdc_il_locations_sample.csv.gz", dtype={"block_geoid": str, "county_fips": str})
footprint = pd.read_csv(f"{DATA}/case1/lakeshore_footprint_blocks.csv", dtype={"block_geoid": str, "county_fips": str})
demo = pd.read_csv(f"{DATA}/case1/il_county_demographics_illustrative.csv", dtype={"county_fips": str})
tech = pd.read_csv(f"{DATA}/case1/bdc_technology_codes.csv")

EXCEL_ROW_LIMIT = 1_048_576
print(f"availability rows: {len(avail):,}  |  Excel row limit: {EXCEL_ROW_LIMIT:,}")
print(f"the full-scale version of this file (20x) would be ~{len(avail)*20:,} rows")
avail.head(3)
"""), None),
    (md("## Step 2 — Clean\nThree hygiene steps that every real dataset needs: drop exact duplicates, normalise names, attach readable labels."), None),
    (code("""
n0 = len(avail)
# TODO: drop exact duplicate rows                         -> avail = avail.drop_duplicates()
# TODO: normalise brand names (strip whitespace, title case), then fix the special cases
#       e.g. .replace({"Midstate Cable": "MidState Cable", "Heritagetel": "HeritageTel",
#                      "Skyreach Wireless": "SkyReach Wireless", "Orbitnet": "OrbitNet", "Starmesh": "StarMesh"})
# TODO: merge the technology names onto avail (left join on "technology")
print(f"dropped {n0-len(avail):,} duplicates; brands: {sorted(avail.brand_name.unique())}")
"""), sol("cs1_step2_clean")),
    (md("## Step 3 — Encode the definitions as parameters\nEverything that is a *judgement* goes in a variable at the top. When the partner asks \"what if we use 25/3 instead?\" you change one line and re-run."), None),
    (code("""
DOWN_MIN, UP_MIN = 100, 20          # 'served' threshold (BEAD uses 100/20; older FCC benchmark was 25/3)
SATELLITE = {60, 61}                # excluded from the terrestrial competition count — defend this choice!
OUR_BRAND = "Lakeshore Broadband"

# TODO: build three filtered frames
# res    = residential-capable offers: business_residential_code in ["R", "X"]
# terr   = res, excluding SATELLITE technologies and requiring low_latency == 1
# served = terr meeting BOTH speed thresholds
"""), sol("cs1_step3_filters")),
    (md("""## Step 4 — Location level: who can each home buy from?
One row per location: number of distinct providers meeting the served definition, and whether *our brand* is one of them. This is a pivot table that Excel would need a helper column and COUNTIFS for; here it is one `groupby`.

**Name things for what they *mean*, not for today's parameter values.** Step 3 made the thresholds variables so you can change them in one place — so a column named `n_providers_100_20` would be a bug waiting to happen: change `DOWN_MIN` to 25 and the name is now a lie. Same for baking "lakeshore" into names: use `our_brand_*` and the whole notebook survives a brand change by editing `OUR_BRAND` once.

> **Aside — `lambda` functions.** `lambda s: (s == OUR_BRAND).any()` is a tiny unnamed function, written inline. `.agg`, `.assign` and `.apply` all accept functions as arguments, and when the function is a one-liner you'll only use once, defining it inline with `lambda` beats cluttering the notebook with a named `def`. Rule of thumb: one expression, used once → lambda; anything longer or reused → `def`."""), None),
    (code("""
# TODO: group `served` by location_id and compute
#   n_providers       = number of unique provider_id
#   our_brand_present = whether any brand_name == OUR_BRAND   (hint: a lambda)
# then left-merge onto `locs`, fill missing with 0/False, and add:
#   is_served     = n_providers > 0
#   n_competitors = n_providers - our_brand_present (as int)
loc = None
"""), sol("cs1_step4_location")),
    (md("""**Pause and notice** (30 seconds each — no answers given, ask your table if unsure):
1. After the `groupby().agg()`, did you remember `.reset_index()`? What does `loc` look like if you skip it — and what goes wrong at the merge?
2. You filled missing values with `.fillna(0).astype(int)` in one column and `.fillna(False).astype(bool)` in another. When is each the right choice, and what does a missing value *mean* here in the first place?
3. For `n_competitors` you subtract `our_brand_present`. Why subtract it at all — and why does it need `.astype(int)` first when it's a True/False column?"""), None),
    (md("## Step 5 — Block level and the whitespace flag"), None),
    (code("""
# TODO: aggregate `loc` to block level (group by block_geoid, county_fips, county_name):
#   locations = count of location_id, units = sum of unit_count, pct_served = mean of is_served,
#   avg_competitors = mean of n_competitors, our_brand_share = mean of our_brand_present
# then add in_footprint (block_geoid in footprint.block_geoid) and merge demo columns
#   relation_to_footprint, county_type, median_household_income  on county_fips
MAX_COMPETITORS = 1
blk = None
# TODO: blk["whitespace"] = not in footprint AND avg_competitors <= MAX_COMPETITORS AND relation_to_footprint == "adjacent"
"""), sol("cs1_step5_blocks")),
    (md("""**Pause and notice:** the groupby keys are `["block_geoid", "county_fips", "county_name"]` — but a block belongs to exactly one county, so `block_geoid` alone already identifies each group. Why carry the other two along anyway? (Think about what the result would look like without them, and what you'd have to do to get the county columns back.)"""), None),
    (md("## Step 6 — Size the opportunity\nAssumptions are variables. Change them, re-run, and the whole analysis updates — *and the change is visible in the code*, not hidden in cell F47."), None),
    (code("""
ARPU_MONTHLY = 75
TARGET_PENETRATION = {"urban": .22, "suburban": .30, "exurban": .35, "rural": .40}
CAPEX_PER_PASSING = {"urban": 900, "suburban": 1100, "exurban": 1400, "rural": 2200}

# TODO: add to blk:  target_pen (map county_type), capex = units * capex/passing, annual_revenue = units * target_pen * ARPU * 12
# TODO: county = whitespace blocks grouped by county_name, county_type with
#       ws_blocks, ws_units, annual_revenue, capex, avg_competitors ; payback_years = capex / annual_revenue
#       sorted by annual_revenue descending
county = None
"""), sol("cs1_step6_sizing")),
    (md("## Step 7 — One chart, one Excel file\nThe chart is for you; the Excel file is for the partner who will never open a notebook."), None),
    (code("""
import matplotlib.pyplot as plt
# TODO: horizontal bar chart of annual_revenue by county (in $M), and a second axis or annotation for payback_years
# TODO: export an Excel workbook with three sheets: `county` ("County ranking"), `blk[blk.whitespace]`
#       ("Whitespace blocks"), and an "Assumptions" sheet written from your constants (ARPU, thresholds, ...)
#       with pd.ExcelWriter("lakeshore_whitespace.xlsx") as xw: county.to_excel(xw, "County ranking", index=False) ...
"""), sol("cs1_step7_chart_excel")),
    (md("""
## Individual check-out (5 min)
Answer in a sentence each, in the cell below:
1. Which county has the largest revenue opportunity, and which has the best payback? Are they the same? Why not?
2. What would this analysis have looked like in Excel? Name the specific step that breaks first.

## Part 2 — group extension (25 min, with your table of 3–4)
Hard switch when the clock says so, wherever you are: start from the table's most complete notebook — an unfinished Part 1 is normal.
1. **Sensitivity.** Loop over `ARPU_MONTHLY in (60, 75, 90)` and `MAX_COMPETITORS in (0, 1, 2)` and build a 3×3 table of total payback years. (Hint: wrap steps 5–6 in a function `run(arpu, max_comp)`.)
2. **Definition risk.** Re-run with satellite *included* in the competition count. How many whitespace blocks disappear? Would you show the partner both numbers?
3. **Recommendation.** Pick a top-2 county expansion plan and write three bullets defending it, including one risk you cannot see in this data (hint: the data says nothing about *who owns the poles*).
"""), None),
    (md("""_group answers here_"""), None),
]

write("02_fcc_whitespace_STARTER.ipynb", [s for s, _ in CS1_CELLS])
if SOLUTIONS:
    write("02_fcc_whitespace_SOLUTION.ipynb", [(sc if sc is not None else s) for s, sc in CS1_CELLS],
          out_dir=SOL_NB_DIR, colab_base=SOLUTIONS_COLAB_BASE)

# =====================================================================================
# 03 — CASE STUDY 2: NEWS DASHBOARD (starter + solution)
# =====================================================================================
CS2_INTRO = md("""
# 03 · Case study 2 — A news screen for an investment target (Day 2)

**Situation.** Your team is evaluating **Marriott International** (NASDAQ: MAR) as a potential investment. Before the first call with management you want a quick, repeatable read of what the news says about the company — and, crucially, you want to know *when to trust that read.*

*(Why Marriott and not, say, Hilton? We tried Hilton first — and the feed drowned in headlines about Steve Hilton, who is running for governor of California next month. "Hilton" is also a person, an island, and a celebrity. That query ambiguity is a real lesson — you will meet it again below — but it makes a bad default target three weeks before an election.)*

**What you will build:** `fetch → clean → baseline sentiment → LLM classification → evaluate → iterate the prompt`.

| Part | When | Output |
|---|---|---|
| A. Fetch, clean, baseline | 10:30–11:10 | a DataFrame of headlines with a naive sentiment score, and a first look at where it is wrong |
| B. Gold labels + LLM run | 11:10–12:00 | your team's 20 hand labels, inter-annotator agreement, a first LLM classification |
| B (cont.) Evaluate + iterate | 1:10–1:40 | the accuracy table, the disagreement list, and a **required** `PROMPT_V3` with a logged score |

**Teamwork rule (teams of 5, finance + tech).** Tech students drive the code; finance students drive every *judgement* — what counts as relevant, what an investor calls negative, which disagreements matter. Rotate the keyboard at Part B. The sprint rubric this afternoon scores exactly this split.

**About `newslib`.** The pipeline's plumbing lives in `app/newslib.py` — roughly 300 lines of Python **written for this workshop**. It is *not* an installed package like `pandas` or `numpy`: it is a plain `.py` file sitting in this repository, which is why the setup cell adds the `app/` folder to Python's path before `import newslib` works. You are encouraged to open it and read it. It follows exactly the library-plus-tests pattern the finance cohort pushed through GitHub Actions yesterday with `case1/whitespace.py` — same idea, different case. Today nobody builds plumbing; you build judgement on top of it.

**No API key?** Every cell runs in *mock mode* (a rule-based stand-in). Your team can also run the prompt in a chat UI on 20 headlines by hand — see Part B.

**If the live feed is blocked** on the classroom network, the notebook automatically switches to a *synthetic* feed about a *fictional* hotel company, "Harborview Hotels", so the pipeline can still be exercised. A banner will tell you.
""")

CS2_CELLS = [
    (CS2_INTRO, None),
    (SETUP_CELL, None),
    (md("""## Part A · Step 1 — Fetch
Google News publishes an **RSS feed** for any search — a machine-readable version of the search results that anyone can download, no API key needed. RSS is written in **XML**, a text format (like HTML's stricter cousin) where each headline sits inside tags: `<item><title>…</title><source>…</source><pubDate>…</pubDate></item>`.

`newslib.fetch_google_news` does two things: builds the search URL, then walks the XML and pulls out one row per `<item>`. Remember: `newslib` is our workshop file, not a package — if you want to see the ~25 lines that do this, open `app/newslib.py` in the repo (or run `nl.fetch_google_news??` in a cell to print the source right here)."""), None),
    (code("""
import pandas as pd
import newslib as nl

COMPANY = "Marriott"
DESCRIPTION = "global hotel franchisor, NASDAQ: MAR"
EXTRA = 'when:30d -library'      # Google search syntax: time window + exclusions (the Marriott Library is a thing)

df = nl.fetch_with_fallback(COMPANY, f"{DATA}/case2/news_sample_cache.csv", extra_terms=EXTRA)
if df.get("is_synthetic", pd.Series([False])).any():
    COMPANY, DESCRIPTION = "Harborview Hotels", "fictional global hotel franchisor (synthetic classroom feed)"
print(len(df), "headlines about", COMPANY)
df.head()
"""), None),
    (md("**Look at the raw titles before doing anything else.** What is *not* about the company? What is duplicated? What would fool a keyword approach?"), None),
    (code("""
pd.set_option("display.max_colwidth", 120)
df[["published", "source", "title"]].head(25)
"""), None),
    (code("""
# OPTIONAL (live feed only): see why we abandoned Hilton as the target.
# df_hilton = nl.fetch_google_news("Hilton", extra_terms="when:7d")
# df_hilton.title.head(20)   # how many are about the hotel company vs the candidate for governor?
"""), None),
    (md("""## Step 2 — Clean

Your TODO below calls `nl.clean_headlines(df)`. Before (or after) you run it, **figure out what the function actually does** — pick whichever of these you'd reach for on the job:

- **Ask the built-in AI.** Colab has Gemini built in (the ✨ *Ask Gemini* button, or the chat icon in the left sidebar): paste the function name and ask "what does `clean_headlines` in the open `newslib.py` do, step by step?" Or paste the function into whatever assistant you have open in another window (Claude, ChatGPT).
- **Read the source yourself.** Run `nl.clean_headlines??` in a cell (two question marks prints the full source), or open `app/newslib.py` via Colab's file browser (📁 icon in the left sidebar → `workshop/app/newslib.py`, double-click to view).

Then verify your understanding against the data: the row counts before/after *are* the answer key."""), None),
    (code("""
# TODO: call nl.clean_headlines(df) and compare row counts before/after. What did it remove?
"""), sol("cs2_step2_clean")),
    (md("""## Step 3 — Baseline: a naive lexicon
`newslib.lexicon_sentiment` counts positive and negative words from a small finance word list (real work would use the Loughran–McDonald dictionary). It is deliberately simple. **Your job is to find where it breaks.**

Expect a lot of neutrals on the live feed: real headlines hedge ("weighs", "expected to", "in talks") and the word list is small. That is itself a finding — a baseline that calls almost everything neutral can still score *high accuracy* if your gold labels are mostly neutral too, which is exactly why you will look at per-class precision/recall this afternoon, not just accuracy."""), None),
    (code("""
df = nl.add_lexicon_sentiment(df)
df["lex_label"].value_counts()
"""), None),
    (code("""
# TODO: show the 10 most positive and 10 most negative headlines by lex_score. Do you agree with them?
# TODO: find at least 3 headlines where the lexicon is clearly wrong and write down WHY (negation? sarcasm? finance jargon?
#       a word that is positive in general English but negative for an investor?)
"""), sol("cs2_step3_lexicon_fails")),
    (md("""
## Part B · Step 4 — Hand labels (the gold set) — 15 min, team
Your team labels **15 headlines**: ten that *every* team labels (the common block) plus five that only your team covers. The cell below slices them out of the cleaned frame by your team number.

Each of the five of you labels all 15 **independently, without talking** (relevant / sentiment / category). Then compare: the share where you agree is your **inter-annotator agreement** — no model can be measured more precisely than that. Argue out a final team label for each headline.

**Class pool:** one person per team enters your 15 agreed rows into the shared class sheet — **the link is on the Case 2 slide being projected right now.** The design: 10 teams × the same common block = ten independent labels for each of those ten headlines (that's the class-wide agreement number for the debrief), and 10 teams × 5 unique = 50 more headlines of coverage. By lunch the room owns a ~60-headline gold set with a measured ceiling.

Categories: `earnings, m&a_or_capital, management, demand_or_macro, legal_or_regulatory, product_or_brand, esg_or_labor, other`.
Sentiment is **from an investor's perspective**: a hotel opening is routine → neutral; a CEO departure → usually negative.
"""), None),
    (code("""
TEAM = 1   # <-- your team number (1-10, from the teams slide)

# Deterministic assignment: sort alphabetically so every team slices the same deck.
pool = df.sort_values("title").reset_index(drop=True)
common = pool.iloc[:10]                                  # every team labels these ten
mine = pool.iloc[10 + (TEAM - 1) * 5 : 10 + TEAM * 5]    # five more, unique to your team
to_label = pd.concat([common, mine])
to_label_out = to_label.reindex(columns=["id", "title"]).assign(relevant="", sentiment="", category="")
to_label_out.to_csv("team_labels_blank.csv", index=False)
print("Wrote team_labels_blank.csv —", len(to_label), "headlines. Each of you labels all of them, separately.")
to_label[["title"]]
"""), None),
    (code("""
# After the team has argued out one agreed label per headline, save the result as our_labels.csv
# and upload it to this Colab session (folder icon on the left -> upload). Then re-run this cell.
import os
gold_path = "our_labels.csv"
if os.path.exists(gold_path):
    gold = pd.read_csv(gold_path)
    key = "id" if "id" in df.columns and "id" in gold.columns else "title"
    m = df.merge(gold[[key, "relevant", "sentiment", "category"]], on=key, how="inner", suffixes=("", "_gold"))
    print(len(m), "labelled headlines")
else:
    print("No labels yet — label your 15 headlines, save as our_labels.csv, upload, re-run.")
"""), None),
    (code("""
# If two group members labelled separately, measure agreement:
# a = pd.read_csv("labels_alice.csv"); b = pd.read_csv("labels_bob.csv")
# print("human-human agreement on sentiment:", nl.agreement_rate(a.sentiment, b.sentiment))
"""), None),
    (md("""
## Step 5 — LLM classification
`newslib.PROMPT_V1` is what most people write first. `PROMPT_V2` applies this morning's **context-engineering** lesson — every line in it earned its place in the window: a role and an investor's perspective (so the model knows *whose* judgement to apply), an explicit relevance check (the exit ramp for headlines about the wrong Marriott), a closed category list (no invented labels), a JSON schema (so code can read the answer), and a rationale field (so *you* can read the model's reasoning when it disagrees with you).

**Track A (you brought your own API key — via your university's AI access):** set `ANTHROPIC_API_KEY` in Colab's *Secrets* panel (🔑 icon) — never paste it in a cell — and create the client below. It's your key and your spend — worth batching calls and using a cheaper model while you iterate.
**Track B (no key):** leave `client = None` (mock mode), *and* have one group member paste `PROMPT_V2` + 20 headlines into a chat UI, record the JSON answers into your labels file as `llm_sentiment`, and compare by hand.
"""), None),
    (code("""
client = None
try:
    from google.colab import userdata          # Colab secrets
    # .strip(): a stray newline from the paste would make every API call fail
    os.environ["ANTHROPIC_API_KEY"] = userdata.get("ANTHROPIC_API_KEY").strip()
except Exception:
    pass
if os.environ.get("ANTHROPIC_API_KEY"):
    os.environ["ANTHROPIC_API_KEY"] = os.environ["ANTHROPIC_API_KEY"].strip()
    !pip -q install anthropic
    import anthropic
    client = anthropic.Anthropic()
    print("Real model mode")
else:
    print("Mock mode — deterministic rule-based stand-in (prompt text is ignored in this mode)")
print(nl.PROMPT_V2.format(company=COMPANY, description=DESCRIPTION, headline="<headline>", source="<source>", categories=nl.CATEGORIES))
"""), None),
    (code("""
# TODO: classify the frame with nl.classify_frame(df, COMPANY, DESCRIPTION, client=client) and look at the new llm_* columns
"""), sol("cs2_step5_classify")),
    (md("## Step 6 — Evaluate: lexicon vs LLM vs humans\nAccuracy alone hides the interesting part. Look at the confusion matrix and the per-class report: *which* mistakes does each method make?"), None),
    (code("""
# TODO (if labels exist): merge llm columns onto m and compute
#   nl.accuracy(m.lex_label[rel], m.sentiment[rel]) and nl.accuracy(m.llm_sentiment[rel], m.sentiment[rel]) where rel = m.relevant
#   nl.accuracy(m.llm_relevant.astype(bool), m.relevant)
#   nl.confusion(...), nl.per_class_report(...)
"""), sol("cs2_step6_evaluate")),
    (code("""
# The most useful table in the whole exercise: where the model disagrees with the humans.
if "m" in globals() and "llm_sentiment" in m.columns:
    rel = m.relevant.astype(bool)
    dis = m[rel & (m.llm_sentiment != m.sentiment)][["title", "sentiment", "llm_sentiment", "llm_rationale"]]
    print(len(dis), "disagreements"); display(dis)
else:
    print("Finish the classification and evaluation TODOs above first.")
"""), None),
    (md("""
## Step 7 — Iterate the prompt (REQUIRED — every team ships a V3)
1. **State the change before you run it.** Pick *one* thing V2 gets wrong (read your disagreement list), write one sentence: "V3 adds ___ because V2 ___".
2. Copy `PROMPT_V2` into `PROMPT_V3`, make exactly that one change (e.g. two labelled examples of the hard case, or a rule for 'cost cuts').
3. Re-run and re-score. **Log both numbers.** Keep V3 even if it is *worse* — say why. A logged regression teaches more than an unlogged win.

**Track A:** re-run `classify_frame(..., prompt_template=PROMPT_V3)`.
**Track B (mock mode ignores prompt text):** run V2 and V3 by hand in a chat UI on the same 20 headlines and score both against your labels — the discipline is identical.
"""), None),
    (code("""
# V3_CHANGE = "adds two labelled examples of cost-cut headlines, because V2 calls them negative and our gold says neutral"
# PROMPT_V3 = nl.PROMPT_V2 + "\\n\\nExamples:\\n- 'Company cuts 200 corporate roles in efficiency push' -> sentiment neutral (cost discipline)\\n"
# df_v3 = nl.classify_frame(df, COMPANY, DESCRIPTION, client=client, prompt_template=PROMPT_V3)
# TODO: re-compute accuracy on df_v3, then log:  print(f"V2: {acc_v2:.0%}  V3: {acc_v3:.0%}  change: {V3_CHANGE}")
"""), None),
    (md("""
## ⚡ Fast lanes (tech students: pick one while the team writes the debrief notes)
1. **Batching.** `classify_frame` calls the model once per headline. Batch 10 headlines per call (one JSON array out) and compare wall-clock time and accuracy.
2. **Model swap.** Re-run with a smaller/cheaper model (change the `model=` argument in `newslib.classify_headline`). Build a 2×2: accuracy vs cost.
3. **Cost tracker.** Count input/output tokens per call (the API response reports usage) and print cost-per-100-headlines.
4. **Dashboard.** Run the Streamlit app (cells below) — and deploy it on Streamlit Community Cloud if you finish early.
"""), None),
    (md("""
## Step 8 (fast lane) — Run the dashboard
`app/news_dashboard.py` is a Streamlit app that calls the same `newslib` functions. **Locally:** `streamlit run app/news_dashboard.py`. **In Colab:** run the two cells below — the first starts the server in the background, the second gives you a URL proxied through Colab (no third-party tunnel needed).
"""), None),
    (code("""
APP_DIR = "workshop/app" if os.path.exists("workshop/app") else "../app"
!pip -q install streamlit
# system_raw truly detaches the server. (A plain `!command &` hangs the cell:
# the server keeps the cell's terminal open and Colab waits for it to close.)
# CORS/XSRF off: Colab's proxy serves the page from a different origin, and
# with the protections on, Streamlit rejects the websocket — the page loads
# as empty grey blocks and never renders.
get_ipython().system_raw(
    f"cd {APP_DIR} && streamlit run news_dashboard.py "
    "--server.port 8501 --server.headless true "
    "--server.enableCORS false --server.enableXsrfProtection false "
    "> /tmp/streamlit.log 2>&1 &")
import time; time.sleep(5)
print(open("/tmp/streamlit.log").read()[-500:] if os.path.exists("/tmp/streamlit.log") else "server log not found — run this cell in Colab")
"""), None),
    (code("""
try:
    from google.colab.output import eval_js
    print("Open your dashboard:", eval_js("google.colab.kernel.proxyPort(8501)"))
except Exception:
    print("Not in Colab — open http://localhost:8501")
"""), None),
    (md("""
## A note on version control (done on Day 1)
The finance cohort lived this loop yesterday with Case 1: they pushed `case1/whitespace.py` plus its tests to GitHub, watched Actions go green, broke a function on purpose, watched the red ✗ catch it, and fixed it. `app/newslib.py` follows exactly the same library-plus-tests pattern (its tests live in `app/tests/`), so the same safety net is one push away for today's prompt work: one repo per team, `app/` + `.github/` at the root, and every `PROMPT_V3` committed with its score in the message. Tech students — if a teammate's test idea captures a failure you found today, add it: `assert lexicon_sentiment("record impairment")["lex_label"] != "positive"` is the house style.

## Case 2 debrief prep (before 1:40)
Have ready: your accuracy table (lexicon vs LLM vs gold, with n), your agreement rate, one disagreement you found interesting, and your V2→V3 numbers with the one-sentence change rationale.
"""), None),
]

write("03_news_dashboard_STARTER.ipynb", [s for s, _ in CS2_CELLS])
if SOLUTIONS:
    write("03_news_dashboard_SOLUTION.ipynb", [(sc if sc is not None else s) for s, sc in CS2_CELLS],
          out_dir=SOL_NB_DIR, colab_base=SOLUTIONS_COLAB_BASE)
else:
    print("(private/ not found — skipped the two SOLUTION notebooks)")
