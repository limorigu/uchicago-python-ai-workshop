# From Excel to Python and AI — UChicago two-day workshop

Materials for a two-day workshop: Day 1 runs 9:00–5:00, Day 2 runs 9:00–3:00.

- **Day 1** — Python bootcamp for 25 finance undergraduates seeing programming
  for the first time: three talks with hands-on labs, data management, Case 1 (a whitespace analysis
  on FCC-shaped broadband data), and Git/CI basics.
- **Day 2** — AI day, joint with 25 UChicago tech majors (50 students total):
  how LLMs work, context engineering, Case 2 (an LLM news screen on an investment target,
  with honest evaluation), closing with a **team extension sprint and 2-minute demos in mixed
  finance + tech teams of five**.

**Start here:** `docs/workshop_design.docx` (agendas, syllabus, pre-reads, exercises, rubrics),
plus the two decks:
`slides/UChicago_Day1_Slides.pptx` (53 slides) and `slides/UChicago_Day2_Slides.pptx` (47 slides).

## For students

Click a notebook's **Open in Colab** link below (or the badge at the top of each notebook).

- **Finance cohort (both days):** before Day 1, follow `docs/day_1_prep_guide.pdf` (two pages, ~45 min).
  It ends with `notebooks/00_setup_check.ipynb`, which should print "You're ready".
- **Tech cohort (Day 2 only):** follow `docs/day_2_prep_guide.pdf` (~45 min) — setup check,
  a skim of Day 1's case, and a short primer on how a PE analyst reads news.

### Day 1 notebooks (finance cohort)

| Notebook | When | What | |
|---|---|---|---|
| `00_setup_check.ipynb` | pre-work | confirm Colab works and the data loads | [Open in Colab](https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/00_setup_check.ipynb) |
| `01_python_bootcamp.ipynb` | 9:55 / 11:20 / 1:30 | Labs A·B·C — one notebook, one section per talk, with fast-lane extras | [Open in Colab](https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/01_python_bootcamp.ipynb) |
| `01b_excel_to_pandas_demo.ipynb` | after lunch | live, step-by-step pivot / lookup / text-cleaning demo (follow along) | [Open in Colab](https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/01b_excel_to_pandas_demo.ipynb) |
| `02_fcc_whitespace_STARTER.ipynb` | 2:45 | Case 1 — whitespace analysis for "Lakeshore Broadband" | [Open in Colab](https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/02_fcc_whitespace_STARTER.ipynb) |

### Day 2 notebooks (everyone)

| Notebook | When | What | |
|---|---|---|---|
| `03_news_dashboard_STARTER.ipynb` | 10:20 | Case 2 — news screen on an investment target (fetch → baseline → LLM → evaluate → dashboard) | [Open in Colab](https://colab.research.google.com/github/limorigu/uchicago-python-ai-workshop/blob/main/notebooks/03_news_dashboard_STARTER.ipynb) |
| any of the above | 1:50 | extension sprint — extend your Case 2 pipeline; menu in the design doc §8 | |

Solution notebooks live in a private companion repo shared with workshop participants only (you'll get an invite via the GitHub username from the pre-workshop survey); open them only when stuck on one step.

## Rebuilding the materials

`build/build_notebooks.py` regenerates the starter notebooks (and, when the private companion repo is checked out at `private/`, the solution notebooks too), `python build/render_prep_guide.py` regenerates both prep-guide PDFs, and `python build/build_two_day_slides.py` rebuilds both decks from the original one-day deck (kept locally, not in this repo).

> **Note:** the checked-in decks and `docs/workshop_design.docx` carry extra visual polish applied by hand after building. Rebuilding reproduces all the content exactly, but not that final formatting — see the warnings at the top of the build scripts before overwriting them.

## Layout

```
data/        case1/ — synthetic, FCC-BDC-shaped Case 1 data (docs §5) · case2/ — synthetic news cache + label template (docs §6)
notebooks/   00 · 01 · 01b demo · 02_fcc_whitespace_STARTER · 03_news_dashboard_STARTER
.github/     workflows/ci.yml — GitHub Actions runs the tests on every push
app/         newslib.py (logic) · news_dashboard.py (Streamlit) · tests/ · requirements.txt
build/       generators: FCC data, notebooks, two-day slide decks, docx, prep-guide PDFs
docs/        workshop_design.docx — the full design document · day_1_prep_guide(.html/.pdf) — finance cohort
             day_2_prep_guide(.html/.pdf) — tech cohort, Day 2 only
slides/      UChicago_Day1_Slides.pptx · UChicago_Day2_Slides.pptx
```

All company names, providers, speeds and headlines in `data/` are invented. County codes and the data
schema mirror the real FCC Broadband Data Collection so the code runs unchanged on the real downloads
(https://broadbandmap.fcc.gov/data-download).
