# From Excel to Python and AI
## A two-day workshop for finance and tech undergraduates — University of Chicago

**Instructor:** Limor Gultchin, Berkshire Partners · **Format:** two days — Day 1 9:00–5:00 (8 hours inclusive of a 1-hour lunch and two coffee breaks), Day 2 9:00–3:00 (6 hours inclusive of a 45-minute working lunch and one coffee break) — in person, laptops required · **Environment:** Google Colab (no installation), GitHub

| | Day 1 | Day 2 |
|---|---|---|
| **Audience** | 25 finance students, most seeing programming outside Excel for the first time | The same 25 finance students **plus** 25 UChicago tech majors (50 total) |
| **Focus** | Python bootcamp, paced for first-time coders; Case 1 (FCC whitespace); Git/GitHub/CI | AI systems, prompt & context engineering, evaluation; Case 2 (news screen) ending in a team extension sprint with demos |
| **Groups** | Fixed tables of 3–4 all day, seeded by survey (one experienced coder per table); labs are individual at the table, Case 1 Part 2 is done as the table | 10 mixed teams of 5 (2–3 finance + 2–3 tech), fixed for the whole day |
| **Materials** | `slides/UChicago_Day1_Slides.pptx`, notebooks 00–02 | `slides/UChicago_Day2_Slides.pptx`, notebook 03 |

---

## 1. Purpose and learning objectives

Finance students leave school fluent in Excel and are then handed datasets Excel cannot open, and AI tools whose output they cannot check. Day 1 closes the first gap slowly and thoroughly, with realistic examples from the daily life of a PE or hedge-fund analyst. Day 2 closes the second gap — and adds 25 tech students so that every team in the room looks like a real deal team: someone who knows the business question sitting next to someone who knows the tooling.

**By 5:00 on Day 1, every finance student should be able to:**

1. Read and write the core Python constructs the cases use — variables, lists, dictionaries, functions, loops, conditions — and explain each one as the Excel concept it generalises.
2. Read a CSV that is too big for Excel, clean it, aggregate it, join it, chart it, and export the result to Excel — in Python, with the cheat sheet open.
3. Turn analytical judgements (thresholds, definitions, assumptions) into named parameters, and re-run an analysis in seconds when the judgement changes.
4. Put code in a GitHub repository, read a CI workflow, and explain what a green check does and does not guarantee.

**By 3:00 on Day 2, every student (both cohorts) should be able to:**

5. Explain how an LLM produces text — including what modern post-training (instruction tuning, RLHF, reasoning training) adds — and predict which tasks it will do well or badly.
6. Write a constrained, testable prompt (role, perspective, closed label set, structured output, rationale) and treat prompt changes as versioned, measured experiments.
7. Build a gold-label set, measure a model against it (accuracy, confusion matrix, per-class precision/recall), and measure the humans against each other.
8. Scope, build, evaluate, and demo an extension to an AI-for-finance pipeline in a 45-minute sprint, in a team where no one person could have done the whole thing.

**Design principles.** Every concept is taught on the case data or a realistic finance example, not toy examples. Day 1 assumes nothing; Day 2 assumes Day 1 (the tech cohort gets a recap slide and a tech-specific prep guide instead). Every exercise has a *fast lane* so experienced students stay engaged. Every "AI" step is paired with an evaluation step. Python is framed as a complement to Excel — Excel remains the delivery format; Python becomes the engine.

## 2. Agendas at a glance

### Day 1 — Python for finance analysts (25 finance students)

| Time | Block | Mode | Materials |
|---|---|---|---|
| 9:00–9:25 | Welcome, why these two days, setup check | Talk + live | Slides D1 §0, `00_setup_check.ipynb` |
| 9:25–9:55 | **Python I — values, variables, types** (incl. the VaR mini-example) | Talk | Slides D1 §1 |
| 9:55–10:30 | Lab A — variables and types on real numbers | Individual | `01_python_bootcamp.ipynb` Lab A |
| 10:30–10:50 | Coffee break | | |
| 10:50–11:20 | **Python II — lists, dictionaries, functions, loops** (with in-slide exercises) | Talk | Slides D1 §2 |
| 11:20–12:00 | Lab B — functions and loops on a mini portfolio | Individual, at your table | `01_python_bootcamp.ipynb` Lab B |
| 12:00–1:00 | Lunch | | |
| 1:00–1:30 | **Python III — the DataFrame; Excel→pandas** (live pivot & VLOOKUP demo in Colab) | Talk + live | Slides D1 §3, `01b_excel_to_pandas_demo.ipynb` |
| 1:30–2:10 | Lab C — pandas on the FCC data | Individual | `01_python_bootcamp.ipynb` Lab C |
| 2:10–2:25 | Coffee break | | |
| 2:25–2:45 | **Data Management & Financial Data Considerations** | Talk | Slides D1 §4 |
| 2:45–2:55 | Case 1 brief: Lakeshore Broadband whitespace | Talk | Slides D1 §5 |
| 2:55–3:40 | Case 1 — individual build | Individual, at your table | `02_fcc_whitespace_STARTER.ipynb` |
| 3:40–4:05 | Case 1 — group extension (sensitivity, recommendation) | Your table of 3–4 (hard switch at 3:40) | same |
| 4:05–4:20 | Case 1 debrief | Discussion | Rubric §9 |
| 4:20–4:50 | **Version control & CI/CD** — talk (15) + push-and-break exercise (15) | Talk + groups | Slides D1 §6, `case1/` |
| 4:50–5:00 | Wrap-up and Day 2 preview | Talk | Slides D1 §7 |

Content totals ≈6 h 25 m: five talks ≈2 h 15 m, hands-on ≈3 h 40 m, debrief/wrap ≈30 m. If the morning runs long, shorten Lab C by 10 minutes rather than the Case 1 debrief.

### Day 2 — AI for investment work (50 students, mixed teams)

| Time | Block | Mode | Materials |
|---|---|---|---|
| 9:00–9:15 | Welcome back; meet the tech cohort (teams posted on screen at 8:45); Day 1 recap on one slide | Talk + activity | Slides D2 §0 |
| 9:15–9:40 | **AI Systems Overview: how they work and where they break** (modernised) | Talk | Slides D2 §1 |
| 9:40–9:55 | **Context Engineering** | Talk | Slides D2 §2 |
| 9:55–10:10 | Coffee break | | |
| 10:10–10:20 | Case 2 brief: news screen on Marriott | Talk | Slides D2 §3 |
| 10:20–12:00 | Case 2 Part A (fetch, clean, baseline) + Part B start (gold labels, LLM run) | Teams | `03_news_dashboard_STARTER.ipynb` |
| 12:00–12:45 | Working lunch — eat together, pick your extension direction | Teams | §8 extension menu |
| 12:45–1:10 | **Evaluating AI Outputs** | Talk | Slides D2 §4 |
| 1:10–1:40 | Case 2 Part B finish — evaluation, disagreement triage, prompt iteration V2→V3 | Teams | same |
| 1:40–1:50 | Case 2 debrief + **extension sprint brief** — rules, rubric | Discussion + talk | Slides D2 §5, §8 |
| 1:50–2:35 | Extension sprint — build | Teams | any notebook + `app/` |
| 2:35–2:55 | Demos — 2 minutes per team, hard stop | Teams | Rubric §8 |
| 2:55–3:00 | Awards and wrap | Talk | Slides D2 §6 |

Content totals ≈5 h: three talks ≈1 h 05 m, case work ≈2 h 10 m, extension sprint + demos ≈1 h 05 m, briefs/debrief/wrap ≈35 m. If Case 2 runs long, trim Part B's fast lanes and keep the sprint start at 1:50; never cut the evaluation talk or the demos.

## 3. Pre-reads and pre-work

Two prep guides, one per cohort. Both end with the same survey (https://forms.gle/12uzGqjYnx3BtkD88): coding background, laptop OS, GitHub username, setup-check status. Responses seed the Day 1 tables of 3–4 and the Day 2 mixed teams, flag anyone whose setup needs help before the day, and drive the invites (by GitHub username) to the private solutions repo.

**Finance cohort — `docs/day_1_prep_guide.pdf` (~90 min, covers both days).** A step-by-step setup checklist for first-time coders: Colab hello-world, package check, GitHub account and first repo, the Colab→GitHub save, the workshop setup check (`00_setup_check.ipynb` until it prints "You're ready"), an AI-assistant account for debugging, and a troubleshooting/contact box. Pre-reads: *Python for Data Analysis* ch. 2 through §2.3 (https://wesmckinney.com/book/python-basics — skim, the labs repeat everything); explore the FCC National Broadband Map for your home address (https://broadbandmap.fcc.gov); 3Blue1Brown, *But what is a GPT?* (https://www.3blue1brown.com/lessons/gpt) at least to next-token prediction; GitHub *Hello World* tutorial. Recommended: the Reinhart–Rogoff Excel error and the London Whale spreadsheet findings (any reputable summary; discussed Day 1 §4).

**Tech cohort — `docs/day_2_prep_guide.pdf` (~45 min, Day 2 only).** Assumes programming fluency; does not assume finance. Setup: run `00_setup_check.ipynb` in Colab, confirm GitHub access. Pre-reads: skim Day 1's Excel→pandas cheat sheet (Appendix B — this is the vocabulary your finance teammates learned); a one-page "how a PE analyst reads news" primer included in the guide (what "investor sentiment" means, why a false negative costs more than a false positive); the same 3Blue1Brown video if transformers are new; Anthropic's prompt-engineering overview (https://docs.claude.com/en/docs/build-with-claude/prompt-engineering/overview). The guide is explicit about the extension sprint: you will be the engineering muscle on a team whose product sense comes from the finance students — the rubric rewards evaluation honesty and finance insight, not just working code.

## 4. Day 1 session syllabus

Slide references are to `slides/UChicago_Day1_Slides.pptx`.

### D1 §0 Welcome (9:00–9:25)

*Objectives:* frame the two days ("engine vs. delivery format" on Day 1; "AI you can check" on Day 2); confirm every laptop runs a cell. *Content:* the two failure modes of finance analysts (data too big, AI unchecked); the arc of both days; ground rules (ask early; fast lanes optional; sit at your assigned table of 3–4 and ask it first; nobody touches anyone else's keyboard). *Activity:* everyone runs `00_setup_check.ipynb` live; instructor triages failures while others read the Case 1 situation paragraph.

### D1 §1 Python I — values, variables, types (9:25–9:55 talk; 9:55–10:30 Lab A)

*Objectives:* students can create variables, know the four basic types, and read a type error. *Content:* a variable is a named cell that cannot silently change type; numbers vs strings vs booleans ("why did `'5' + '5'` give `'55'`?"); f-strings for formatting output like a tear sheet; `type()`, `print()`, reading the error message bottom line first. **Worked finance example (addresses slide note "VaR model define"):** a parametric VaR calculation built from nothing but variables — `position_value`, `daily_vol`, `z_99 = 2.33`, `var_1d = position_value * daily_vol * z_99` — then "the PM asks for 95% instead of 99%": change one variable, re-run, done. The point is made before pandas ever appears: assumptions live at the top, with names. *In-slide exercises (addresses "Exercise" notes):* two 90-second predict-the-output exercises done on paper before the lab. *Lab A:* variables and types on real numbers — position sizing, simple returns, a VaR re-run with new assumptions; fast lane: holding-period VaR scaling by √t.

### D1 §2 Python II — lists, dictionaries, functions, loops (10:50–11:20 talk; 11:20–12:00 Lab B)

*Objectives:* students can store a column of values, look up by key, wrap a calculation in a function, and repeat it over a portfolio. *Content:* a list is a column; a dictionary is a two-column lookup table (ticker → position); a function is a named formula with documentation — define once, reuse, test; loops and conditions are fill-down and IF; indentation is structure, not decoration. **Running example:** a five-position book — loop over it, compute P&L per name, flag breaches of a position limit. *In-slide exercises:* "what does this loop print?" and "write the one-line dictionary lookup" — 90 seconds each, neighbours compare answers. *Lab B:* build `position_pnl()` and `portfolio_var()` functions over the mini portfolio; loop with a threshold alert; fast lane: sort positions by risk contribution without pandas. *Common sticking points:* indentation errors; `=` vs `==`; off-by-one in ranges; forgetting `return`.

### D1 §3 Python III — the DataFrame and Excel→pandas (1:00–1:30 talk + live demo; 1:30–2:10 Lab C)

*Objectives:* students can load, inspect, filter, aggregate, join, and export a DataFrame, and can map each operation to the Excel move they already know. *Content:* the DataFrame as a spreadsheet in a variable; `df.info()` first, always; the Excel→pandas translation table (Appendix B) — filter, calculated column, SUM/AVERAGE, pivot → `groupby`, VLOOKUP → `merge`, remove duplicates, sort, Save As → `to_excel`. **Live Colab demo (addresses slide 13 note):** `01b_excel_to_pandas_demo.ipynb`, projected and stepped through — the *same* pivot table built in Excel (screenshot) and in `groupby` (live), then the *same* VLOOKUP built with `XLOOKUP` (screenshot) and `merge` (live), including the two classic failure cases: lookup key missing (Excel `#N/A` vs pandas `NaN`) and duplicated keys on the right side (Excel silently takes the first match; `merge` multiplies rows — shown, named, and fixed with `drop_duplicates`). The harder translation rows (the "gray ones": `np.where`, text clean-up chains, multi-key merges) sit at the end of the table and are covered in the demo notebook step by step rather than on the slide. *Lab C:* pandas on the real FCC availability sample — load with dtypes, filter to ≥100/20, groupby county, merge demographics, export; fast lane: the same in method-chaining style. *Common sticking points:* boolean masks without parentheses; Series vs DataFrame; reading IDs as numbers.

### D1 §4 Data Management & Financial Data Considerations (2:25–2:45)

*Objectives:* students treat data hygiene as analysis, not overhead, and can name the specific places Excel breaks. *Content:* where financial data comes from and what each source lies about (filings, vendors, government open data, scraped web); the FCC BDC as real open data — hundreds of millions of rows, biannual vintages, challenge process, retroactive corrections; codes vs numbers (FRN, GEOID, CUSIP, ZIP: keep as text); duplicates and their causes; definitional choices (what is "served"? include satellite?) and why they must be explicit parameters; provenance and reproducibility (a notebook is an audit trail; cell F47 is not); Excel's hard limits (1,048,576 rows, silent type coercion, no diff, no test) and the two famous failures; when Excel is still the right tool. *Live moment:* open the 4.7-million-row file in Excel; read the truncation warning aloud. *Bridge to Case 1:* the four tables, the three definitions, the deliverable.

### D1 §5 Case 1 brief (2:45–2:55) — see §6.

### D1 §6 Version control & CI/CD (4:20–4:50)

*Objectives:* students can explain repo/commit/branch/pull request, read `ci.yml` line by line, and have pushed code that ran on a machine that is not theirs. *Content (15 min):* the `model_v12_FINAL_final2.xlsx` problem; Git as a full history with names and reasons; GitHub as the shared copy plus review; the four verbs (clone, commit, push, pull request); `.gitignore` and why an API key must never be committed; what CI is (a script another computer runs on every change) and what tests are (assertions that stay true); reading `ci.yml` aloud: trigger → checkout → set up Python → install → test; CD as "tests green → merge → deploy"; the honest limits (a green check proves the tests you wrote pass, nothing more). *Exercise (15 min):* the repo ships with the afternoon's case distilled into a small library (`case1/whitespace.py`) and four tests (`case1/tests/`) that pin down bugs students just met — the GEOID dtype, the satellite rule, the whitespace cut. Create repo → upload `case1/` and `.github/` → watch Actions go green → break `served_offers` on purpose (delete the satellite exclusion) → red → fix → green → fast lane: add a test of your own (25/3 must serve at least as many rows as 100/20), then a branch + pull request. The key moment is the red X: a test failing on a machine you do not control.

### D1 §7 Wrap-up and Day 2 preview (4:50–5:00)

What to do Monday: rebuild your last Excel analysis in a notebook; keep the cheat sheet open; use an LLM to write the first draft of code and read every line. Tomorrow: 25 tech students join; you know the FCC data, the case, and the business question — they don't. You are not the junior partner tomorrow.

## 5. Day 2 session syllabus

Slide references are to `slides/UChicago_Day2_Slides.pptx`.

### D2 §0 Welcome and team formation (9:00–9:15)

*Objectives:* mixed teams formed and talking before any content; both cohorts share one picture of Day 1. *Content:* roles and the team contract; the Day 1 recap on a single slide — the Lakeshore case and its answer (Cook biggest, DuPage best payback), the constructs, the engine-vs-delivery rule (three-sheet Excel export with an assumptions sheet), and the Git/CI loop every group ran through `case1/` (tech students saw Case 1 in the prep pack; this slide is the shared vocabulary). *Day 1 material is assumed known from here on — there is no separate bridge section or Git/CI module on Day 2; that time goes to Case 2 Part A.* *Activity:* 10 pre-assigned teams of 5 (2–3 finance + 2–3 tech, seeded from surveys — spread the strongest coders and anyone with internship experience) are posted on screen at 8:45 so students sit with their team from the start; each team does a 2-minute icebreaker: finance students explain the Lakeshore case in 60 seconds, tech students explain their favourite tool in 60 seconds. Teams persist through the extension sprint.

### D2 §1 AI Systems Overview (9:15–9:40)

*Objectives:* students can explain next-token prediction in one sentence, say what each training stage adds, and predict which tasks an LLM will do well or badly. *Content (modernised per slide 35 note):* tokens, embeddings, attention in plain language; the modern training pipeline — pre-training (predict the next token on the internet) → supervised instruction tuning (follow instructions) → preference tuning / RLHF and successors (be helpful, refuse well) → **reasoning training** (reinforcement learning on verifiably correct answers; why models now "think out loud" and when that helps); context window and why the model does not "know" your data; determinism (sampling is model-managed in 2026 — no temperature dial; measure run-to-run disagreement instead); model vs chat product vs API; tool use and agents in one slide (a model that can call a search or run code — and why that changes the failure modes without eliminating them). *Failure taxonomy (memorise these four):* hallucination, sycophancy, inconsistency, brittleness to phrasing and position (name collisions: "Hilton" the company vs the person vs the island). *Strong/weak task table (refreshed per slide 39 comment):* strong — classification, extraction, summarisation, format translation, first-draft code, synthesis across supplied documents; weak — arithmetic over many numbers (unless it can run code), precise recall of specific facts, anything requiring data it was not given, consistency across repeated runs, knowing what it does not know. Cost and latency intuition. *Bridge:* "we will use one as a classifier, and we will measure it."

### D2 §2 Context Engineering (9:40–9:55)

*Objectives:* students can decide what earns a place in the model's context window, guide an agent with instructions, examples and tools rather than magic words, and change one thing / measure / keep what wins. *Content:* the model sees one thing — its context window — and attention over it is finite ("context rot"), so context engineering is choosing the smallest set of high-signal tokens; what goes in the window (role/rules/perspective, this item's data clearly delimited) and what stays out (computation and lookup belong to code, secrets/MNPI never go in); examples beat instructions — a few curated canonical examples outperform rule-stuffing, held out of the eval set; structured output is now schema-enforced at the API level — the design work is the schema itself (closed labels with an exit, a rationale field); guiding an agent with a standing instruction file (CLAUDE.md / AGENTS.md — short beats long, a living document) and tool descriptions as context; long tasks — write it down, compress it, isolate it; and the habit that replaces prompt lore: one change, one measurement ("a context change without a measurement is a vibe"). Deliberately out of scope: prompt anatomy taxonomies, move libraries, politeness/threats folklore, model-specific tricks. *Live demo:* `PROMPT_V1` vs `PROMPT_V2` from `newslib.py` on five headlines including a name collision.

### D2 §3 Case 2 brief (10:10–10:20) — see §7.

### D2 §4 Evaluating AI Outputs (12:45–1:10)

*Objectives:* students can design a small evaluation and read its results honestly. *Content:* why a demo is not evidence; the gold set — who labels, how many, the labelling guide; inter-annotator agreement as the ceiling ("if two analysts agree 80% of the time, a 95% model is measuring something else"); accuracy, confusion matrix, per-class precision/recall, and which matters for which decision (missing a negative story costs more than mislabelling a neutral one); baselines (the lexicon) so you know what the model adds; the disagreement list as the primary reading material; repeat runs and the disagreement rate as the error bar (no temperature dial in 2026 SDKs — determinism is measured, not set); evaluation as a living asset — every production failure becomes a test case; LLM-as-judge and its limits; what to report to a partner (n, metric, baseline, three example failures, the decision the number supports). *Bridge:* Part B finish, then the sprint rubric — 25 of 100 points are for exactly this.

### D2 §5 Case 2 debrief and extension sprint brief (1:40–1:50) — see §8.

### D2 §6 Awards and wrap (2:55–3:00)

Winner announced (rubric §8); what to do Monday (both cohorts); the repo stays public — fork it; feedback form.

## 6. Case Study 1 — Whitespace analysis for Lakeshore Broadband (Day 1)

**Situation.** Lakeshore Broadband is an imaginary regional cable/fiber operator in northern Illinois (Lake, McHenry, Kane and part of Boone County). A sponsor is evaluating an investment and wants to know where Lakeshore could expand next and how large that opportunity is.

**Data** (`/data`). Synthetic, but shaped exactly like the FCC Broadband Data Collection fixed-availability download so the code transfers to the real thing (https://broadbandmap.fcc.gov/data-download).

| File | Rows | Role |
|---|---|---|
| `bdc_il_fixed_availability_sample.csv.gz` | 236,549 | One row per provider × location × technology: `frn, provider_id, brand_name, location_id, technology, max_advertised_download_speed, max_advertised_upload_speed, low_latency, business_residential_code, state_usps, block_geoid, h3_res8_id`. Contains ~1% exact duplicates and ~2% inconsistent brand-name casing on purpose. |
| `bdc_il_locations_sample.csv.gz` | 60,500 | Stand-in for the (licensed) FCC Fabric: location → block, county, lat/lon, `unit_count`, land use. |
| `lakeshore_footprint_blocks.csv` | 336 | Census blocks where Lakeshore operates today. |
| `il_county_demographics_illustrative.csv` | 12 | Households, income, county type, relation to footprint (`core`/`adjacent`/`outer`). Illustrative, not Census figures. |
| `bdc_technology_codes.csv` | 9 | FCC technology codes (10 copper, 40 cable, 50 fiber, 60/61 satellite, 70–72 fixed wireless). |

`build/generate_fcc_data.py --scale 20` produces a 4.7-million-row version. Open it in Excel live during D1 §4 — it truncates at 1,048,576 rows with a warning most students have never read.

**Definitions (parameters in the notebook).** Served location: at least one terrestrial, low-latency, residential-capable offer at ≥100/20 Mbps. Competitors in a block: distinct providers other than Lakeshore at ≥100/20. Whitespace block: not in footprint, in an `adjacent` county, average competitors ≤1. Opportunity: units × target penetration (by county type) × ARPU ($75/month) × 12; capex = units × capex per passing (by county type); payback = capex / annual revenue.

**Individual (45 min).** Complete the seven TODO steps in `02_fcc_whitespace_STARTER.ipynb`: load with correct dtypes → clean (duplicates, names, technology labels) → encode definitions → location-level provider counts → block-level aggregation and whitespace flag → sizing → one chart and a three-sheet Excel export (county ranking, whitespace blocks, assumptions). By Day 1 afternoon, every operation in steps 1–4 has been done once in Labs A–C — the case is assembly, not first contact. Check-out: two sentences answering "largest revenue vs best payback — same county?" and "which step breaks first in Excel?". Students work at the same tables of 3–4 they've sat at all day, so the individual→group transition costs nothing: at 3:40 the table switches to Part 2 together regardless of where anyone is — the table starts from its most complete notebook, and an unfinished Part 1 is explicitly fine (fast finishers have the fast lane; anyone stuck >3 minutes has the SOLUTION-notebook valve).

**Group (25 min, same table — no reshuffle).** (1) Sensitivity: wrap steps 5–6 in a function and produce a 3×3 payback table over ARPU ∈ {60, 75, 90} × max competitors ∈ {0, 1, 2}. (2) Definition risk: include satellite in the competition count; count the whitespace blocks that disappear; decide whether to show the partner both numbers. (3) Recommendation: top-2 counties, three bullets, one risk invisible in the data (pole ownership, permitting, incumbent response).

## 7. Case Study 2 — A news screen for an investment target (Marriott) (Day 2)

**Situation.** The team is evaluating Marriott International (NASDAQ: MAR). Before the first management call, build a repeatable read of the news flow — and establish when that read can be trusted. *(Marriott rather than Hilton: Steve Hilton is on the California gubernatorial ballot three weeks after the workshop, so a Hilton feed is dominated by campaign coverage. Hilton appears instead as the name-collision teaching example — tokenization/attention slides, mock classifier noise, and an optional "try Hilton and see" cell in the notebook.)*

**Pipeline.** `fetch (Google News RSS, no key) → clean (near-duplicates) → baseline lexicon sentiment → LLM classification (relevant / investor sentiment / category / rationale) → evaluation against hand labels → Streamlit dashboard`. All logic lives in `app/newslib.py` (≈250 lines); the notebook and the dashboard only call it. Students meet this library for the first time on Day 2 — but they already know the push/CI/test loop from Day 1's exercise on `case1/`.

**What makes Case 2 demanding.** Prompt iteration is *required*, not optional: every team ships a `PROMPT_V3` with a logged score against V2. Gold labels are pooled across the class (each team labels 15 headlines: a 10-headline common block every team labels — giving ten independent labels per headline for a class-wide agreement number — plus 5 team-unique headlines for coverage; 10 teams yield a ~60-headline class gold set by lunch, entered into a shared sheet whose link is on the projected Case 2 slide) so inter-annotator agreement is computed on real volume. Tech fast lanes: batch the LLM calls, swap the model and compare cost/accuracy, add a cost-per-headline tracker. CI/dashboard engineering belongs to Day 1; teams that finish early deploy the dashboard as a fast lane.

**Two contingencies are built in.** If the classroom network blocks the feed, the notebook switches to `data/case2/news_sample_cache.csv`, a synthetic feed about a *fictional* company ("Harborview Hotels & Resorts") with 45 relevant items and 15 name-collision noise items, plus a held-out answer key (in the private solutions repo); a banner announces the switch. If no API key is available, every LLM cell runs in *mock mode* (a rule-based stand-in that is better than the lexicon but visibly imperfect); teams then run the prompt by hand in a chat UI on 20 headlines ("Track B"). Track A assumes students bring their own API key, obtained in advance through their university's AI access, and store it in Colab's Secrets panel; classifying 100 headlines costs each student only cents.

**Part A — teams, within 10:20–12:00 (≈50 min).** Fetch → inspect raw titles (what is not about the company? what is duplicated?) → clean → lexicon baseline → find three headlines where the lexicon is wrong and name the failure type (negation, jargon, "record impairment", investor vs consumer valence). Finance students drive the "is this about the company / does an investor care" judgements; tech students drive the code.

**Part B — teams (start before lunch, finish 1:10–1:40).** (1) Each member labels the team's 15 headlines (10 common + 5 team-unique, sliced deterministically by team number) silently, then compare: compute inter-annotator agreement; argue out agreed labels and enter them in the shared class sheet (link on the projected Case 2 slide). (2) Classify with `PROMPT_V2`. (3) Accuracy, confusion matrix, per-class report for lexicon vs LLM vs gold; relevance accuracy; category accuracy. (4) Read the disagreement list; classify each disagreement as model error, label error, or genuinely ambiguous. (5) **Required:** write `PROMPT_V3` (one change, stated before running), re-score, log both numbers — keep V3 even if it is worse, and say why. (6) Fast lanes: batching, model swap with cost comparison, cost tracker, dashboard deploy.

## 8. Extension sprint — AI for the deal team (Day 2, 1:40–3:00)

**Format.** The final act of Case 2, in the same 10 mixed teams of 5. A 10-minute debrief-plus-brief (1:40–1:50), a 45-minute build (1:50–2:35), a 2-minute demo per team with a hard stop (2:35–2:55), awards (2:55–3:00). Teams choose their direction over the working lunch, so the build starts with a plan, not a blank page. Each team extends its Case 2 pipeline in **one** direction — the menu below is a starting point, not a fence:

1. **Watchlist.** Generalise from one company to a watchlist of three: per-company sentiment rollup, cross-company ranking by materiality, "what moved today" summary. The name-collision problem now exists three times over.
2. **Cost-aware classifier.** Batch the LLM calls, add a cost-per-headline tracker, swap the model, and report the accuracy-per-dollar trade-off a partner would ask about.
3. **Digest.** Turn the day's relevant headlines into an analyst-ready digest: key facts extracted, tone read, three questions for the next management call — separating *extraction* (checkable) from *interpretation* (labelled as such).
4. **Dashboard.** Deploy the Streamlit dashboard with the team's own V3 prompt and evaluation table wired in, so the demo is a product, not a notebook.

**Rules.** Any workshop dataset or the synthetic feeds; students' own API keys (Track A) or mock mode; any AI assistant for writing code — but the team must be able to explain every line it demos. Every demo must include: a working run end-to-end, **one honest evaluation** (a gold set of ≥10 items, a baseline comparison, or a spot-check table — something that could have come out badly), and one sentence on what they would *not* trust it for.

**Judging rubric (100 points).**

| Criterion | Points | 4-point behaviours scaled to range |
|---|---|---|
| Working demo | 30 | Runs live end-to-end on real input; graceful failure beats fake success |
| Evaluation honesty | 25 | A measurement that could have embarrassed the team, reported anyway, with n and baseline |
| Finance insight | 25 | The tool answers a question an investor actually asks; outputs an analyst could act on; limitations framed in deal terms |
| Teamwork & presentation | 20 | Both cohorts visibly contributed; demo within 2 minutes; one clear takeaway |

## 9. Evaluation of exercise outputs

Each criterion scores 1–4 (1 missing, 2 attempted, 3 correct, 4 correct and explained). Rubrics are shared before the exercises; evaluation is formative, not graded. The sprint rubric (§8) is the only scored one.

**Day 1 labs.** Lab A: VaR re-run done by changing variables, not retyping the formula. Lab B: function has a docstring and a `return`; the loop uses the function rather than repeating the formula. Lab C: dtypes set at load; mask parenthesised; groupby output merged, not VLOOKUP-by-eye.

**Case 1 — individual**

| Criterion | 4 looks like |
|---|---|
| Data hygiene | Duplicates dropped with a count printed; brand names normalised; dtypes set so IDs stay text |
| Definitions as parameters | Threshold, satellite set, max competitors, ARPU and capex all named variables at the top; nothing hard-coded downstream |
| Aggregation logic | Location-level provider count excludes Lakeshore from "competitors"; block-level averages computed over locations, not rows |
| Output | Ranked county table with units, revenue, capex, payback; one legible chart; three-sheet Excel export including the assumptions sheet |
| Reflection | Names the specific Excel breaking point (row limit, COUNTIFS over 200k rows, re-running with new assumptions) rather than "Excel is slow" |

**Case 1 — group.** Sensitivity table produced by a *function*, not copy-paste (4) vs re-run by hand (2). Definition-risk answer states the number of blocks lost and takes a position on disclosure. Recommendation cites a risk not in the data.

**Case 2 — Part A.** Raw titles inspected before analysis; cleaning step justified; three lexicon failures with named failure types.

**Case 2 — Part B**

| Criterion | 4 looks like |
|---|---|
| Gold set | 15 headlines (10 common + 5 team-unique) labelled independently by all five members; agreement rate reported and interpreted as a ceiling; agreed labels entered in the class sheet |
| Measurement | Accuracy for lexicon and LLM, confusion matrix, per-class precision/recall, relevance and category accuracy |
| Reading the errors | Each disagreement classified (model error / label error / ambiguous); at least one label changed as a result |
| Prompt discipline | `PROMPT_V3` with exactly one stated change; both scores logged; kept even if worse, with a reason |
| Honesty | Reports n and baseline alongside accuracy; does not claim the model is "good" without saying compared to what |

**Common mistakes to watch for.** Reading IDs as numbers; boolean masks without parentheses; averaging percentages across blocks of different sizes without noting it; counting Lakeshore as its own competitor; treating the mock classifier's output as evidence about the prompt; computing accuracy on all headlines including irrelevant ones (inflates the score); confusing "the model agrees with me" with "the model is right"; on Day 2, tech students doing everything while finance students watch (the rubric's teamwork line exists for this).

## 10. Discussion questions

**Day 1, after Lab C (5 min).** What did you do here that would have been painful in Excel? What was harder in Python than in Excel, honestly?

**Day 1, Case 1 debrief (4:05).** Cook has the largest revenue and the most competitors; DuPage has the best payback. Which do you recommend and what would change your mind? We excluded satellite from competition — a satellite executive would object. Who is right, and does the answer depend on the customer? The data says nothing about pole ownership, permitting or the incumbent's response. How do you present a number you know is missing its biggest risk? If the partner asks for the analysis at 25/3 instead of 100/20, how long does it take you now? How long in Excel?

**Day 1, version control (4:50).** What did the green check prove? What did it not prove? What would you want in the tests before trusting this in a real diligence?

**Day 2, after the AI overview (9:40).** Which parts of your last internship (or last course project) could be described as "classification, extraction or summarisation"? Which could not? If a model is right 90% of the time on headlines, is it useful? For which decision?

**Day 2, after context engineering (9:55).** Why does "from an investor's perspective" change the answer for "Marriott opens 20th property in Asia"? What is the risk of giving the model examples? (It learns your labelling biases too.)

**Day 2, Case 2 debrief (1:40).** Where did the humans disagree with each other, and what does that say about the metric? Which is worse here: a false positive or a false negative? Would you put this dashboard in front of a partner? What would need to be true first? The lexicon was free and instant; the LLM costs money and time. When is the lexicon enough? Did V3 beat V2 — and if it didn't, what did you learn that a win would have hidden?

**Day 2, after demos (2:55).** Which team's evaluation would you trust most, and why? What is the smallest version of your tool that a real deal team could use next week?

## Appendix A — Repository layout

```
uchicago-python-ai-workshop/
├── README.md
├── data/            case1/ (FCC-shaped Case 1 data, §6) · case2/ (synthetic news cache + label template, §7)
├── notebooks/       00_setup_check · 01_python_bootcamp (Labs A/B/C) · 01b_excel_to_pandas_demo
│                    02_fcc_whitespace_STARTER · 03_news_dashboard_STARTER
│                    (SOLUTION notebooks + answer keys live in a private, invite-only companion repo)
├── .github/workflows/ci.yml   GitHub Actions: run the tests on every push / pull request
├── case1/           whitespace.py · tests/        (Day 1 CI exercise: the case distilled into a tested library)
├── app/             newslib.py · news_dashboard.py · tests/ · requirements.txt · .gitignore
├── build/           generate_fcc_data.py · build_notebooks.py · run_notebook.py
├── docs/            this document · day_1_prep_guide.pdf · day_2_prep_guide.pdf
└── slides/          UChicago_Day1_Slides.pptx · UChicago_Day2_Slides.pptx
```

## Appendix B — Excel → pandas cheat sheet

Rows are ordered easy → hard; the last four (gray in the slides) are demonstrated step by step in `01b_excel_to_pandas_demo.ipynb` rather than taught from the table.

| You want to… | Excel | pandas |
|---|---|---|
| Open a file | File → Open | `df = pd.read_csv("file.csv.gz", dtype={"id": str})` |
| See what's in it | scroll | `df.info()`, `df.head()`, `df.describe()` |
| Filter rows | AutoFilter | `df[(df.a >= 100) & (df.b == "R")]` |
| Add a calculated column | `=A2*B2`, fill down | `df["rev"] = df.units * df.arpu * 12` |
| Sum / average / count | `SUM`, `AVERAGE`, `COUNT` | `df.rev.sum()`, `.mean()`, `.count()` |
| Count if / sum if | `COUNTIFS`, `SUMIFS` | `df[df.tech == 50].rev.sum()` |
| Sort | Data → Sort | `df.sort_values("rev", ascending=False)` |
| Remove duplicates | Data → Remove Duplicates | `df.drop_duplicates()` |
| Chart | Insert → Chart | `df.plot(kind="barh", x="county", y="rev")` |
| Save for the partner | Save As .xlsx | `df.to_excel("out.xlsx", index=False)` |
| Pivot table | Insert → PivotTable | `df.groupby("county").agg(units=("units","sum"), pct=("served","mean"))` |
| Look up a value | `VLOOKUP`, `XLOOKUP` | `df.merge(other, on="key", how="left")` |
| If / then | `IF(A2>100,"hi","lo")` | `np.where(df.a > 100, "hi", "lo")` |
| Text clean-up | `TRIM`, `PROPER` | `df.name.str.strip().str.title()` |
| Change an assumption | find the cell, pray | change the variable at the top, re-run |
