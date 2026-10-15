"""
newslib.py — shared functions for Case Study 2 (news analysis on an investment target).

Pipeline:  fetch  ->  clean  ->  baseline (lexicon) sentiment  ->  LLM classification  ->  evaluate

Design choices made for a classroom setting:
  * No API key needed for fetching: Google News publishes an RSS feed per search query.
  * The LLM step runs against the Anthropic API when a key is present, and in a
    deterministic "mock" mode otherwise, so every notebook cell runs for every student.
  * An offline cache (data/case2/news_sample_cache.csv, clearly labelled SYNTHETIC) lets the
    exercise proceed if the classroom network blocks the feed.
"""
from __future__ import annotations

import json
import os
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from email.utils import parsedate_to_datetime

import pandas as pd

# --------------------------------------------------------------------------------------
# 1. FETCH
# --------------------------------------------------------------------------------------
GOOGLE_NEWS_RSS = "https://news.google.com/rss/search?q={query}&hl=en-US&gl=US&ceid=US:en"


def build_query(company: str, extra_terms: str = "") -> str:
    """Google News search syntax works here: quotes, OR, -exclusions, when:7d."""
    q = f'"{company}"'
    if extra_terms:
        q += f" {extra_terms}"
    return q


def fetch_google_news(company: str, extra_terms: str = "when:30d", timeout: int = 15) -> pd.DataFrame:
    """Return a DataFrame of headlines for `company` from the Google News RSS feed.

    Columns: title, source, published (UTC datetime), link, query
    Google News caps a feed at ~100 items. Links are Google redirect URLs, not publisher URLs.
    """
    q = build_query(company, extra_terms)
    url = GOOGLE_NEWS_RSS.format(query=urllib.parse.quote(q))
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (classroom-workshop)"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        xml_bytes = resp.read()
    return parse_rss(xml_bytes, query=q)


def parse_rss(xml_bytes: bytes, query: str = "") -> pd.DataFrame:
    """RSS is just XML. Each <item> has title / link / pubDate / source."""
    root = ET.fromstring(xml_bytes)
    rows = []
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        source_el = item.find("source")
        source = (source_el.text or "").strip() if source_el is not None else ""
        # Google News appends " - Source" to the title; strip it when we already have the source
        if source and title.endswith(f" - {source}"):
            title = title[: -len(source) - 3].strip()
        pub = item.findtext("pubDate")
        try:
            published = parsedate_to_datetime(pub) if pub else None
        except (TypeError, ValueError):
            published = None
        rows.append({"title": title, "source": source, "published": published,
                     "link": (item.findtext("link") or "").strip(), "query": query})
    df = pd.DataFrame(rows, columns=["title", "source", "published", "link", "query"])
    if len(df):
        df["published"] = pd.to_datetime(df["published"], utc=True, errors="coerce")
    return df


def load_cached_sample(path: str) -> pd.DataFrame:
    """Offline fallback. The file is SYNTHETIC classroom data — the notebook prints a banner."""
    df = pd.read_csv(path)
    df["published"] = pd.to_datetime(df["published"], utc=True, errors="coerce")
    print("*** USING SYNTHETIC CACHED HEADLINES — these are NOT real news items. ***")
    return df


def fetch_with_fallback(company: str, cache_path: str, extra_terms: str = "when:30d") -> pd.DataFrame:
    try:
        df = fetch_google_news(company, extra_terms)
        if len(df) == 0:
            raise RuntimeError("feed returned 0 items")
        df["is_synthetic"] = False
        return df
    except Exception as e:  # network blocked, feed changed, etc.
        print(f"Live fetch failed ({type(e).__name__}: {e}). Falling back to cache.")
        df = load_cached_sample(cache_path)
        df["is_synthetic"] = True
        return df


# --------------------------------------------------------------------------------------
# 2. CLEAN
# --------------------------------------------------------------------------------------
def clean_headlines(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["title"] = out["title"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    out = out[out["title"].str.len() > 10]
    # near-duplicate removal: same headline syndicated by several outlets
    out["title_key"] = out["title"].str.lower().str.replace(r"[^a-z0-9 ]", "", regex=True)
    out = out.drop_duplicates("title_key").drop(columns="title_key")
    out = out.sort_values("published", ascending=False, na_position="last").reset_index(drop=True)
    return out


# --------------------------------------------------------------------------------------
# 3. BASELINE: lexicon sentiment (a tiny, finance-flavoured word list)
#    Real work would use the Loughran-McDonald dictionary: https://sraf.nd.edu/loughranmcdonald-master-dictionary/
# --------------------------------------------------------------------------------------
POSITIVE = {"beat", "beats", "record", "growth", "grows", "raises", "raised", "upgrade", "upgrades", "upgraded",
            "strong", "expands", "expansion", "outperform", "outperforms", "surge", "surges", "soars", "soar",
            "jumps", "jump", "climbs", "climb", "rallies", "rally", "gain", "gains", "profit", "profitable",
            "momentum", "accelerates", "rebound", "recovery", "resilient", "buyback", "dividend", "milestone",
            "wins", "win", "award", "launches", "opens", "partnership", "exceeds", "tops", "boost", "boosts",
            "optimistic", "bullish", "upbeat", "robust"}
NEGATIVE = {"miss", "misses", "missed", "cuts", "cut", "downgrade", "downgrades", "downgraded", "weak", "weakens",
            "falls", "fall", "drop", "drops", "plunge", "plunges", "plummets", "plummet", "tumbles", "tumble",
            "sinks", "sink", "slides", "slide", "slips", "slip", "slumps", "slump", "lawsuit", "sued", "probe",
            "investigation", "fine", "fined", "layoffs", "decline", "declines", "warning", "warns", "loss",
            "losses", "debt", "default", "breach", "strike", "scandal", "recall", "risk", "risks", "slowdown",
            "concern", "concerns", "fraud", "disappoints", "disappointing", "shortfall", "halts", "halt",
            "delays", "delay", "bearish", "struggles", "struggling"}
NEGATORS = {"not", "no", "never", "without", "fails", "failed"}


def lexicon_sentiment(text: str) -> dict:
    """Score = (#positive - #negative) / #words. Deliberately naive — students will find where it breaks."""
    words = re.findall(r"[a-z']+", text.lower())
    pos = neg = 0
    for i, w in enumerate(words):
        negated = i > 0 and words[i - 1] in NEGATORS
        if w in POSITIVE:
            pos, neg = (pos, neg + 1) if negated else (pos + 1, neg)
        elif w in NEGATIVE:
            pos, neg = (pos + 1, neg) if negated else (pos, neg + 1)
    score = (pos - neg) / max(len(words), 1)
    label = "positive" if score > 0 else "negative" if score < 0 else "neutral"
    return {"lex_score": round(score, 3), "lex_label": label, "lex_pos_hits": pos, "lex_neg_hits": neg}


def add_lexicon_sentiment(df: pd.DataFrame) -> pd.DataFrame:
    scores = pd.DataFrame([lexicon_sentiment(t) for t in df["title"]])
    return pd.concat([df.reset_index(drop=True), scores], axis=1)


# --------------------------------------------------------------------------------------
# 4. LLM CLASSIFICATION
# --------------------------------------------------------------------------------------
CATEGORIES = ["earnings", "m&a_or_capital", "management", "demand_or_macro", "legal_or_regulatory",
              "product_or_brand", "esg_or_labor", "other"]

PROMPT_V1 = """Classify this news headline about {company}. Return JSON with keys sentiment and category.
Headline: {headline}"""

PROMPT_V2 = """You are an analyst at a private equity firm screening news about {company} ({description}).

For the headline below, return ONLY a JSON object with these keys:
- "relevant": true if the headline is about {company} the company (its business, stock, management,
   properties, deals or regulation); false if it is about something else that merely shares the name
   (e.g. a person, a place, a different company) or is generic listicle/advertising content.
- "sentiment": one of "positive", "negative", "neutral" — from the perspective of an INVESTOR in
   {company}. Ask: would this make an owner of the business more or less confident? Routine operational
   news (a hotel opening, a promotion) is "neutral" unless it signals a change in trajectory.
- "category": one of {categories}.
- "rationale": one sentence.

If "relevant" is false, set sentiment to "neutral" and category to "other".

Headline: {headline}
Source: {source}"""

_MOCK_NOISE = ("paris hilton", "hilton head", "perez hilton", "hilton als", "steve hilton", "marriott library",
               "medical center", "high school",
               "middle school", "community college", "neighborhood", "neighbourhood", "yacht club", "stadium",
               "bridge repairs", "apartments", "restaurant", "park through", "capital partners", "capital backs",
               "cheap hotel deals", "best hotels near")
_MOCK_NEG = ("trim", "slip", "soften", "flatten", "step down", "class action", "strike", "impairment", "fined",
             "push back", "warns", "warning", "slower", "pull back", "not immune", "settles", "inquiry", "lawsuit")
_MOCK_POS = ("beats", "raises", "record", "expands", "partners", "completes", "recognized", "surpasses", "rebound",
             "best-positioned", "talks to acquire", "asset-light", "not expected to miss", "signs franchise",
             "urges faster buybacks", "launches", "extends maturities")
_MOCK_NEU = ("reaffirms", "announces quarterly dividend", "names new", "appoints", "prices senior", "opens ",
             "rolls out", "to present", "publishes", "chart of the day", "insiders sell", "debuts", "cuts roughly",
             "slow new hotel construction")


def _mock_classify(headline: str, source: str, company: str) -> dict:
    """Deterministic stand-in for an LLM so the notebook runs without an API key.
    Better than the naive lexicon (handles name collisions, some phrases) but still clearly imperfect —
    students should be able to find its mistakes. A real model call replaces this when a key is present."""
    h = headline.lower()
    if any(n in h for n in _MOCK_NOISE):
        return {"relevant": False, "sentiment": "neutral", "category": "other",
                "rationale": "Name collision — not about the company.", "mode": "mock"}
    lex = lexicon_sentiment(headline)
    if any(p in h for p in _MOCK_NEU):
        lex["lex_label"] = "neutral"
    elif any(p in h for p in _MOCK_NEG):
        lex["lex_label"] = "negative"
    elif any(p in h for p in _MOCK_POS):
        lex["lex_label"] = "positive"
    cat = "other"
    if re.search(r"\b(q[1-4]|quarter|earnings|results|revpar|guidance|outlook|eps)\b", h):
        cat = "earnings"
    elif re.search(r"\b(acqui|merger|deal|buyback|bond|debt|notes|offering|stake|spin)", h):
        cat = "m&a_or_capital"
    elif re.search(r"\b(ceo|cfo|executive|appoint|names|steps down|resign|board)\b", h):
        cat = "management"
    elif re.search(r"\b(travel demand|occupancy|tourism|consumer|recession|rates|macro|bookings)\b", h):
        cat = "demand_or_macro"
    elif re.search(r"\b(lawsuit|sued|probe|regulat|fine|antitrust|court|settle)", h):
        cat = "legal_or_regulatory"
    elif re.search(r"\b(opens|launch|brand|loyalty|honors|property|hotel|resort|app)\b", h):
        cat = "product_or_brand"
    elif re.search(r"\b(union|strike|workers|sustainab|emissions|diversity)\b", h):
        cat = "esg_or_labor"
    sent = lex["lex_label"]
    if cat == "product_or_brand" and sent == "positive":
        sent = "neutral"  # routine openings are not thesis-changing
    return {"relevant": True, "sentiment": sent, "category": cat,
            "rationale": "Mock classifier (keyword rules).", "mode": "mock"}


def llm_classify(headline: str, source: str, company: str, description: str,
                 prompt_template: str = PROMPT_V2, model: str | None = None, client=None) -> dict:
    """Classify one headline. Uses the Anthropic API if `client` is provided, else mock mode.

    To use the real model:
        import anthropic; client = anthropic.Anthropic()   # reads ANTHROPIC_API_KEY from the environment
        llm_classify(h, s, "Marriott", "global hotel franchisor, NASDAQ: MAR", client=client)
    """
    if client is None:
        return _mock_classify(headline, source, company)
    prompt = prompt_template.format(company=company, description=description, headline=headline,
                                    source=source, categories=CATEGORIES)
    model = model or os.environ.get("WORKSHOP_MODEL", "claude-sonnet-4-5")  # update to the current model id
    # SDK >= 1.0 (Aug 2026) removed the temperature/top_p/top_k arguments from
    # messages.create(); current models manage sampling themselves. Exact-run
    # reproducibility was never guaranteed anyway — which is why we evaluate.
    msg = client.messages.create(model=model, max_tokens=300,
                                 messages=[{"role": "user", "content": prompt}])
    text = msg.content[0].text
    try:
        data = json.loads(text[text.find("{"): text.rfind("}") + 1])
    except (ValueError, json.JSONDecodeError):
        data = {"relevant": None, "sentiment": "parse_error", "category": "parse_error", "rationale": text[:200]}
    data["mode"] = model
    return data


def classify_frame(df: pd.DataFrame, company: str, description: str, client=None,
                   prompt_template: str = PROMPT_V2, model: str | None = None) -> pd.DataFrame:
    results = [llm_classify(r.title, r.source, company, description, prompt_template, model, client)
               for r in df.itertuples()]
    res = pd.DataFrame(results).add_prefix("llm_")
    return pd.concat([df.reset_index(drop=True), res], axis=1)


# --------------------------------------------------------------------------------------
# 5. EVALUATE
# --------------------------------------------------------------------------------------
def accuracy(pred: pd.Series, gold: pd.Series) -> float:
    mask = gold.notna()
    return float((pred[mask] == gold[mask]).mean()) if mask.any() else float("nan")


def confusion(pred: pd.Series, gold: pd.Series) -> pd.DataFrame:
    return pd.crosstab(gold.rename("gold"), pred.rename("predicted"), dropna=False)


def per_class_report(pred: pd.Series, gold: pd.Series) -> pd.DataFrame:
    labels = sorted(set(gold.dropna()) | set(pred.dropna()))
    rows = []
    for lab in labels:
        tp = ((pred == lab) & (gold == lab)).sum()
        fp = ((pred == lab) & (gold != lab) & gold.notna()).sum()
        fn = ((pred != lab) & (gold == lab)).sum()
        p = tp / (tp + fp) if tp + fp else 0.0
        r = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * p * r / (p + r) if p + r else 0.0
        rows.append({"label": lab, "support": int((gold == lab).sum()), "precision": round(p, 2),
                     "recall": round(r, 2), "f1": round(f1, 2)})
    return pd.DataFrame(rows)


def agreement_rate(labels_a: pd.Series, labels_b: pd.Series) -> float:
    """Inter-annotator agreement — how often two humans agree. This is the ceiling for any model."""
    mask = labels_a.notna() & labels_b.notna()
    return float((labels_a[mask] == labels_b[mask]).mean()) if mask.any() else float("nan")


@dataclass
class EvalSummary:
    n: int
    accuracy_lexicon: float
    accuracy_llm: float
    human_agreement: float | None = None

    def as_frame(self) -> pd.DataFrame:
        return pd.DataFrame({"metric": ["n labelled", "lexicon accuracy", "LLM accuracy", "human-human agreement"],
                             "value": [self.n, self.accuracy_lexicon, self.accuracy_llm, self.human_agreement]})
