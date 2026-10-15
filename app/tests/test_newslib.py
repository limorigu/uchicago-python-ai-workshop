"""Tests for newslib — these run in CI on every push (see .github/workflows/ci.yml).

A finance analogy: these are the 'check-sum' cells at the bottom of a model tab, except they run
automatically, every time, and block a merge when they fail.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import newslib as nl  # noqa: E402

FEED = b"""<rss><channel>
<item><title>Hilton Reports Results - Reuters</title><link>https://news.google.com/a</link>
<pubDate>Tue, 22 Sep 2026 12:00:00 GMT</pubDate><source url="https://reuters.com">Reuters</source></item>
<item><title>Hilton Reports Results - Reuters</title><link>https://news.google.com/b</link>
<pubDate>Tue, 22 Sep 2026 13:00:00 GMT</pubDate><source url="https://reuters.com">Reuters</source></item>
<item><title>Paris Hilton Launches Podcast - People</title><link>https://news.google.com/c</link>
<pubDate>Mon, 21 Sep 2026 09:00:00 GMT</pubDate><source url="https://people.com">People</source></item>
</channel></rss>"""


def test_parse_rss_strips_source_suffix_and_parses_dates():
    df = nl.parse_rss(FEED)
    assert len(df) == 3
    assert df.loc[0, "title"] == "Hilton Reports Results"
    assert df.loc[0, "source"] == "Reuters"
    assert str(df["published"].dtype).startswith("datetime64")


def test_clean_removes_syndicated_duplicates():
    df = nl.clean_headlines(nl.parse_rss(FEED))
    assert len(df) == 2  # the two identical Reuters headlines collapse to one


def test_lexicon_handles_negation():
    assert nl.lexicon_sentiment("Company beats estimates")["lex_label"] == "positive"
    assert nl.lexicon_sentiment("Company does not beat estimates")["lex_label"] == "negative"
    assert nl.lexicon_sentiment("Company to present at conference")["lex_label"] == "neutral"


def test_mock_classifier_flags_name_collisions():
    out = nl.llm_classify("Paris Hilton Launches Podcast", "People", "Hilton", "hotel company")
    assert out["relevant"] is False and out["category"] == "other"


def test_classify_frame_adds_prefixed_columns():
    df = pd.DataFrame({"title": ["Hilton Beats Estimates"], "source": ["Reuters"]})
    out = nl.classify_frame(df, "Hilton", "hotel company")
    for col in ("llm_relevant", "llm_sentiment", "llm_category", "llm_rationale"):
        assert col in out.columns


def test_metrics():
    gold = pd.Series(["positive", "negative", "neutral", "neutral"])
    pred = pd.Series(["positive", "neutral", "neutral", "neutral"])
    assert nl.accuracy(pred, gold) == 0.75
    rep = nl.per_class_report(pred, gold).set_index("label")
    assert rep.loc["neutral", "recall"] == 1.0
    assert nl.agreement_rate(gold, pred) == 0.75


def test_prompt_template_has_all_fields():
    txt = nl.PROMPT_V2.format(company="X", description="d", headline="h", source="s", categories=nl.CATEGORIES)
    assert "relevant" in txt and "sentiment" in txt and "category" in txt
