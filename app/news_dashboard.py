"""
Streamlit dashboard — news screen for an investment target.

Run locally:      streamlit run news_dashboard.py
Run from Colab:   see notebooks/03_news_dashboard.ipynb, section 6 (port proxy)

Everything the app does is a function in newslib.py; the app is only the presentation layer.
That separation is what makes the logic testable (tests/test_newslib.py) and CI-able.
"""
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import newslib as nl

HERE = Path(__file__).resolve().parent
CACHE = HERE.parent / "data" / "case2" / "news_sample_cache.csv"

st.set_page_config(page_title="Target News Screen", layout="wide")
st.title("Investment target — news screen")
st.caption("Workshop build. Headlines from Google News RSS; sentiment from a naive lexicon and an LLM classifier.")

# ---------------- sidebar: inputs -----------------------------------------------------
with st.sidebar:
    company = st.text_input("Company", value="Marriott")
    description = st.text_input("One-line description (helps the LLM disambiguate)",
                                value="global hotel franchisor, NASDAQ: MAR")
    window = st.selectbox("Look-back window", ["when:7d", "when:30d", "when:90d"], index=1)
    exclude = st.text_input("Exclude terms (Google syntax)", value="-library")
    use_llm = st.toggle("Use LLM classifier (needs ANTHROPIC_API_KEY; otherwise mock mode)", value=True)
    go = st.button("Fetch & analyse", type="primary")


@st.cache_data(show_spinner=False, ttl=900)
def load(company: str, window: str, exclude: str) -> pd.DataFrame:
    df = nl.fetch_with_fallback(company, str(CACHE), extra_terms=f"{window} {exclude}".strip())
    df = nl.clean_headlines(df)
    return nl.add_lexicon_sentiment(df)


def get_client():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return None
    try:
        import anthropic
        return anthropic.Anthropic()
    except Exception:
        return None


if go or "df" in st.session_state:
    if go:
        with st.spinner("Fetching headlines…"):
            df = load(company, window, exclude)
        if use_llm:
            with st.spinner("Classifying…"):
                df = nl.classify_frame(df, company, description, client=get_client())
        st.session_state["df"] = df
    df = st.session_state["df"]

    if df.get("is_synthetic", pd.Series([False])).any():
        st.warning("Showing SYNTHETIC cached headlines about a fictional company (live feed unavailable).")

    # ---------------- KPI row --------------------------------------------------------
    rel = df["llm_relevant"].astype(bool) if "llm_relevant" in df else pd.Series(True, index=df.index)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Headlines fetched", len(df))
    k2.metric("Relevant to target", int(rel.sum()))
    sent_col = "llm_sentiment" if "llm_sentiment" in df else "lex_label"
    k3.metric("Negative (relevant)", int((df.loc[rel, sent_col] == "negative").sum()))
    k4.metric("Positive (relevant)", int((df.loc[rel, sent_col] == "positive").sum()))

    # ---------------- charts ---------------------------------------------------------
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Sentiment mix (relevant headlines)")
        mix = df.loc[rel, sent_col].value_counts().reindex(["positive", "neutral", "negative"]).fillna(0)
        st.bar_chart(mix)
    with c2:
        st.subheader("Topics")
        if "llm_category" in df:
            st.bar_chart(df.loc[rel, "llm_category"].value_counts())
        else:
            st.info("Enable the LLM classifier to see topic categories.")

    st.subheader("Headline volume by day")
    daily = df.loc[rel].set_index("published").resample("D").size() if df["published"].notna().any() else pd.Series(dtype=int)
    st.line_chart(daily)

    # ---------------- table ----------------------------------------------------------
    st.subheader("Headlines")
    show_cols = [c for c in ["published", "source", "title", "lex_label", "llm_relevant", "llm_sentiment",
                             "llm_category", "llm_rationale", "link"] if c in df]
    flt = st.multiselect("Filter sentiment", ["positive", "neutral", "negative"], default=[])
    view = df if not flt else df[df[sent_col].isin(flt)]
    st.dataframe(view[show_cols], use_container_width=True, hide_index=True,
                 column_config={"link": st.column_config.LinkColumn("link")})

    # ---------------- disagreement view: where the two methods differ -------------------
    if "llm_sentiment" in df:
        st.subheader("Where the lexicon and the LLM disagree")
        dis = df[rel & (df["lex_label"] != df["llm_sentiment"])]
        st.caption(f"{len(dis)} of {int(rel.sum())} relevant headlines. Read these first — they are where evaluation matters.")
        st.dataframe(dis[["title", "lex_label", "llm_sentiment", "llm_rationale"]], use_container_width=True, hide_index=True)

    st.download_button("Download results (CSV)", df.to_csv(index=False).encode(), file_name=f"{company}_news_screen.csv")
else:
    st.info("Set the company in the sidebar and click **Fetch & analyse**.")
