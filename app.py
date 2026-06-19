import streamlit as st

from state import ConsultingState
from graph import build_graph
from report_builder import build_presentation

st.set_page_config(
    page_title="AI Consultant Agent Team",
    page_icon="🧭",
    layout="centered"
)

st.title("🧭 AI Consultant Agent Team")
st.caption(
    "Market Intelligence → Strategy → Executive Advisory"
)

st.success("✅ Running locally with Ollama (qwen2.5:3b)")

query = st.text_area(
    "Describe the strategy you want",
    value="Create a market entry strategy for electric vehicles in India.",
    height=80,
)

run = st.button("🚀 Run Analysis", type="primary")


# --------------------------------------------------------------------------
# Rendering helpers — the agents now return structured dicts/lists instead
# of raw text blobs, so we render them as real markdown instead of dumping
# them with st.write(raw_text), which used to produce an untidy wall of text.
# --------------------------------------------------------------------------

def render_market_analysis(ma: dict):
    if not ma:
        st.write("Not available.")
        return
    st.markdown(f"**Overview:** {ma.get('overview', '')}")
    st.markdown(f"**Market size:** {ma.get('market_size', '')}")

    st.markdown("**Growth trends**")
    for t in ma.get("growth_trends", []) or []:
        st.markdown(f"- {t}")

    competitors = ma.get("competitors", []) or []
    if competitors:
        st.markdown("**Top competitors**")
        st.table([{"Company": c.get("name", ""), "Why it matters": c.get("note", "")} for c in competitors])

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Opportunities**")
        for o in ma.get("opportunities", []) or []:
            st.markdown(f"- {o}")
    with col2:
        st.markdown("**Risks**")
        for r in ma.get("risks", []) or []:
            st.markdown(f"- {r}")


def render_strategy(strategy: dict):
    if not strategy:
        st.write("Not available.")
        return
    swot = strategy.get("swot", {}) or {}
    cols = st.columns(2)
    labels = [("strengths", "💪 Strengths"), ("weaknesses", "⚠️ Weaknesses"),
              ("opportunities", "🌱 Opportunities"), ("threats", "🚧 Threats")]
    for i, (key, label) in enumerate(labels):
        with cols[i % 2]:
            st.markdown(f"**{label}**")
            for item in swot.get(key, []) or []:
                st.markdown(f"- {item}")

    st.markdown("**Market entry strategy**")
    for i, step in enumerate(strategy.get("entry_strategy", []) or [], start=1):
        st.markdown(f"{i}. {step}")


def render_financials(fin: dict):
    if not fin:
        st.write("Not available.")
        return
    c1, c2, c3 = st.columns(3)
    c1.metric("Initial Investment", fin.get("initial_investment", "—"))
    c2.metric("Year 1 Revenue", fin.get("year1_revenue", "—"))
    c3.metric("Breakeven", fin.get("breakeven_timeline", "—"))

    projection = fin.get("projection", []) or []
    if projection:
        st.markdown("**Revenue projection**")
        st.table([{"Period": p.get("year", ""), "Revenue": p.get("revenue", "")} for p in projection])


def render_recommendations(recs: list):
    if not recs:
        st.write("Not available.")
        return
    for i, r in enumerate(recs, start=1):
        st.markdown(f"{i}. {r}")


if run:

    if not query.strip():
        st.error("Please enter a query.")
        st.stop()

    status_box = st.status(
        "Launching agent team...",
        expanded=True
    )

    try:
        # NOTE: industry/market used to be extracted via a separate LLM call
        # here before the graph even started. That's now folded into the
        # Market Intelligence agent itself (see agents.py), saving one full
        # model round trip per run — a meaningful chunk of the latency the
        # original version had.
        initial_state: ConsultingState = {
            "query": query,
            "industry": "Target Industry",
            "market": "Target Market",

            "market_analysis": None,
            "strategy": None,

            "executive_summary": None,
            "financials": None,
            "recommendations": None,

            "status_log": [],
            "timings": {},
        }

        graph = build_graph()

        accumulated_state = dict(initial_state)

        with status_box:

            for step_output in graph.stream(initial_state):

                for node_name, node_update in step_output.items():

                    accumulated_state.update(node_update)

                    log = node_update.get(
                        "status_log",
                        []
                    )

                    if log:
                        st.write(log[-1])

        status_box.update(
            label="✅ Analysis complete!",
            state="complete",
            expanded=False
        )

        full_state = accumulated_state

        st.success(
            f"Report generated for **{full_state.get('industry')}** in **{full_state.get('market')}**."
        )

        st.subheader("Executive Summary")
        st.write(
            full_state.get(
                "executive_summary",
                ""
            ) or "Not available."
        )

        with st.expander("📈 Market Intelligence", expanded=True):
            render_market_analysis(full_state.get("market_analysis"))

        with st.expander("🎯 Strategy & SWOT"):
            render_strategy(full_state.get("strategy"))

        with st.expander("💰 Financial Estimate"):
            render_financials(full_state.get("financials"))

        with st.expander("✅ Recommendations"):
            render_recommendations(full_state.get("recommendations"))

        timings = full_state.get("timings") or {}
        if timings:
            st.caption(
                "Agent timing — "
                + " · ".join(f"{k.replace('_', ' ').title()}: {v}s" for k, v in timings.items())
            )

        ppt_path = build_presentation(
            full_state,
            "market_entry_report.pptx"
        )

        with open(ppt_path, "rb") as f:

            st.download_button(
                "⬇️ Download PowerPoint Report",
                data=f.read(),
                file_name=f"{full_state.get('industry')}_{full_state.get('market')}_strategy_report.pptx".replace(" ", "_"),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

    except Exception as e:

        status_box.update(
            label="❌ Failed",
            state="error"
        )

        st.error(
            f"Something went wrong: {e}"
        )

        st.exception(e)



"""uvicorn server:app --reload --port 8000"""