"""
LangGraph orchestration for the 3-Agent Consulting Team

Flow:

Market Intelligence Agent  (also identifies industry/market)
        ↓
Strategy Agent
        ↓
Executive Advisory Agent

`extract_industry_market` used to run as a *4th*, separate LLM call before
this graph even started. It's been folded into the Market Intelligence
agent (agents.py) instead — same information, one fewer full model round
trip per run.
"""

from langgraph.graph import StateGraph, END

from state import ConsultingState

from agents import (
    market_intelligence_agent,
    strategy_agent,
    executive_advisory_agent,
)


def build_graph():

    graph = StateGraph(ConsultingState)

    graph.add_node(
        "market_intelligence",
        market_intelligence_agent
    )

    graph.add_node(
        "strategy",
        strategy_agent
    )

    graph.add_node(
        "executive_advisory",
        executive_advisory_agent
    )

    graph.set_entry_point(
        "market_intelligence"
    )

    graph.add_edge(
        "market_intelligence",
        "strategy"
    )

    graph.add_edge(
        "strategy",
        "executive_advisory"
    )

    graph.add_edge(
        "executive_advisory",
        END
    )

    return graph.compile()