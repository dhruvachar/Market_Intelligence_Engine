"""
Shared state schema for the 3-Agent Consulting Workflow.

NOTE: this file was missing entirely from the original project (every other
module imports `ConsultingState` from here, so the app could not have run
as-uploaded). It has been (re)created with a *structured* schema instead of
free-text blobs — that's what lets the frontend and the PPTX builder render
clean headings/bullets/tables instead of dumping raw model output into a box.
"""

from typing import Dict, List, Optional, TypedDict


class CompetitorInfo(TypedDict, total=False):
    name: str
    note: str


class MarketAnalysis(TypedDict, total=False):
    overview: str
    market_size: str
    growth_trends: List[str]
    competitors: List[CompetitorInfo]
    opportunities: List[str]
    risks: List[str]

    customer_segments: List[str]
    market_drivers: List[str]
    barriers_to_entry: List[str]

    raw: str     # populated only if the model output couldn't be parsed as JSON


class SwotAnalysis(TypedDict, total=False):
    strengths: List[str]
    weaknesses: List[str]
    opportunities: List[str]
    threats: List[str]


class Strategy(TypedDict, total=False):
    swot: SwotAnalysis

    entry_strategy: List[str]

    go_to_market: List[str]

    partnerships: List[str]

    implementation_roadmap: List[str]

    raw: str


class RevenueProjection(TypedDict, total=False):
    year: str
    revenue: str
    profit: str


class Financials(TypedDict, total=False):
    initial_investment: str
    year1_revenue: str
    breakeven_timeline: str
    projection: List[RevenueProjection]
    raw: str


class ConsultingState(TypedDict, total=False):
    query: str
    industry: str
    market: str

    market_analysis: Optional[MarketAnalysis]
    strategy: Optional[Strategy]

    executive_summary: Optional[str]
    financials: Optional[Financials]
    recommendations: Optional[List[str]]

    status_log: List[str]
    timings: Dict[str, float]