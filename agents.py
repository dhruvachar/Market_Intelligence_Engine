"""
3-Agent Consulting Workflow

1. Market Intelligence Agent  (also identifies industry/market from the
   raw query — this used to be a separate 4th LLM call before every run;
   folding it in here cuts total latency by ~25%)
2. Strategy Agent
3. Executive Advisory Agent

Every agent now asks the model for a single structured JSON object
(clients.py forces Ollama's `format="json"` mode) instead of free-form
prose. That fixes two real bugs in the original version:

  - The executive advisory agent used to return the SAME raw text for
    `executive_summary`, `financials`, AND `recommendations` — i.e. the
    financials and recommendations sections you saw in the UI/PPTX were
    never actually the financials or recommendations, just three copies
    of the whole response. Structured JSON means each field is now the
    real, distinct thing.
  - Free-form text was being dumped verbatim into the UI and into PPTX
    text boxes with no parsing, which is why everything looked "untidy"
    and the PPTX overflowed its slides. Structured fields let both the
    frontend and report_builder.py render real headings/bullets/tables.
"""

import json
import re
import time
from urllib import response

from clients import llm
from state import ConsultingState
import state

# Word/array-length budgets below are intentionally tight — shorter asks
# means fewer tokens generated means lower latency per call, on top of the
# format="json" / num_predict caps already set in clients.py.


def _strip_to_json(text: str) -> str:
    """Best-effort isolation of a JSON object from model output, in case
    the model still wraps it in markdown fences or adds stray text despite
    format="json" mode."""
    if not text:
        return ""
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?", "", cleaned.strip())
    cleaned = re.sub(r"```$", "", cleaned.strip())
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start != -1 and end != -1 and end > start:
        cleaned = cleaned[start : end + 1]
    return cleaned.strip()


def _safe_json(text: str, fallback: dict) -> dict:
    """Parse model output as JSON, falling back to a safe default shape
    (with the raw text preserved under "raw") if parsing fails, so a single
    malformed response never crashes the whole pipeline."""
    candidate = _strip_to_json(text)
    if candidate:
        try:
            data = json.loads(candidate)
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass
    out = dict(fallback)
    out["raw"] = (text or "").strip()[:1500]
    return out


def _log_timing(state: ConsultingState, node: str, seconds: float) -> dict:
    timings = dict(state.get("timings", {}))
    timings[node] = round(seconds, 1)
    return timings


# --------------------------------------------------
# Agent 1: Market Intelligence (+ industry/market extraction)
# --------------------------------------------------

def market_intelligence_agent(state: ConsultingState) -> dict:

    print("MARKET INTELLIGENCE AGENT STARTED")

    state["status_log"].append(
        "📈 Market Intelligence Agent: identifying market & analyzing..."
    )

    prompt = f"""
You are a senior market intelligence consultant.

User request:
{state['query']}

Return ONLY valid JSON.

{{
  "industry":"industry",
  "market":"market",

  "overview":"Detailed market overview (150-200 words)",

  "market_size":"Detailed market size and forecast",

  "growth_trends":[
    "trend 1",
    "trend 2",
    "trend 3"
  ],

  "competitors":[
    {{
      "name":"competitor",
      "note":"detailed competitor description"
    }},
    {{
      "name":"competitor",
      "note":"detailed competitor description"
    }},
    {{
      "name":"competitor",
      "note":"detailed competitor description"
    }}
  ],

  "customer_segments":[
    "segment",
    "segment",
    "segment"
  ],

  "market_drivers":[
    "driver",
    "driver",
    "driver"
  ],

  "barriers_to_entry":[
    "barrier",
    "barrier",
    "barrier"
  ],

  "opportunities":[
    "opportunity",
    "opportunity",
    "opportunity"
  ],

  "risks":[
    "risk",
    "risk",
    "risk"
  ]
}}
"""

    t0 = time.perf_counter()
    response = llm.invoke(prompt)
    elapsed = time.perf_counter() - t0

    fallback = {
        "industry": state.get("industry") or "Target Industry",
        "market": state.get("market") or "Target Market",
        "overview": "",
        "market_size": "",
        "growth_trends": [],
        "competitors": [],
        "opportunities": [],
        "risks": [],
    }
    data = _safe_json(response.content, fallback)

    industry = (data.get("industry") or fallback["industry"]).strip()
    market = (data.get("market") or fallback["market"]).strip()

    print("MARKET INTELLIGENCE AGENT FINISHED")

    state["status_log"].append(
        f"✅ Market Intelligence Agent: done in {elapsed:.1f}s — {industry} / {market}"
    )

    return {
        "industry": industry,
        "market": market,
        "market_analysis": data,
        "status_log": state["status_log"],
        "timings": _log_timing(state, "market_intelligence", elapsed),
    }


# --------------------------------------------------
# Agent 2: Strategy Agent
# --------------------------------------------------

def strategy_agent(state: ConsultingState) -> dict:

    print("STRATEGY AGENT STARTED")

    state["status_log"].append(
        "🎯 Strategy Agent: building strategy..."
    )

    market_analysis_json = json.dumps(state.get("market_analysis", {}))

    prompt = f"""
You are a senior strategy consultant.

Industry:
{state['industry']}

Market:
{state['market']}

Market Analysis:
{market_analysis_json}

Return ONLY valid JSON.

{{
  "swot": {{
      "strengths":[
        "strength",
        "strength",
        "strength"
      ],

      "weaknesses":[
        "weakness",
        "weakness",
        "weakness"
      ],

      "opportunities":[
        "opportunity",
        "opportunity",
        "opportunity"
      ],

      "threats":[
        "threat",
        "threat",
        "threat"
      ]
  }},

  "entry_strategy":[
      "detailed recommendation",
      "detailed recommendation",
      "detailed recommendation"
  ],

  "go_to_market":[
      "step",
      "step",
      "step"
  ],

  "partnerships":[
      "partnership",
      "partnership",
      "partnership"
  ],

  "implementation_roadmap":[
      "phase 1",
      "phase 2",
      "phase 3"
  ]
}}
"""

    t0 = time.perf_counter()
    response = llm.invoke(prompt)
    elapsed = time.perf_counter() - t0

    fallback = {
        "swot": {
            "strengths": [],
            "weaknesses": [],
            "opportunities": [],
            "threats": [],
        },
        "entry_strategy": [],
    }
    data = _safe_json(response.content, fallback)
    # Guard against a model returning swot as something other than a dict
    if not isinstance(data.get("swot"), dict):
        data["swot"] = fallback["swot"]

    print("STRATEGY AGENT FINISHED")

    state["status_log"].append(
        f"✅ Strategy Agent: done in {elapsed:.1f}s"
    )

    return {
        "strategy": data,
        "status_log": state["status_log"],
        "timings": _log_timing(state, "strategy", elapsed),
    }


# --------------------------------------------------
# Agent 3: Executive Advisory Agent
# --------------------------------------------------

def executive_advisory_agent(state: ConsultingState) -> dict:

    print("EXECUTIVE ADVISORY AGENT STARTED")

    state["status_log"].append(
        "🧭 Executive Advisory Agent: preparing board-level report..."
    )

    market_analysis_json = json.dumps(state.get("market_analysis", {}))
    strategy_json = json.dumps(state.get("strategy", {}))

    prompt = f"""
You are a senior McKinsey/Bain strategy partner.

Industry:
{state['industry']}

Market:
{state['market']}

Market Analysis:
{market_analysis_json}

Strategy:
{strategy_json}

Create a realistic executive consulting report.

Use assumptions that fit THIS industry and THIS market.
Do NOT use generic placeholder values.
Do NOT reuse example numbers.
Provide realistic estimates with reasoning.

Return ONLY valid JSON.

{{
  "executive_summary":"Detailed board-level executive summary of at least 300 words",

  "financials": {{
    "investment_reasoning":"Explain what drives investment needs"
    "year1_revenue":"Estimated range"
    "breakeven_timeline":"Estimated timeline"
       "key_cost_drivers":[
          "cost driver",
          "cost driver",
          "cost driver",
          "cost driver"
      ],

      "projection":[
        {{
          "year":"Year 1",
          "revenue":"value",
          "profit":"value"
        }},
        {{
          "year":"Year 2",
          "revenue":"value",
          "profit":"value"
        }},
        {{
          "year":"Year 3",
          "revenue":"value",
          "profit":"value"
        }}
      ]
  }},

  "recommendations":[
      "detailed recommendation",
      "detailed recommendation",
      "detailed recommendation"
  ]
}}
"""

    t0 = time.perf_counter()

    response = llm.invoke(prompt)

    elapsed = time.perf_counter() - t0
    fallback = {
    "executive_summary": "",
    "financials": {
        "investment_reasoning": "",
        "initial_investment": "",
        "year1_revenue": "",
        "year2_revenue": "",
        "year3_revenue": "",
        "breakeven_timeline": "",
        "key_cost_drivers": [],
        "projection": []
    },
    "recommendations": []
}

    data = _safe_json(response.content, fallback)
    return {
        "executive_summary": data.get("executive_summary", ""),
        "financials": data.get("financials", {}),
        "recommendations": data.get("recommendations", []),
        "status_log": state["status_log"],
        "timings": _log_timing(
            state,
            "executive_advisory",
            elapsed
        ),
    }