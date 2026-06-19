"""
FastAPI backend for AI Consultant Agent Team

3-Agent Flow:

Market Intelligence (also identifies industry/market)
↓
Strategy
↓
Executive Advisory

Changes from the original:
  - Removed the separate `extract_industry_market` call that used to run
    before the graph even started — that's a 4th full LLM round trip per
    request that's now folded into the Market Intelligence agent. This is
    the single biggest latency win available without changing models.
  - RESULTS now stores the full structured report (not just the pptx path),
    so the frontend can show a "recent reports" list and re-open a past
    report without re-running the agents.
  - New endpoints: GET /history, GET /report/{run_id}, GET
    /export/{run_id}/markdown, GET /health.
  - SSE events now include per-agent timing so the frontend can show how
    long each agent actually took.
"""

import json
import time
import uuid
import asyncio

from pathlib import Path
from typing import AsyncGenerator

from fastapi import FastAPI, HTTPException
from fastapi.responses import (
    HTMLResponse,
    StreamingResponse,
    FileResponse,
    PlainTextResponse,
)
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from state import ConsultingState
from graph import build_graph
from report_builder import build_presentation, build_markdown_report

app = FastAPI(title="AI Consultant Agent Team")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# run_id -> {"pptx_path", "data", "created_at", "query"}
RESULTS: dict = {}

AGENT_NAMES = {
    "market_intelligence": "Market Intelligence Agent",
    "strategy": "Strategy Agent",
    "executive_advisory": "Executive Advisory Agent",
}


class AnalyzeRequest(BaseModel):
    query: str


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = Path(__file__).parent / "frontend.html"
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))


@app.get("/health")
async def health():
    return {"status": "ok", "active_reports": len(RESULTS)}


@app.post("/analyze")
async def analyze(req: AnalyzeRequest):

    async def event_stream() -> AsyncGenerator[str, None]:

        def sse(payload: dict):
            return f"data: {json.dumps(payload)}\n\n"

        run_started = time.perf_counter()

        try:
            yield sse({
                "type": "status",
                "agent": "init",
                "message": "Launching agent pipeline...",
                "state": "running",
            })

            initial_state: ConsultingState = {
                "query": req.query,
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
            accumulated = dict(initial_state)

            for step_output in graph.stream(initial_state):
                for node_name, node_update in step_output.items():
                    if node_name == "__end__":
                        continue

                    accumulated.update(node_update)

                    timing = (node_update.get("timings") or {}).get(node_name)
                    message = f"{AGENT_NAMES.get(node_name, node_name)} complete"
                    if timing is not None:
                        message += f" ({timing:.1f}s)"

                    yield sse({
                        "type": "status",
                        "agent": node_name,
                        "message": message,
                        "state": "done",
                        "elapsed": timing,
                    })

                    await asyncio.sleep(0)

            yield sse({
                "type": "status",
                "agent": "pptx",
                "message": "Generating PowerPoint...",
                "state": "running",
            })

            run_id = str(uuid.uuid4())
            reports_dir = Path("reports")
            reports_dir.mkdir(exist_ok=True)
            pptx_path = str(reports_dir / f"report_{run_id}.pptx")

            build_presentation(accumulated, pptx_path)

            total_elapsed = time.perf_counter() - run_started

            report_data = {
                "industry": accumulated.get("industry"),
                "market": accumulated.get("market"),
                "market_analysis": accumulated.get("market_analysis"),
                "strategy": accumulated.get("strategy"),
                "executive_summary": accumulated.get("executive_summary"),
                "financials": accumulated.get("financials"),
                "recommendations": accumulated.get("recommendations"),
                "timings": accumulated.get("timings"),
                "total_seconds": round(total_elapsed, 1),
            }

            RESULTS[run_id] = {
                "pptx_path": pptx_path,
                "data": report_data,
                "created_at": time.time(),
                "query": req.query,
            }

            yield sse({
                "type": "status",
                "agent": "pptx",
                "message": "PowerPoint ready",
                "state": "done",
            })

            yield sse({
                "type": "result",
                "run_id": run_id,
                "data": report_data,
            })

        except Exception as e:
            import traceback
            yield sse({
                "type": "error",
                "message": str(e),
                "detail": traceback.format_exc(),
            })

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/history")
async def history():
    """Recent reports generated this server session, newest first — powers
    the frontend's 'Recent Reports' sidebar."""
    items = [
        {
            "run_id": run_id,
            "industry": entry["data"].get("industry"),
            "market": entry["data"].get("market"),
            "query": entry["query"],
            "created_at": entry["created_at"],
        }
        for run_id, entry in RESULTS.items()
    ]
    items.sort(key=lambda x: x["created_at"], reverse=True)
    return items[:25]


@app.get("/report/{run_id}")
async def report(run_id: str):
    """Full structured data for a past report, so the frontend can re-open
    it from history without re-running the agents."""
    entry = RESULTS.get(run_id)
    if not entry:
        raise HTTPException(404, "Report not found.")
    return entry["data"]


@app.get("/download/{run_id}")
async def download(run_id: str):
    entry = RESULTS.get(run_id)
    if not entry or not Path(entry["pptx_path"]).exists():
        raise HTTPException(404, "Report not found.")

    data = entry["data"]
    filename = f"{data.get('industry', 'report')}_{data.get('market', '')}_strategy_report.pptx".replace(" ", "_")

    return FileResponse(
        entry["pptx_path"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


@app.get("/export/{run_id}/markdown")
async def export_markdown(run_id: str):
    entry = RESULTS.get(run_id)
    if not entry:
        raise HTTPException(404, "Report not found.")
    md = build_markdown_report(entry["data"])
    return PlainTextResponse(md, media_type="text/markdown")