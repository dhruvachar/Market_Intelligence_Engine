"""
LLM client configuration.

Why this file changed
----------------------
The original 300s-for-3-agents slowness came from four separate full LLM
round trips (one for industry/market extraction + one per agent), each
generating free-form prose with no cap on length, against a default Ollama
context window much bigger than these prompts need.

Fixes applied here (the agents.py rewrite removes the 4th call entirely):

1. format="json" — constrains Ollama's decoding to valid JSON. This removes
   markdown fences / preamble / "Sure, here's..." filler the model would
   otherwise generate, and means agents.py gets real structured data every
   time instead of fighting brittle regex parsing of prose.
2. num_predict — hard cap on output tokens. Small local models sometimes
   ramble well past what you asked for; this puts a ceiling on worst-case
   latency per call.
3. num_ctx — smaller context window than Ollama's default. Our prompts are
   short, so a large window just costs extra allocation/processing time for
   no benefit.
4. keep_alive — keeps the model resident in memory between the 3 agent
   calls (and across requests), avoiding reload latency.

If generation is still slow on your machine, the remaining lever is the
model itself — see the comment at the bottom of this file.
"""

import os

from langchain_ollama import ChatOllama

OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "qwen2.5:3b")
OLLAMA_BASE_URL = os.environ.get("OLLAMA_BASE_URL")  # None -> langchain default

llm = ChatOllama(
    model=OLLAMA_MODEL,
    base_url=OLLAMA_BASE_URL,
    temperature=0.2,
    format="json",      # force structured, parseable output -> faster + reliable
    num_predict=700,     # hard cap on generated tokens per call
    num_ctx=4096,         # right-sized context window for our prompts
    keep_alive="30m",     # keep the model loaded between agent calls
)


def web_search(query: str, max_results: int = 5):
    # Placeholder hook for grounding agents with live web data. Not wired
    # into the workflow yet — agents currently rely on the model's own
    # knowledge. Swap this for a real search API call if you want the
    # Market Intelligence agent to cite live data.
    return ""


# ---------------------------------------------------------------------------
# Want it faster still?
# ---------------------------------------------------------------------------
# qwen2.5:3b on CPU is the main remaining bottleneck, not this app's code.
# Options, roughly fastest -> most capable:
#   - `ollama pull qwen2.5:1.5b` then OLLAMA_MODEL=qwen2.5:1.5b  (2-3x faster,
#      noticeably less detailed)
#   - `ollama pull llama3.2:3b`  similar speed/quality tier to qwen2.5:3b
#   - Run Ollama with GPU acceleration if one is available
#   - Set num_thread above to match your CPU's physical core count
# ---------------------------------------------------------------------------