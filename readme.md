# Market Intelligence Engine

An AI-powered multi-agent consulting platform that generates market intelligence reports, business strategies, executive recommendations, financial projections, SWOT analysis, and downloadable PowerPoint presentations.

## Features

### Multi-Agent Architecture

The platform uses specialized AI agents that collaborate to produce consulting-grade reports:

* **Market Intelligence Agent**

  * Market overview
  * Industry trends
  * Market size estimates
  * Competitive landscape

* **Strategy Agent**

  * Market entry strategy
  * Growth opportunities
  * Risk assessment
  * Strategic recommendations

* **Executive Advisory Agent**

  * Executive summary
  * Financial projections
  * Investment estimates
  * Breakeven analysis
  * Board-level recommendations

### Automated Report Generation

* SWOT Analysis
* Financial Estimates
* Strategic Recommendations
* PowerPoint (.pptx) Report Export
* Interactive Web Dashboard

### Modern Frontend

* Futuristic consulting dashboard
* Real-time agent pipeline visualization
* Report history tracking
* Interactive report viewer

## Tech Stack

### Backend

* Python
* FastAPI
* LangGraph
* LangChain
* Groq API
* Llama 3.1 8B Instant

### Frontend

* HTML
* CSS
* JavaScript

### Reporting

* python-pptx

## Project Structure

```text
Market_Intelligence_Engine/
│
├── agents.py              # AI agents
├── clients.py             # LLM configuration
├── graph.py               # LangGraph workflow
├── state.py               # Shared state schema
├── server.py              # FastAPI server
├── app.py                 # Application entry
├── report_builder.py      # PowerPoint generator
├── frontend.html          # User interface
├── requirements.txt
│
├── reports/               # Generated reports
│
└── README.md
```

## Installation

### Clone Repository

```bash
git clone https://github.com/dhruvachar/Market_Intelligence_Engine.git
cd Market_Intelligence_Engine
```

### Create Virtual Environment

```bash
python -m venv venv
```

Activate:

Windows

```bash
venv\Scripts\activate
```

Mac/Linux

```bash
source venv/bin/activate
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

## Environment Variables

Create a `.env` file:

```env
GROQ_API_KEY=your_groq_api_key
```

Get a Groq API key from:

https://console.groq.com

## Running Locally

Start the FastAPI server:

```bash
uvicorn server:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

## Workflow

1. Enter industry and market information.
2. Launch analysis.
3. Market Intelligence Agent performs research.
4. Strategy Agent develops market-entry strategy.
5. Executive Advisory Agent generates business recommendations and financial estimates.
6. PowerPoint report is generated automatically.
7. Results are displayed in the dashboard and available for download.

## Example Use Cases

* Market Entry Analysis
* Startup Feasibility Studies
* Competitive Intelligence
* Investment Research
* Industry Opportunity Assessment
* Strategic Consulting

## Future Improvements

* Live web search integration
* Real-time market data
* Advanced financial modeling
* Interactive charts and dashboards
* 3D AI consultant interface
* Multi-language support
* PDF report export

## Author

**Dhruv Achar**

GitHub:
https://github.com/dhruvachar
