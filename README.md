# EasyPres

A lightweight, open-source AI agent that generates polished PowerPoint presentations on demand.

## Why EasyPres?

Everybody is trying to sell you a different presentation SaaS nowadays. Gamma, Beautiful.ai, Tome -- and now even Anthropic is entering the PowerPoint generation game with their Microsoft Powerpoint plugin. The pitch is always the same: hand over your content, pay a subscription, and hope the output looks decent.

But there are very few simple, standalone, and functional open-source alternatives. As individuals, companies, and enterprises evolve, they deserve to **own the infrastructure** powering their AI processes. They should be able to control it -- choose which models generate their content instead of being perpetually dependent on SaaS providers where power is consolidating.

EasyPres is just that: a self-hosted presentation agent that generates real-time, professional `.pptx` files for any content that fits within your chosen model's context window. Bring your own API key, pick your model, and keep full control.

## Features

- **6 slide types** -- intro, bar chart, bulleted boxes, numeric highlights, split bullet, and table slides, each with consistent theming and auto-scaling
- **AI-driven layout** -- the agent analyzes your content and picks the best slide type for each section automatically
- **Configurable slide count** -- request an exact number of slides or let the agent decide
- **REST API** -- a single `POST /generate` endpoint that returns a downloadable `.pptx` binary
- **Webhook events** -- stream raw agent loop events (tool calls, reasoning) to any URL in real time
- **Model-agnostic** -- built on the OpenAI Agents SDK, defaults to GPT-4o, but you can swap in any supported model
- **Fully open source** -- no vendor lock-in, no subscriptions, run it anywhere

## Project Structure

```
src/
  agents/
    presentation_agent.py   # Agent definition, tools, and runner
  api/
    server.py               # FastAPI server with POST /generate
  slides/
    intro_slide.py           # Title/intro slide
    bar_chart_slide.py       # Single and multi-series bar charts
    bulleted_boxes_slide.py  # Titled cards with bullet points
    numeric_highlight_slide.py # Metric highlight grid
    split_bullet_slide.py    # Split layout with sections
    table_slide.py           # Tabular data
requirements.txt
```

## Quickstart

### Prerequisites

- Python 3.11+
- An OpenAI API key

### Setup

```bash
git clone <repo-url> && cd EasyPres
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export OPENAI_API_KEY=sk-...
```

### Run the API server

```bash
uvicorn src.api.server:app --host 0.0.0.0 --port 8000
```

Then generate a presentation:

```bash
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "content": "Your presentation content goes here...",
    "num_slides": 5
  }' \
  -o presentation.pptx
```

The response is a valid `.pptx` file that opens in PowerPoint, Keynote, and Google Slides.

### Run from the command line

```bash
python -m src.agents.presentation_agent "Your content here" -o output.pptx -n 5
```

Or from a file:

```bash
python -m src.agents.presentation_agent -f content.txt -o output.pptx
```

## API Reference

### `POST /generate`

**Request body** (JSON):

| Field | Type | Required | Description |
|---|---|---|---|
| `content` | string | yes | The source material to turn into slides |
| `num_slides` | integer | no | Exact number of slides (including intro). Omit to let the agent decide. |
| `webhook_url` | string | no | URL to POST raw agent events to as JSON |
| `webhook_headers` | object | no | Extra HTTP headers sent with every webhook POST |

**Response**: binary `.pptx` file with content type `application/vnd.openxmlformats-officedocument.presentationml.presentation`.

### Webhook Events

Pass a `webhook_url` in your request and every agent loop event (tool calls, reasoning steps, completions) will be POSTed to that URL as a JSON body with at minimum a `type` field and event-specific data. Headers can be included for authentication via `webhook_headers`.

If no `webhook_url` is provided, events are silently discarded.

## Configuration

| Environment Variable | Description |
|---|---|
| `OPENAI_API_KEY` | Your OpenAI API key (required) |

To change the model, set the `model` parameter on the `Agent` definition in `src/agents/presentation_agent.py`.

## License

Open source. See LICENSE for details.
