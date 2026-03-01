import sys
import os
import asyncio
import inspect
import logging
from dataclasses import dataclass, field
from typing import Optional, Callable

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "slides"))

from pptx import Presentation
from pydantic import BaseModel, Field
from agents import Agent, ModelSettings, Runner, RunContextWrapper, function_tool
from agents.stream_events import StreamEvent

from intro_slide import create_intro_slide
from bar_chart_slide import create_bar_chart_slide
from bulleted_boxes_slide import create_bulleted_boxes_slide
from numeric_highlight_slide import create_numeric_highlight_slide
from split_bullet_slide import create_split_bullet_slide
from table_slide import create_table_slide

DEFAULT_MODEL = os.environ.get("EASYPRES_MODEL", "gpt-5.2")


@dataclass
class PresentationContext:
    prs: Presentation = field(default_factory=Presentation)
    output_path: str = "output.pptx"


# ---------------------------------------------------------------------------
# Pydantic models for structured tool parameters
# ---------------------------------------------------------------------------

class BulletCard(BaseModel):
    title: str = Field(description="Card heading text")
    bullets: list[str] = Field(description="List of bullet point strings")


class NumericCard(BaseModel):
    label: str = Field(description="Label describing the metric")
    value: str = Field(description="Formatted value string, e.g. '$132.4 B'")


class SplitSection(BaseModel):
    title: str = Field(description="Section heading")
    descriptor: str = Field(description="Section body text")


class BarDataPoint(BaseModel):
    category: str = Field(description="Category or label for this data point")
    value: float = Field(description="Numeric value for this data point")


class BarChartSeries(BaseModel):
    name: str = Field(description="Series name for the legend")
    data_points: list[BarDataPoint] = Field(description="Data points in this series")


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@function_tool
def add_intro_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    subtitle: str,
    date: str,
) -> str:
    """Add a title/intro slide to the presentation.

    Args:
        title: The main title text (max 30 characters; overflow goes to subtitle).
        subtitle: The subtitle text.
        date: The date string to display.
    """
    create_intro_slide(ctx.context.prs, title=title, subtitle=subtitle, date=date)
    return f"Added intro slide: '{title}'"


@function_tool
def add_bar_chart_slide_single(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    data_points: list[BarDataPoint],
    horizontal: bool = True,
    descriptor: Optional[str] = None,
) -> str:
    """Add a bar chart slide with a single data series.

    Args:
        title: Chart title text (max 39 characters; overflow goes to descriptor).
        data_points: List of data points, each with a category name and numeric value.
        horizontal: True for horizontal bars, False for vertical columns.
        descriptor: Optional description text below the title.
    """
    data = {dp.category: dp.value for dp in data_points}
    create_bar_chart_slide(
        ctx.context.prs,
        title=title,
        data=data,
        horizontal=horizontal,
        descriptor=descriptor,
    )
    return f"Added bar chart slide: '{title}'"


@function_tool
def add_bar_chart_slide_multi(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    series: list[BarChartSeries],
    horizontal: bool = True,
    descriptor: Optional[str] = None,
) -> str:
    """Add a bar chart slide with multiple data series.

    Args:
        title: Chart title text (max 39 characters; overflow goes to descriptor).
        series: List of data series, each with a name and data points.
        horizontal: True for horizontal bars, False for vertical columns.
        descriptor: Optional description text below the title.
    """
    data = [
        {"name": s.name, "values": {dp.category: dp.value for dp in s.data_points}}
        for s in series
    ]
    create_bar_chart_slide(
        ctx.context.prs,
        title=title,
        data=data,
        horizontal=horizontal,
        descriptor=descriptor,
    )
    return f"Added multi-series bar chart slide: '{title}'"


@function_tool
def add_bulleted_boxes_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    cards: list[BulletCard],
) -> str:
    """Add a slide with titled cards containing bullet points.

    Args:
        title: Centered title text at the top of the slide.
        cards: List of cards, each with a title and bullet points.
    """
    cards_dicts = [c.model_dump() for c in cards]
    create_bulleted_boxes_slide(ctx.context.prs, title=title, cards=cards_dicts)
    return f"Added bulleted boxes slide: '{title}'"


@function_tool
def add_numeric_highlight_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    subtitle: str,
    cards: list[NumericCard],
    cols: Optional[int] = None,
) -> str:
    """Add a slide with a grid of numeric highlight cards.

    Args:
        title: Main title text (max 39 characters; overflow goes to subtitle).
        subtitle: Description text below the title.
        cards: List of metric cards, each with a label and a value.
        cols: Number of columns in the grid. Auto-picked if omitted.
    """
    cards_dicts = [c.model_dump() for c in cards]
    create_numeric_highlight_slide(
        ctx.context.prs,
        title=title,
        subtitle=subtitle,
        cards=cards_dicts,
        cols=cols,
    )
    return f"Added numeric highlight slide: '{title}'"


@function_tool
def add_split_bullet_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    subtitle: str,
    sections: list[SplitSection],
) -> str:
    """Add a split-layout slide with title/subtitle on the left and sections on the right.

    Args:
        title: Large title text for the left column.
        subtitle: Descriptor text below the title in the left column.
        sections: List of sections, each with a title and descriptor.
    """
    sections_dicts = [s.model_dump() for s in sections]
    create_split_bullet_slide(
        ctx.context.prs,
        title=title,
        subtitle=subtitle,
        sections=sections_dicts,
    )
    return f"Added split bullet slide: '{title}'"


@function_tool
def add_table_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    descriptor: str,
    headers: list[str],
    rows: list[list[str]],
) -> str:
    """Add a table slide to the presentation.

    Args:
        title: Bold title text displayed at the top-left.
        descriptor: Description text displayed at the top-right.
        headers: List of column header strings.
        rows: List of lists, where each inner list is one row of cell values.
    """
    create_table_slide(
        ctx.context.prs,
        title=title,
        descriptor=descriptor,
        headers=headers,
        rows=rows,
    )
    return f"Added table slide: '{title}'"


# ---------------------------------------------------------------------------
# Agent
# ---------------------------------------------------------------------------

AGENT_INSTRUCTIONS = """\
You are a professional presentation builder. Given content, create a polished \
PowerPoint presentation by adding slides one at a time using the available tools.

Guidelines:
- Always start with an intro slide (add_intro_slide).
- Analyze the content and choose the best slide type for each section:
  * Numeric data with labels -> add_numeric_highlight_slide
  * Comparative bar/column data -> add_bar_chart_slide_single (one series) \
or add_bar_chart_slide_multi (multiple series)
  * Categorized bullet points -> add_bulleted_boxes_slide
  * Strategy, overview, or multi-section text -> add_split_bullet_slide
  * Tabular data with rows and columns -> add_table_slide
- Use clear, concise text on each slide. Summarize verbose content.
- Intro slide titles MUST be 30 characters or fewer. Bar chart and numeric highlight titles MUST be 39 characters or fewer. Move any elaborative text to the subtitle or descriptor.
- Once all slides are added, respond with a brief summary of the slides you created.\
"""

presentation_agent = Agent[PresentationContext](
    name="Presentation Builder",
    instructions=AGENT_INSTRUCTIONS,
    model=DEFAULT_MODEL,
    model_settings=ModelSettings(parallel_tool_calls=False),
    tools=[
        add_intro_slide,
        add_bar_chart_slide_single,
        add_bar_chart_slide_multi,
        add_bulleted_boxes_slide,
        add_numeric_highlight_slide,
        add_split_bullet_slide,
        add_table_slide,
    ],
)


# ---------------------------------------------------------------------------
# Event handling
# ---------------------------------------------------------------------------

logger = logging.getLogger(__name__)


def _serialize_event(event: StreamEvent) -> dict:
    """Best-effort serialization of a StreamEvent to a JSON-safe dict."""
    payload: dict = {"type": event.type}
    data = getattr(event, "data", None)
    if data is not None:
        if hasattr(data, "model_dump"):
            payload["data"] = data.model_dump()
        elif hasattr(data, "__dict__"):
            payload["data"] = str(data)
        else:
            payload["data"] = str(data)
    for attr in ("new_agent", "item", "name"):
        val = getattr(event, attr, None)
        if val is not None:
            payload[attr] = str(val)
    return payload


def create_webhook_event_handler(
    webhook_url: Optional[str] = None,
    webhook_headers: Optional[dict[str, str]] = None,
) -> Callable[[StreamEvent], None]:
    """Factory that returns a StreamEvent callback.

    When *webhook_url* is ``None`` the returned callback is a no-op.
    Otherwise every event is POSTed as JSON to the URL (fire-and-forget;
    errors are logged but never raised).
    """
    if webhook_url is None:
        def _noop(event: StreamEvent) -> None:
            pass
        return _noop

    import httpx

    async def _webhook_handler(event: StreamEvent) -> None:
        payload = _serialize_event(event)
        try:
            async with httpx.AsyncClient() as client:
                await client.post(
                    webhook_url,
                    json=payload,
                    headers=webhook_headers or {},
                    timeout=10.0,
                )
        except Exception:
            logger.warning("Webhook POST to %s failed", webhook_url, exc_info=True)

    return _webhook_handler


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

async def build_presentation(
    content: str,
    output_path: str = "output.pptx",
    num_slides: Optional[int] = None,
    on_event: Optional[Callable[[StreamEvent], None]] = None,
    webhook_url: Optional[str] = None,
    webhook_headers: Optional[dict[str, str]] = None,
) -> str:
    """Build a presentation from content using the AI agent.

    Args:
        content: The source material to turn into slides.
        output_path: Where to write the .pptx file.
        num_slides: If provided, instruct the agent to create exactly this
                    many slides (including the intro). Otherwise the agent
                    decides on its own.
        on_event: Callback invoked for every ``StreamEvent`` emitted during
                  the agent run. May be sync or async.  Takes precedence over
                  *webhook_url* when both are given.
        webhook_url: URL to POST each raw event to as JSON.  Ignored when
                     *on_event* is supplied.
        webhook_headers: Extra HTTP headers sent with every webhook POST.

    Returns the path to the saved .pptx file.
    """
    if on_event is not None:
        handler = on_event
    else:
        handler = create_webhook_event_handler(webhook_url, webhook_headers)

    instructions = AGENT_INSTRUCTIONS
    if num_slides is not None:
        instructions += (
            f"\n- You MUST create exactly {num_slides} slides "
            f"(including the intro slide)."
        )

    agent = Agent[PresentationContext](
        name=presentation_agent.name,
        instructions=instructions,
        model=DEFAULT_MODEL,
        model_settings=ModelSettings(parallel_tool_calls=False),
        tools=list(presentation_agent.tools),
    )

    context = PresentationContext(output_path=output_path)
    result = Runner.run_streamed(
        agent,
        input=content,
        context=context,
    )
    async for event in result.stream_events():
        ret = handler(event)
        if inspect.isawaitable(ret):
            await ret

    context.prs.save(context.output_path)
    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Build a PowerPoint presentation from text content.")
    parser.add_argument("content", nargs="?", default=None, help="Text content (or pass via --file)")
    parser.add_argument("--file", "-f", type=str, default=None, help="Path to a text file with the content")
    parser.add_argument("--output", "-o", type=str, default="output.pptx", help="Output .pptx path")
    parser.add_argument("--slides", "-n", type=int, default=None, help="Exact number of slides to create (including intro)")
    args = parser.parse_args()

    if args.file:
        with open(args.file, "r") as fh:
            content = fh.read()
    elif args.content:
        content = args.content
    else:
        parser.error("Provide content as a positional argument or via --file")

    path = asyncio.run(build_presentation(
        content,
        output_path=args.output,
        num_slides=args.slides,
    ))
    print(f"Presentation saved to: {path}")
