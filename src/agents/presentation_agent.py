import sys
import os
import asyncio
import inspect
import logging
import re
from dataclasses import dataclass, field
from typing import Any, Optional, Callable, Union
from urllib.parse import urlparse, urlunparse, unquote

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "slides"))

from pptx import Presentation
from pydantic import BaseModel, Field
from agents import Agent, ModelSettings, Runner, RunContextWrapper, function_tool
from agents.stream_events import StreamEvent

from themes import (
    ALL_BUILTIN_AGENT_TOOLS,
    FINANCE_DECK_AGENT_TOOLS,
    STANDARD_DECK_AGENT_TOOLS,
    allowed_agent_tools_for_theme,
    get_theme,
    merge_theme,
)

from intro_slide import create_intro_slide
from bar_chart_slide import create_bar_chart_slide
from bulleted_boxes_slide import create_bulleted_boxes_slide
from numeric_highlight_slide import create_numeric_highlight_slide
from split_bullet_slide import create_split_bullet_slide
from table_slide import create_table_slide
from sources_cited_slide import create_sources_cited_slide
from finance_dual_column_pie_slide import create_finance_dual_column_pie_slide
from finance_narrative_table_slide import create_finance_narrative_table_slide
from finance_dual_column_bar_slide import create_finance_dual_column_bar_slide
from finance_quad_slide import create_finance_quad_slide
from finance_tear_sheet_slide import create_finance_tear_sheet_slide

DEFAULT_MODEL = os.environ.get("EASYPRES_MODEL", "gpt-5.2")


def _default_deck_theme() -> dict:
    return get_theme("default")


@dataclass
class PresentationContext:
    prs: Presentation = field(default_factory=Presentation)
    output_path: str = "output.pptx"
    theme: dict = field(default_factory=_default_deck_theme)


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


class FinanceBulletItem(BaseModel):
    lead: Optional[str] = Field(default=None, description="Bold lead-in before the colon")
    body: str = Field(default="", description="Sentence body after the colon")
    cites: list[int] = Field(
        default_factory=list,
        description="Footnote numbers rendered as [n]; each needs a matching CitationLink on the slide tool",
    )


class FinanceColumn(BaseModel):
    heading: str = Field(description="Blue section title above this column")
    bullets: list[Union[str, FinanceBulletItem]] = Field(
        description="Bullet lines: plain strings or structured lead/body/cites",
    )


class FinancePieSlice(BaseModel):
    label: str = Field(description="Slice / category label")
    value: float = Field(description="Numeric value (share, percent, or count)")


class FinancePieChartSpec(BaseModel):
    title: str = Field(default="", description="Subtitle above the pie")
    slices: list[FinancePieSlice] = Field(
        default_factory=list,
        description="Pie segments as label + value pairs",
    )
    show_legend: bool = True


class FinanceChart1D(BaseModel):
    title: str = Field(default="", description="Chart title")
    categories: list[str] = Field(default_factory=list)
    values: list[float] = Field(default_factory=list)


class FinanceNarrativeSection(BaseModel):
    heading: str = ""
    bullets: list[Union[str, FinanceBulletItem]] = Field(default_factory=list)


class FinanceTableSpec(BaseModel):
    heading: str = Field(default="", description="Blue title above the table")
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class FinanceQuadTextBlock(BaseModel):
    heading: str = ""
    bullets: list[Union[str, FinanceBulletItem]] = Field(default_factory=list)


class FinanceLeftBlock(BaseModel):
    type: str = Field(description="'section' (heading+bullets) or 'table'")
    heading: str = ""
    bullets: list[Union[str, FinanceBulletItem]] = Field(default_factory=list)
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class CitationLink(BaseModel):
    ref_id: int = Field(description="Citation number matching FinanceBulletItem.cites entries")
    url: str = Field(description="https URL opened when the presenter clicks [ref_id]")


# ---------------------------------------------------------------------------
# Source extraction
# ---------------------------------------------------------------------------

MARKDOWN_LINK_RE = re.compile(r"\[([^\]]+)\]\((https?://[^\s)]+)\)")
RAW_URL_RE = re.compile(r"https?://[^\s<>()\]]+")


def _normalize_source_url(url: str) -> str:
    """Normalize URLs for deterministic de-duplication."""
    parsed = urlparse(url.strip())
    scheme = parsed.scheme.lower()
    netloc = parsed.netloc.lower()
    path = parsed.path or ""
    if path != "/":
        path = path.rstrip("/")
    return urlunparse((scheme, netloc, path, "", parsed.query, ""))


def _derive_source_title(url: str) -> str:
    """Generate a readable fallback title from a raw URL."""
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    if host.startswith("www."):
        host = host[4:]
    path = unquote(parsed.path.strip("/"))
    if not path:
        return host
    path_parts = [part.replace("-", " ").replace("_", " ") for part in path.split("/") if part]
    return f"{host} - {' / '.join(path_parts)}"


def _dump_finance_bullet(b: Union[str, FinanceBulletItem, dict]) -> Union[str, dict]:
    if isinstance(b, str):
        return b
    if isinstance(b, FinanceBulletItem):
        return b.model_dump()
    return b


def _finance_column_dict(c: FinanceColumn) -> dict:
    return {"heading": c.heading, "bullets": [_dump_finance_bullet(x) for x in c.bullets]}


def _citation_links_to_dict(links: Optional[list[CitationLink]]) -> Optional[dict[int, str]]:
    """Build ref_id -> URL map; duplicate ref_ids keep the last URL."""
    if not links:
        return None
    out: dict[int, str] = {}
    for link in links:
        out[link.ref_id] = link.url.strip()
    return out


def _extract_sources_from_content(content: str) -> list[dict[str, str]]:
    """Extract linked sources from markdown links and raw URLs."""
    sources: list[dict[str, str]] = []
    seen_urls: set[str] = set()

    for match in MARKDOWN_LINK_RE.finditer(content):
        title = match.group(1).strip()
        link = match.group(2).strip()
        normalized = _normalize_source_url(link)
        if title and normalized and normalized not in seen_urls:
            seen_urls.add(normalized)
            sources.append({"title": title, "link": link})

    markdown_urls = {match.group(2).strip() for match in MARKDOWN_LINK_RE.finditer(content)}
    for match in RAW_URL_RE.finditer(content):
        link = match.group(0).rstrip(".,;:")
        if link in markdown_urls:
            continue
        normalized = _normalize_source_url(link)
        if normalized and normalized not in seen_urls:
            seen_urls.add(normalized)
            sources.append({"title": _derive_source_title(link), "link": link})

    return sources


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
    create_intro_slide(
        ctx.context.prs,
        title=title,
        subtitle=subtitle,
        date=date,
        theme=ctx.context.theme,
    )
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
        theme=ctx.context.theme,
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
        theme=ctx.context.theme,
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
    create_bulleted_boxes_slide(
        ctx.context.prs,
        title=title,
        cards=cards_dicts,
        theme=ctx.context.theme,
    )
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
        theme=ctx.context.theme,
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
        theme=ctx.context.theme,
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
        theme=ctx.context.theme,
    )
    return f"Added table slide: '{title}'"


@function_tool
def add_finance_dual_column_pie_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    columns: list[FinanceColumn],
    sources_line: str,
    chart: Optional[FinancePieChartSpec] = None,
    logo_text: Optional[str] = None,
    citations: Optional[list[CitationLink]] = None,
) -> str:
    """Two-column finance narrative with optional pie in the lower-right (equity-research style)."""
    cols = [_finance_column_dict(c) for c in columns]
    ch = None
    if chart:
        ch = {
            "title": chart.title,
            "slices": {s.label: s.value for s in chart.slices},
            "show_legend": chart.show_legend,
        }
    create_finance_dual_column_pie_slide(
        ctx.context.prs,
        title=title,
        columns=cols,
        chart=ch,
        sources_line=sources_line,
        logo_text=logo_text,
        citation_urls=_citation_links_to_dict(citations),
        theme=ctx.context.theme,
    )
    return f"Added finance dual-column + pie slide: '{title}'"


@function_tool
def add_finance_narrative_table_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    left_sections: list[FinanceNarrativeSection],
    table: FinanceTableSpec,
    sources_line: str,
    logo_text: Optional[str] = None,
    citations: Optional[list[CitationLink]] = None,
) -> str:
    """Left-column narrative with a comparison table on the right (navy header, zebra rows)."""
    sections = [
        {"heading": s.heading, "bullets": [_dump_finance_bullet(b) for b in s.bullets]}
        for s in left_sections
    ]
    create_finance_narrative_table_slide(
        ctx.context.prs,
        title=title,
        left_sections=sections,
        table=table.model_dump(),
        sources_line=sources_line,
        logo_text=logo_text,
        citation_urls=_citation_links_to_dict(citations),
        theme=ctx.context.theme,
    )
    return f"Added finance narrative + table slide: '{title}'"


@function_tool
def add_finance_dual_column_bar_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    columns: list[FinanceColumn],
    chart: FinanceChart1D,
    sources_line: str,
    logo_text: Optional[str] = None,
    citations: Optional[list[CitationLink]] = None,
) -> str:
    """Two-column bullets on top and a full-width column chart below."""
    cols = [_finance_column_dict(c) for c in columns]
    create_finance_dual_column_bar_slide(
        ctx.context.prs,
        title=title,
        columns=cols,
        chart=chart.model_dump(),
        sources_line=sources_line,
        logo_text=logo_text,
        citation_urls=_citation_links_to_dict(citations),
        theme=ctx.context.theme,
    )
    return f"Added finance dual-column + bar slide: '{title}'"


@function_tool
def add_finance_quad_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    top_left: FinanceQuadTextBlock,
    top_right: FinanceChart1D,
    bottom_left: FinanceQuadTextBlock,
    bottom_right: FinanceTableSpec,
    sources_line: str,
    logo_text: Optional[str] = None,
    citations: Optional[list[CitationLink]] = None,
) -> str:
    """2x2 layout: text, line chart, text, small table (dense operational / financial slide)."""
    create_finance_quad_slide(
        ctx.context.prs,
        title=title,
        top_left={
            "heading": top_left.heading,
            "bullets": [_dump_finance_bullet(b) for b in top_left.bullets],
        },
        top_right=top_right.model_dump(),
        bottom_left={
            "heading": bottom_left.heading,
            "bullets": [_dump_finance_bullet(b) for b in bottom_left.bullets],
        },
        bottom_right=bottom_right.model_dump(),
        sources_line=sources_line,
        logo_text=logo_text,
        citation_urls=_citation_links_to_dict(citations),
        theme=ctx.context.theme,
    )
    return f"Added finance quad slide: '{title}'"


@function_tool
def add_finance_tear_sheet_slide(
    ctx: RunContextWrapper[PresentationContext],
    title: str,
    left_blocks: list[FinanceLeftBlock],
    right_chart: FinanceChart1D,
    sources_line: str,
    logo_text: Optional[str] = None,
    citations: Optional[list[CitationLink]] = None,
) -> str:
    """Dense left column (sections and/or tables) with horizontal bar chart on the right and bottom bar."""
    blocks = []
    for b in left_blocks:
        d = b.model_dump()
        d["bullets"] = [_dump_finance_bullet(x) for x in d.get("bullets", [])]
        blocks.append(d)
    create_finance_tear_sheet_slide(
        ctx.context.prs,
        title=title,
        left_blocks=blocks,
        right_chart=right_chart.model_dump(),
        sources_line=sources_line,
        logo_text=logo_text,
        citation_urls=_citation_links_to_dict(citations),
        theme=ctx.context.theme,
    )
    return f"Added finance tear sheet slide: '{title}'"


# ---------------------------------------------------------------------------
# Agent — tool ordering and theme filtering
# ---------------------------------------------------------------------------

ORDERED_AGENT_TOOLS: list[Callable] = [
    add_intro_slide,
    add_bar_chart_slide_single,
    add_bar_chart_slide_multi,
    add_bulleted_boxes_slide,
    add_numeric_highlight_slide,
    add_split_bullet_slide,
    add_table_slide,
    add_finance_dual_column_pie_slide,
    add_finance_narrative_table_slide,
    add_finance_dual_column_bar_slide,
    add_finance_quad_slide,
    add_finance_tear_sheet_slide,
]


def _agent_tool_name(tool: Any) -> str:
    """Resolve registry id for SDK-wrapped or plain callables."""
    n = getattr(tool, "name", None)
    if isinstance(n, str) and n:
        return n
    fn = getattr(tool, "__name__", None)
    if isinstance(fn, str) and fn:
        return fn
    return type(tool).__name__


def select_agent_tools_for_theme(theme: dict) -> list[Callable]:
    """Return tool callables allowed for *theme* (stable order)."""
    allow = allowed_agent_tools_for_theme(theme)
    return [t for t in ORDERED_AGENT_TOOLS if _agent_tool_name(t) in allow]


AGENT_INSTRUCTIONS_CORE = """\
You are a professional presentation builder. Given content, create a polished \
PowerPoint presentation by adding slides one at a time using the available tools.

You ONLY have the tools exposed for this job — never assume a tool exists unless you can call it.

Deck styling (colors, fonts, chrome) is controlled by the job — do not ask for theme; use \
the tools as-is.

Guidelines:
- Always start with an intro slide (add_intro_slide) when that tool is available.
- Analyze the content and choose the best slide type for each section from the tools you have.
"""

AGENT_INSTRUCTIONS_STANDARD_LAYOUTS = """\
Standard layouts (use when available):
  * Numeric data with labels -> add_numeric_highlight_slide
  * Comparative bar/column data -> add_bar_chart_slide_single (one series) \
or add_bar_chart_slide_multi (multiple series)
  * Categorized bullet points -> add_bulleted_boxes_slide
  * Strategy, overview, or multi-section text -> add_split_bullet_slide
  * Tabular data with rows and columns -> add_table_slide
"""

AGENT_INSTRUCTIONS_FINANCE_LAYOUTS = """\
Equity-research / finance layouts (use when available):
  * Two columns of thesis bullets + pie or mix chart (lower-right) -> \
add_finance_dual_column_pie_slide (set chart to null if no pie)
  * Narrative bullets left + wide comparison table right -> add_finance_narrative_table_slide
  * Two-column bullets on top + single column chart below -> add_finance_dual_column_bar_slide
  * Four-quadrant: two text blocks, line chart, small table -> add_finance_quad_slide
  * Dense left stack (sections and/or embedded tables) + horizontal bar + footnotes -> \
add_finance_tear_sheet_slide

Finance bullets: use FinanceBulletItem with lead (bold lead-in), body, and cites \
([n] in deck) when the slide lists sourced claims; plain strings are fine for simple bullets.
Whenever you use cites on a finance slide, pass citations: a list of {ref_id, url} with one \
https URL per cited ref_id (each [n] in the deck becomes a clickable link).
Every finance slide needs a sources_line (short grey footer text).
"""

AGENT_INSTRUCTIONS_GENERAL = """\
General:
- Use clear, concise text. Summarize verbose content.
- Intro slide titles MUST be 30 characters or fewer. Bar chart and numeric highlight \
titles MUST be 39 characters or fewer. Move elaboration to subtitle or descriptor.
- Once all slides are added, respond with a brief summary of the slides you created.
"""


def build_agent_instructions(theme: dict) -> str:
    """Compose system prompt from theme allowlists."""
    allow = allowed_agent_tools_for_theme(theme)
    parts = [AGENT_INSTRUCTIONS_CORE.rstrip()]
    std_names = allow & STANDARD_DECK_AGENT_TOOLS
    if std_names - {"add_intro_slide"}:
        parts.append(AGENT_INSTRUCTIONS_STANDARD_LAYOUTS.rstrip())
    if allow & FINANCE_DECK_AGENT_TOOLS:
        parts.append(AGENT_INSTRUCTIONS_FINANCE_LAYOUTS.rstrip())
    parts.append(AGENT_INSTRUCTIONS_GENERAL.rstrip())
    return "\n\n".join(parts) + "\n"


# Default agent (full tool surface) for introspection; runtime uses select_agent_tools_for_theme.
AGENT_INSTRUCTIONS = build_agent_instructions(
    {"ALLOWED_AGENT_TOOLS": ALL_BUILTIN_AGENT_TOOLS},
)

presentation_agent = Agent[PresentationContext](
    name="Presentation Builder",
    instructions=AGENT_INSTRUCTIONS,
    model=DEFAULT_MODEL,
    model_settings=ModelSettings(parallel_tool_calls=False),
    tools=list(ORDERED_AGENT_TOOLS),
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
    passthrough_data: Optional[dict[str, Any]] = None,
) -> Callable[[StreamEvent], None]:
    """Factory that returns a StreamEvent callback.

    When *webhook_url* is ``None`` the returned callback is a no-op.
    Otherwise every event except ``raw_response_event`` is POSTed as JSON
    to the URL (fire-and-forget; errors are logged but never raised).
    If *passthrough_data* is provided, it is included in each POST body
    under the key ``passthrough_data`` (must be JSON-serializable).
    """
    if webhook_url is None:
        def _noop(event: StreamEvent) -> None:
            pass
        return _noop

    import httpx

    async def _webhook_handler(event: StreamEvent) -> None:
        if event.type == "raw_response_event":
            return
        payload = _serialize_event(event)
        if passthrough_data is not None:
            payload["passthrough_data"] = passthrough_data
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
    include_sources_cited: Optional[bool] = None,
    theme_name: str = "default",
    theme_overrides: Optional[dict[str, Any]] = None,
    on_event: Optional[Callable[[StreamEvent], None]] = None,
    webhook_url: Optional[str] = None,
    webhook_headers: Optional[dict[str, str]] = None,
    passthrough_data: Optional[dict[str, Any]] = None,
) -> str:
    """Build a presentation from content using the AI agent.

    Args:
        content: The source material to turn into slides.
        output_path: Where to write the .pptx file.
        num_slides: If provided, instruct the agent to create exactly this
                    many slides (including the intro). Otherwise the agent
                    decides on its own.
        include_sources_cited: If True, append Sources Cited slide(s). If False, \
never append. If None, use the theme's INCLUDE_SOURCES_CITED_DEFAULT.
        on_event: Callback invoked for every ``StreamEvent`` emitted during
                  the agent run. May be sync or async.  Takes precedence over
                  *webhook_url* when both are given.
        webhook_url: URL to POST each raw event to as JSON.  Ignored when
                     *on_event* is supplied.
        webhook_headers: Extra HTTP headers sent with every webhook POST.
        passthrough_data: Optional dict included in each webhook POST under
                         the key ``passthrough_data`` (JSON-serializable).
                         Ignored when *on_event* is supplied.
        theme_name: Preset from themes.get_theme (e.g. ``\"default\"``, ``\"finance\"``).
        theme_overrides: Optional shallow dict merged onto the preset (RGBColor values, etc.).

    Returns the path to the saved .pptx file.
    """
    if on_event is not None:
        handler = on_event
    else:
        handler = create_webhook_event_handler(
            webhook_url, webhook_headers, passthrough_data=passthrough_data
        )

    resolved_theme = merge_theme(get_theme(theme_name), theme_overrides)
    effective_include_sources = include_sources_cited
    if effective_include_sources is None:
        effective_include_sources = bool(
            resolved_theme.get("INCLUDE_SOURCES_CITED_DEFAULT", False),
        )

    instructions = build_agent_instructions(resolved_theme)
    if num_slides is not None:
        instructions += (
            f"\n- You MUST create exactly {num_slides} slides "
            f"(including the intro slide)."
        )

    agent_tools = select_agent_tools_for_theme(resolved_theme)
    agent = Agent[PresentationContext](
        name=presentation_agent.name,
        instructions=instructions,
        model=DEFAULT_MODEL,
        model_settings=ModelSettings(parallel_tool_calls=False),
        tools=agent_tools,
    )

    context = PresentationContext(output_path=output_path, theme=resolved_theme)
    result = Runner.run_streamed(
        agent,
        input=content,
        context=context,
        max_turns=100,
    )
    async for event in result.stream_events():
        ret = handler(event)
        if inspect.isawaitable(ret):
            await ret

    if effective_include_sources:
        extracted_sources = _extract_sources_from_content(content)
        if extracted_sources:
            create_sources_cited_slide(
                context.prs, sources=extracted_sources, theme=context.theme,
            )

    context.prs.save(context.output_path)
    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Build a PowerPoint presentation from text content.")
    parser.add_argument("content", nargs="?", default=None, help="Text content (or pass via --file)")
    parser.add_argument("--file", "-f", type=str, default=None, help="Path to a text file with the content")
    parser.add_argument("--output", "-o", type=str, default="output.pptx", help="Output .pptx path")
    parser.add_argument("--slides", "-n", type=int, default=None, help="Exact number of slides to create (including intro)")
    parser.add_argument(
        "--include-sources-cited",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Append Sources Cited slides: true/false, or omit to use theme default",
    )
    parser.add_argument(
        "--theme",
        type=str,
        default="default",
        help="Theme preset: default, finance, or a name registered via themes.register_theme",
    )
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
        include_sources_cited=args.include_sources_cited,  # None → theme INCLUDE_SOURCES_CITED_DEFAULT
        theme_name=args.theme,
    ))
    print(f"Presentation saved to: {path}")
