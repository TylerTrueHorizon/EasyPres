"""
Runtime presentation themes (colors, typography, chrome).

Schema
------
Required (compatibility with existing slides):
    PRIMARY_COLOR, SECONDARY_COLOR, TERTIARY_COLOR, NEUTRAL_DARK, NEUTRAL_LIGHT

Deck builder (presentation agent):
    ALLOWED_AGENT_TOOLS — frozenset of tool function names (strings). Limits which \
add_* tools are registered for this theme. Built-ins: STANDARD_DECK_AGENT_TOOLS (no \
finance layouts) and FINANCE_DECK_AGENT_TOOLS (intro + finance layouts only). Omit or \
override via merge_theme when registering custom themes.
    INCLUDE_SOURCES_CITED_DEFAULT — if True, append Sources Cited slides when the caller \
does not pass include_sources_cited (None) to src.agents.presentation_agent.build_presentation.

Optional (finance layouts and chrome — readers should use theme.get with fallbacks):
    FONT_FAMILY — default "Albert Sans"
    TITLE_TEXT_COLOR — slide titles (fallback PRIMARY_COLOR in legacy intro, or NEUTRAL_DARK)
    SECTION_HEADING_COLOR — subsection titles (fallback SECONDARY_COLOR)
    CITATION_COLOR — inline [n] references (fallback SECONDARY_COLOR)
    TABLE_HEADER_BG / TABLE_HEADER_TEXT — comparison tables
    TABLE_FIRST_COL_BG / TABLE_ZEBRA_ALT — row/column shading
    TOP_BAR_COLOR, TOP_BAR_HEIGHT_IN, SHOW_TOP_BAR
    BOTTOM_BAR_COLOR, BOTTOM_BAR_HEIGHT_IN, SHOW_BOTTOM_BAR
    FOOTER_RULE_COLOR, SOURCES_TEXT_COLOR
    CHART_LINE_COLOR — line charts (fallback PRIMARY_COLOR)
    LOGO_DEFAULT_TEXT — e.g. "AI"
    LOGO_UNDERLINE_COLOR — fallback SECTION_HEADING_COLOR
"""

from __future__ import annotations

from typing import Callable, FrozenSet, Optional, Union

from pptx.dml.color import RGBColor

THEME_DEFAULT = "default"
THEME_FINANCE = "finance"

# Agent tool names (must match @function_tool function __name__)

TOOL_ADD_INTRO = "add_intro_slide"
TOOL_ADD_BAR_SINGLE = "add_bar_chart_slide_single"
TOOL_ADD_BAR_MULTI = "add_bar_chart_slide_multi"
TOOL_ADD_BULLETED_BOXES = "add_bulleted_boxes_slide"
TOOL_ADD_NUMERIC_HIGHLIGHT = "add_numeric_highlight_slide"
TOOL_ADD_SPLIT_BULLET = "add_split_bullet_slide"
TOOL_ADD_TABLE = "add_table_slide"
TOOL_ADD_FINANCE_PIE = "add_finance_dual_column_pie_slide"
TOOL_ADD_FINANCE_NARRATIVE_TABLE = "add_finance_narrative_table_slide"
TOOL_ADD_FINANCE_BAR = "add_finance_dual_column_bar_slide"
TOOL_ADD_FINANCE_QUAD = "add_finance_quad_slide"
TOOL_ADD_FINANCE_TEAR = "add_finance_tear_sheet_slide"

STANDARD_DECK_AGENT_TOOLS: FrozenSet[str] = frozenset({
    TOOL_ADD_INTRO,
    TOOL_ADD_BAR_SINGLE,
    TOOL_ADD_BAR_MULTI,
    TOOL_ADD_BULLETED_BOXES,
    TOOL_ADD_NUMERIC_HIGHLIGHT,
    TOOL_ADD_SPLIT_BULLET,
    TOOL_ADD_TABLE,
})

FINANCE_DECK_AGENT_TOOLS: FrozenSet[str] = frozenset({
    TOOL_ADD_INTRO,
    TOOL_ADD_FINANCE_PIE,
    TOOL_ADD_FINANCE_NARRATIVE_TABLE,
    TOOL_ADD_FINANCE_BAR,
    TOOL_ADD_FINANCE_QUAD,
    TOOL_ADD_FINANCE_TEAR,
})

ALL_BUILTIN_AGENT_TOOLS: FrozenSet[str] = STANDARD_DECK_AGENT_TOOLS | FINANCE_DECK_AGENT_TOOLS


def _coerce_allowed_agent_tools(val: Union[frozenset, set, list, tuple, None]) -> Optional[frozenset]:
    """Normalize theme override into frozenset of tool names."""
    if val is None:
        return None
    if isinstance(val, frozenset):
        return val
    return frozenset(str(x) for x in val)


def allowed_agent_tools_for_theme(theme: dict) -> frozenset:
    """Tools the presentation agent may expose for this theme."""
    raw = theme.get("ALLOWED_AGENT_TOOLS")
    coerced = _coerce_allowed_agent_tools(raw)
    if coerced is not None:
        return coerced
    # Custom theme without explicit allowlist: permit everything we know about
    return ALL_BUILTIN_AGENT_TOOLS


def default_theme() -> dict:
    return {
        "PRIMARY_COLOR": RGBColor(0, 51, 102),
        "SECONDARY_COLOR": RGBColor(0, 174, 239),
        "TERTIARY_COLOR": RGBColor(255, 127, 0),
        "NEUTRAL_DARK": RGBColor(51, 51, 51),
        "NEUTRAL_LIGHT": RGBColor(242, 242, 242),
        "FONT_FAMILY": "Albert Sans",
        "SHOW_TOP_BAR": False,
        "SHOW_BOTTOM_BAR": False,
        "ALLOWED_AGENT_TOOLS": STANDARD_DECK_AGENT_TOOLS,
        "INCLUDE_SOURCES_CITED_DEFAULT": False,
    }


def finance_theme() -> dict:
    royal = RGBColor(0, 71, 171)
    navy = RGBColor(0, 32, 71)
    black = RGBColor(0, 0, 0)
    cite_blue = RGBColor(0, 102, 204)
    return {
        "PRIMARY_COLOR": navy,
        "SECONDARY_COLOR": royal,
        "TERTIARY_COLOR": RGBColor(255, 127, 0),
        "NEUTRAL_DARK": black,
        "NEUTRAL_LIGHT": RGBColor(242, 242, 242),
        "FONT_FAMILY": "Arial",
        "TITLE_TEXT_COLOR": black,
        "SECTION_HEADING_COLOR": royal,
        "CITATION_COLOR": cite_blue,
        "TABLE_HEADER_BG": navy,
        "TABLE_HEADER_TEXT": RGBColor(255, 255, 255),
        "TABLE_FIRST_COL_BG": RGBColor(230, 230, 230),
        "TABLE_ZEBRA_ALT": RGBColor(245, 245, 245),
        "TABLE_ZEBRA_BASE": RGBColor(255, 255, 255),
        "TOP_BAR_COLOR": royal,
        "TOP_BAR_HEIGHT_IN": 0.14,
        "SHOW_TOP_BAR": True,
        "BOTTOM_BAR_COLOR": navy,
        "BOTTOM_BAR_HEIGHT_IN": 0.14,
        "SHOW_BOTTOM_BAR": False,
        "FOOTER_RULE_COLOR": RGBColor(200, 200, 200),
        "SOURCES_TEXT_COLOR": RGBColor(120, 120, 120),
        "CHART_LINE_COLOR": navy,
        "SHOW_LOGO": False,
        "LOGO_DEFAULT_TEXT": "AI",
        "LOGO_UNDERLINE_COLOR": royal,
        "ALLOWED_AGENT_TOOLS": FINANCE_DECK_AGENT_TOOLS,
        "INCLUDE_SOURCES_CITED_DEFAULT": True,
    }


_PRESETS: dict[str, Callable[[], dict]] = {
    THEME_DEFAULT: default_theme,
    THEME_FINANCE: finance_theme,
}


def get_theme(name: str) -> dict:
    """Resolve a named preset. Unknown names fall back to *default*."""
    key = (name or THEME_DEFAULT).strip().lower()
    factory = _PRESETS.get(key)
    if factory is None:
        return default_theme()
    return factory()


def merge_theme(base: dict, overrides: Optional[dict]) -> dict:
    """Shallow merge; values from *overrides* replace *base* (skip None values)."""
    if not overrides:
        return dict(base)
    out = dict(base)
    for k, v in overrides.items():
        if v is not None:
            out[k] = v
    return out


def register_theme(name: str, factory: Callable[[], dict]) -> None:
    """Register an additional preset at runtime (lowercase key)."""
    _PRESETS[name.strip().lower()] = factory


def font_family(theme: dict) -> str:
    """Body/title font name for the active theme."""
    return theme.get("FONT_FAMILY", "Albert Sans")
