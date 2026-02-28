import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

from intro_slide import DEFAULT_THEME
from numeric_highlight_slide import _blend_toward_white


def create_table_slide(
    prs,
    title,
    descriptor,
    headers,
    rows,
    theme=None,
):
    """
    Create a table slide with a bold title (top-left), a descriptor (top-right),
    and a styled table below.

    Args:
        prs: Presentation object
        title: Bold title text (top-left)
        descriptor: Description text (top-right)
        headers: List of column header strings
        rows: List of lists, each inner list is one row of cell values
        theme: Color theme dict (uses DEFAULT_THEME if None)
    """
    if theme is None:
        theme = DEFAULT_THEME

    slide_width = prs.slide_width
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # ── Title (top-left) ───────────────────────────────────────────────
    title_left = Inches(0.75)
    title_top = Inches(0.4)
    title_width = Inches(4.5)
    title_height = Inches(0.8)

    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.text = title
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.font.name = "Albert Sans"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = theme['PRIMARY_COLOR']

    # ── Descriptor (top-right) ─────────────────────────────────────────
    desc_width = Inches(4.5)
    desc_left = slide_width - desc_width - Inches(0.75)
    desc_top = Inches(0.4)
    desc_height = Inches(0.8)

    desc_box = slide.shapes.add_textbox(desc_left, desc_top, desc_width, desc_height)
    df = desc_box.text_frame
    df.word_wrap = True
    df.text = descriptor
    dp = df.paragraphs[0]
    dp.alignment = PP_ALIGN.LEFT
    dp.font.name = "Albert Sans"
    dp.font.size = Pt(12)
    dp.font.color.rgb = theme['NEUTRAL_DARK']

    # ── Table ──────────────────────────────────────────────────────────
    slide_height = prs.slide_height
    n_cols = len(headers)
    n_rows = len(rows) + 1
    table_width = Inches(8.5)
    table_left = (slide_width - table_width) // 2

    content_top = Inches(1.5)
    content_bottom = slide_height - Inches(0.5)
    available_height = content_bottom - content_top

    row_height = available_height / n_rows
    row_height = min(row_height, Inches(0.85))
    row_height = max(row_height, Inches(0.35))

    table_height = row_height * n_rows
    table_top = content_top + (available_height - table_height) // 2

    graphic_frame = slide.shapes.add_table(
        n_rows, n_cols, int(table_left), int(table_top),
        int(table_width), int(table_height),
    )
    table = graphic_frame.table

    table.first_row = False
    table.first_col = False
    table.horz_banding = False
    table.vert_banding = False

    header_fill = _blend_toward_white(theme['SECONDARY_COLOR'], opacity=0.35)
    data_fill = _blend_toward_white(theme['SECONDARY_COLOR'], opacity=0.10)

    # ── Header row ─────────────────────────────────────────────────────
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Albert Sans"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = theme['NEUTRAL_DARK']

    # ── Data rows ──────────────────────────────────────────────────────
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(cell_text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = data_fill
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.name = "Albert Sans"
            p.font.size = Pt(13)
            p.font.color.rgb = theme['NEUTRAL_DARK']

    return slide


if __name__ == "__main__":
    prs = Presentation()

    create_table_slide(
        prs,
        title="FY25 METRICS",
        descriptor=(
            "Highlights include strong Service growth (+14%) and "
            "Net income (+19%), driven by strategic mix shifts and "
            "robust hardware performance."
        ),
        headers=["Metric", "FY25 (USD m)", "YoY"],
        rows=[
            ["Net sales", "416,161", "+6%"],
            ["Net income", "112,010", "+19%"],
            ["Services net sales", "109,158", "+14%"],
              ["Net sales", "416,161", "+6%"],
            ["Net income", "112,010", "+19%"],
            ["Services net sales", "109,158", "+14%"],
              ["Net sales", "416,161", "+6%"],
            ["Net income", "112,010", "+19%"],
            ["Services net sales", "109,158", "+14%"],
              ["Net sales", "416,161", "+6%"],
            ["Net income", "112,010", "+19%"],
            ["Services net sales", "109,158", "+14%"]
        ],
    )

    tests_dir = os.path.join(os.path.dirname(__file__), "..", "tests")
    os.makedirs(tests_dir, exist_ok=True)

    output_path = os.path.join(tests_dir, "test_table_slide.pptx")
    prs.save(output_path)
    print(f"Presentation saved to: {output_path}")
