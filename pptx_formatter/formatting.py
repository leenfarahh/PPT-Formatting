"""
Stage 3 - Detailed Formatting (technical plan, Section 5.3) - Phase 1-2 scope only.

Implements:
    5.3.1 Typography & text boxes  (Phase 1)
    5.3.2 Color & branding application (Phase 1)
    5.3.3 Layout, spacing & grid alignment (Phase 2)

NOT implemented here (Phase 3 scope, Section 5.3.4):
    5.3.4 Charts, tables & icons - table/icon restyling is a reasonable
    Phase 1-2 stretch (see format_table below, included as a bonus since it
    doesn't need Phase 3's chart-XML work) but full chart recoloring is left
    out, matching the phase boundary in the plan.
"""
from __future__ import annotations

from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER

from .style_spec import StyleSpec, ACCENT_ROLES

TITLE_FONT_SIZE = Pt(28)
BODY_FONT_SIZE = Pt(18)
TABLE_HEADER_FONT_SIZE = Pt(14)
TABLE_BODY_FONT_SIZE = Pt(12)

TITLE_PLACEHOLDER_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}


def _hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    hex_str = hex_str.lstrip("#")
    return int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)


def _rgb_distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def nearest_accent_role(rgb_hex: str, style_spec: StyleSpec) -> str | None:
    """
    Nearest-color match (Section 5.3.2): map an arbitrary RGB found on a
    content slide to the closest theme accent role, by simple Euclidean
    distance in RGB space. Good enough to demonstrate and test the rule;
    a production version might use a perceptual color space (Lab/OKLab)
    instead of raw RGB distance.
    """
    target = _hex_to_rgb(rgb_hex)
    best_role, best_dist = None, float("inf")
    for role in ACCENT_ROLES:
        if role not in style_spec.theme_colors:
            continue
        dist = _rgb_distance(target, _hex_to_rgb(style_spec.theme_colors[role]))
        if dist < best_dist:
            best_role, best_dist = role, dist
    return best_role


def _is_title_shape(shape) -> bool:
    if not shape.is_placeholder:
        return False
    return shape.placeholder_format.type in TITLE_PLACEHOLDER_TYPES


def apply_typography(slide, style_spec: StyleSpec) -> None:
    """5.3.1 - set font family/size per shape role, normalize alignment/line spacing.

    Also forces word-wrap on and disables shape-autofit. Plain text boxes
    (python-pptx's `add_textbox`, and PowerPoint's own Insert > Text Box)
    default to `wrap="none"` + `spAutoFit`, which lets the shape silently
    grow/reposition around a single unwrapped line instead of respecting
    the box you drew - exactly the kind of overflow this stage is meant to
    fix per Section 5.3.1 ("apply autofit rules ... shrink-to-fit vs. fixed
    size"). Without this, a rough content box can render far outside its
    nominal bounds (including off the edge of the slide).
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        is_title = _is_title_shape(shape)
        font_name = style_spec.theme_fonts.major_latin if is_title else style_spec.theme_fonts.minor_latin
        font_size = TITLE_FONT_SIZE if is_title else BODY_FONT_SIZE

        tf = shape.text_frame
        tf.word_wrap = True
        if tf.auto_size != MSO_AUTO_SIZE.NONE:
            tf.auto_size = MSO_AUTO_SIZE.NONE

        for para in shape.text_frame.paragraphs:
            if para.alignment is None:
                para.alignment = PP_ALIGN.LEFT
            para.line_spacing = 1.15
            for run in para.runs:
                run.font.name = font_name
                run.font.size = font_size


def apply_color_mapping(slide, style_spec: StyleSpec) -> list[str]:
    """
    5.3.2 - map ad hoc shape fill colors to the nearest theme accent role.
    Returns a list of human-readable notes about what was remapped, mainly
    so tests (and curious callers) can see the mapping happened.
    """
    notes: list[str] = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or shape.shape_type == MSO_SHAPE_TYPE.CHART:
            continue
        try:
            fill = shape.fill
        except (AttributeError, TypeError):
            continue
        if fill is None or fill.type is None:
            continue
        # MSO_FILL.SOLID == 1
        if int(fill.type) != 1:
            continue
        try:
            current_hex = str(fill.fore_color.rgb)
        except (AttributeError, TypeError):
            continue  # already a theme-color reference, or otherwise not a literal RGB

        role = nearest_accent_role(current_hex, style_spec)
        if role and role in style_spec.theme_colors:
            new_hex = style_spec.theme_colors[role]
            if new_hex.upper() != current_hex.upper():
                fill.fore_color.rgb = RGBColor.from_string(new_hex)
                notes.append(f"{shape.name}: #{current_hex} -> #{new_hex} ({role})")
    return notes


def apply_grid_alignment(slide, style_spec: StyleSpec) -> None:
    """
    5.3.3 - snap shapes to the grid implied by the master (margins + column
    count), and enforce the margins as hard bounds. This is a simple
    snap-to-nearest-column-line implementation, not a full constraint
    solver: it moves each shape's left/top to the nearest grid line without
    trying to resolve overlaps between shapes.
    """
    grid = style_spec.grid
    usable_width = style_spec.slide_width - grid.margin_left - grid.margin_right
    if grid.columns <= 0:
        return
    col_width = (usable_width - grid.gutter * (grid.columns - 1)) / grid.columns
    col_stride = col_width + grid.gutter

    row_unit = Emu(228600)  # 0.25" vertical rhythm

    for shape in slide.shapes:
        if shape.left is None or shape.top is None:
            continue

        # Snap left to the nearest column start line, clamped inside the margins.
        rel_left = shape.left - grid.margin_left
        col_index = round(rel_left / col_stride) if col_stride else 0
        col_index = max(0, min(grid.columns - 1, col_index))
        new_left = int(grid.margin_left + col_index * col_stride)

        # Snap top to the nearest vertical rhythm line, clamped inside the top margin.
        rel_top = max(0, shape.top - grid.margin_top)
        row_index = round(rel_top / row_unit)
        new_top = int(grid.margin_top + row_index * row_unit)

        shape.left = Emu(new_left)
        shape.top = Emu(new_top)


def format_table(table_shape, style_spec: StyleSpec) -> None:
    """
    Bonus, not required for Phase 1-2: restyle a table's header row and
    fonts to the theme. Included because - unlike charts - tables don't
    need a separate XML part, so it's not blocked on Phase 3 the way chart
    recoloring is. Safe to ignore/remove if out of scope for this drop.
    """
    table = table_shape.table
    header_fill_hex = style_spec.theme_colors.get("accent1")
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            if header_fill_hex and r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(header_fill_hex)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = style_spec.theme_fonts.minor_latin
                    run.font.size = TABLE_HEADER_FONT_SIZE if r_idx == 0 else TABLE_BODY_FONT_SIZE
                    if r_idx == 0:
                        run.font.bold = True


def format_slide(slide, style_spec: StyleSpec) -> dict:
    """Run the full Phase 1-2 Stage 3 pass on one slide. Returns a small
    report dict (useful for logging / the API response / tests)."""
    apply_typography(slide, style_spec)
    color_notes = apply_color_mapping(slide, style_spec)
    apply_grid_alignment(slide, style_spec)

    table_count = 0
    for shape in slide.shapes:
        if shape.has_table:
            format_table(shape, style_spec)
            table_count += 1

    return {"color_remaps": color_notes, "tables_formatted": table_count}
