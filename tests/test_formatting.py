import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from pptx_formatter.style_spec import StyleSpec, ThemeFonts, Grid
from pptx_formatter.formatting import (
    nearest_accent_role, apply_typography, apply_color_mapping, apply_grid_alignment,
)


def make_spec():
    return StyleSpec(
        theme_colors={
            "accent1": "1F6FEB", "accent2": "D9480F", "accent3": "2F9E44",
            "dk1": "000000", "lt1": "FFFFFF",
        },
        theme_fonts=ThemeFonts(major_latin="Georgia", minor_latin="Verdana"),
        slide_width=9144000, slide_height=6858000,
        grid=Grid(margin_left=457200, margin_top=457200, margin_right=457200,
                   margin_bottom=457200, columns=12, gutter=91440),
    )


def test_nearest_accent_role_picks_the_closest_color():
    spec = make_spec()
    # very close to accent1 (1F6FEB) but not exact
    role = nearest_accent_role("1F70EC", spec)
    assert role == "accent1"


def test_apply_typography_sets_font_by_role():
    spec = make_spec()
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    slide.shapes.title.text_frame.text = "Hello"
    body = slide.placeholders[1]
    body.text_frame.text = "Body text"

    apply_typography(slide, spec)

    title_run = slide.shapes.title.text_frame.paragraphs[0].runs[0]
    body_run = body.text_frame.paragraphs[0].runs[0]
    assert title_run.font.name == "Georgia"
    assert body_run.font.name == "Verdana"
    assert title_run.font.size == Pt(28)
    assert body_run.font.size == Pt(18)


def test_apply_color_mapping_remaps_offtheme_fill():
    spec = make_spec()
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x1F, 0x70, 0xEC)  # near accent1

    notes = apply_color_mapping(slide, spec)

    assert str(shape.fill.fore_color.rgb).upper() == "1F6FEB"
    assert len(notes) == 1


def test_apply_grid_alignment_snaps_to_column_lines():
    spec = make_spec()
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.03), Inches(1.02), Inches(1), Inches(1))

    apply_grid_alignment(slide, spec)

    usable_width = spec.slide_width - spec.grid.margin_left - spec.grid.margin_right
    col_width = (usable_width - spec.grid.gutter * (spec.grid.columns - 1)) / spec.grid.columns
    col_stride = col_width + spec.grid.gutter

    rel_left = shape.left - spec.grid.margin_left
    # after snapping, rel_left should be an (near-)exact multiple of col_stride
    remainder = rel_left % col_stride
    assert remainder < 1 or (col_stride - remainder) < 1
