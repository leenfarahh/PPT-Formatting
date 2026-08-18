import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_formatter.style_spec import StyleSpec, ThemeFonts
from pptx_formatter.layout_generator import generate_master_layouts, list_available_layouts, DEFAULT_TEMPLATE
from pptx_formatter.extraction import get_theme_root, A_NS


def test_generate_master_layouts_rewrites_colors_and_fonts(tmp_path):
    spec = StyleSpec(
        theme_colors={"accent1": "112233", "accent2": "445566"},
        theme_fonts=ThemeFonts(major_latin="Georgia", minor_latin="Verdana"),
    )

    prs = generate_master_layouts(spec, template_path=DEFAULT_TEMPLATE)

    theme_root = get_theme_root(prs)
    scheme = theme_root.find(".//a:clrScheme", A_NS)
    assert scheme.find("a:accent1/a:srgbClr", A_NS).get("val") == "112233"
    assert scheme.find("a:accent2/a:srgbClr", A_NS).get("val") == "445566"

    font_scheme = theme_root.find(".//a:fontScheme", A_NS)
    assert font_scheme.find("a:majorFont/a:latin", A_NS).get("typeface") == "Georgia"
    assert font_scheme.find("a:minorFont/a:latin", A_NS).get("typeface") == "Verdana"


def test_generated_deck_keeps_multiple_layouts(tmp_path):
    spec = StyleSpec(theme_colors={"accent1": "112233"})
    prs = generate_master_layouts(spec, template_path=DEFAULT_TEMPLATE)
    layouts = list_available_layouts(prs)
    # the bundled default template ships 11 standard layouts
    assert len(layouts) >= 10
    assert "Title and Content" in layouts
    assert "Section Header" in layouts


def test_generated_deck_survives_save_and_reload(tmp_path):
    from pptx import Presentation
    spec = StyleSpec(theme_colors={"accent1": "998877"})
    prs = generate_master_layouts(spec, template_path=DEFAULT_TEMPLATE)
    out = tmp_path / "restyled.pptx"
    prs.save(str(out))

    reloaded = Presentation(str(out))
    theme_root = get_theme_root(reloaded)
    assert theme_root.find(".//a:clrScheme/a:accent1/a:srgbClr", A_NS).get("val") == "998877"
