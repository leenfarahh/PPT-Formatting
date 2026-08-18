import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_formatter.extraction import extract_style_spec


def test_extracts_default_theme_colors_and_fonts(tmp_path):
    from pptx import Presentation
    sample = tmp_path / "master.pptx"
    Presentation().save(str(sample))  # unmodified default template

    spec = extract_style_spec(sample)

    assert spec.theme_colors.get("accent1") == "4F81BD"  # default Office theme accent1 (this python-pptx version)
    assert spec.theme_fonts.minor_latin  # some font name was found
    assert spec.slide_width > 0
    assert spec.slide_height > 0


def test_extracts_custom_theme_colors(tmp_path):
    from lxml import etree
    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    prs = Presentation()
    theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    root = etree.fromstring(theme_part.blob)
    accent1 = root.find(".//a:clrScheme/a:accent1/a:srgbClr", ns)
    accent1.set("val", "AA00AA")
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    sample = tmp_path / "master_custom.pptx"
    prs.save(str(sample))

    spec = extract_style_spec(sample)
    assert spec.theme_colors["accent1"] == "AA00AA"


def test_no_footer_text_is_none_not_error(tmp_path):
    from pptx import Presentation
    sample = tmp_path / "master_no_footer.pptx"
    Presentation().save(str(sample))

    spec = extract_style_spec(sample)
    # default template's footer placeholder is empty - should be None, not raise
    assert spec.brand_footer.text in (None, "")
