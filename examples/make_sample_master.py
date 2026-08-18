"""
Generates a small synthetic "master slide" .pptx to test the pipeline
against locally, standing in for a real designer-submitted master.

It sets custom theme colors/fonts (so Stage 1 extraction has something
non-default to find) and adds a rectangle labeled "LOGO" on the slide
master (so the logo-detection heuristic in extraction.py has a picture...
actually a rectangle isn't a picture - see note below) and footer text.

Note: extraction.py's logo detector looks for a PICTURE shape on the slide
master. This sample doesn't add one, so running the pipeline against it
exercises the "no logo detected" path - that's intentional (a real logo
picture needs an actual image file, which this synthetic sample doesn't
have one of). Also worth knowing for anyone extending this: python-pptx's
`MasterShapes` (prs.slide_masters[0].shapes) doesn't support add_shape /
add_textbox / add_picture at all - those only exist on a slide's shape
tree, not a master's or layout's. That's a real python-pptx limitation
(distinct from the "can't add new layouts" one in layout_generator.py) -
worth knowing before assuming Stage 2 can decorate a master this way.

Run: python examples/make_sample_master.py
"""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

A_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

OUT_PATH = Path(__file__).resolve().parent / "sample_master.pptx"

# A small palette standing in for a client's brand colors.
CUSTOM_COLORS = {
    "accent1": "1F6FEB",
    "accent2": "D9480F",
    "accent3": "2F9E44",
    "accent4": "AE3EC9",
}
MAJOR_FONT = "Georgia"
MINOR_FONT = "Verdana"
FOOTER_TEXT = "Prezlab | Confidential"


def main():
    prs = Presentation()

    # --- rewrite theme colors/fonts, same technique as layout_generator.py ---
    master_part = prs.slide_masters[0].part
    theme_part = master_part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    scheme = root.find(".//a:clrScheme", A_NS)
    for role, hex_val in CUSTOM_COLORS.items():
        role_el = scheme.find(f"a:{role}", A_NS)
        srgb = role_el.find("a:srgbClr", A_NS)
        srgb.set("val", hex_val)
    font_scheme = root.find(".//a:fontScheme", A_NS)
    font_scheme.find("a:majorFont/a:latin", A_NS).set("typeface", MAJOR_FONT)
    font_scheme.find("a:minorFont/a:latin", A_NS).set("typeface", MINOR_FONT)
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # --- footer text on the master ---
    master = prs.slide_masters[0]
    for ph in master.placeholders:
        if ph.placeholder_format.type is not None and "FOOTER" in str(ph.placeholder_format.type):
            ph.text_frame.text = FOOTER_TEXT

    prs.save(str(OUT_PATH))
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
