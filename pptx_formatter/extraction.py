"""
Stage 1 - Master Slide Ingestion & Style Extraction (technical plan, Section 5.1).

Reads a designer-submitted master-slide .pptx and produces a StyleSpec:
    - theme colors (from the theme part's <a:clrScheme>)
    - theme fonts (from the theme part's <a:fontScheme>)
    - a logo, detected heuristically as a picture shape on the slide master
    - footer text, read from the master's footer placeholder if present
    - slide dimensions

python-pptx doesn't expose theme colors/fonts as a public API (this is the
gap called out in Section 5.1/7 of the plan), so this module reads the
theme XML part directly via lxml, the same way layout_generator.py writes
it back in Phase 2.
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .style_spec import StyleSpec, ThemeFonts, BrandLogo, BrandFooter, THEME_COLOR_ROLES

A_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def get_theme_part(prs: Presentation):
    """Return the OPC Part for the presentation's (first) master theme."""
    master_part = prs.slide_masters[0].part
    return master_part.part_related_by(RT.THEME)


def get_theme_root(prs: Presentation) -> etree._Element:
    """Return the parsed <a:theme> XML root for the presentation's master theme."""
    return etree.fromstring(get_theme_part(prs).blob)


def _extract_theme_colors(theme_root: etree._Element) -> dict:
    colors = {}
    scheme = theme_root.find(".//a:clrScheme", A_NS)
    if scheme is None:
        return colors
    for role in THEME_COLOR_ROLES:
        role_el = scheme.find(f"a:{role}", A_NS)
        if role_el is None:
            continue
        srgb = role_el.find("a:srgbClr", A_NS)
        if srgb is not None and srgb.get("val"):
            colors[role] = srgb.get("val").upper()
            continue
        sys_clr = role_el.find("a:sysClr", A_NS)
        if sys_clr is not None and sys_clr.get("lastClr"):
            colors[role] = sys_clr.get("lastClr").upper()
    return colors


def _extract_theme_fonts(theme_root: etree._Element) -> ThemeFonts:
    fonts = ThemeFonts()
    scheme = theme_root.find(".//a:fontScheme", A_NS)
    if scheme is None:
        return fonts
    major = scheme.find("a:majorFont/a:latin", A_NS)
    minor = scheme.find("a:minorFont/a:latin", A_NS)
    if major is not None and major.get("typeface"):
        fonts.major_latin = major.get("typeface")
    if minor is not None and minor.get("typeface"):
        fonts.minor_latin = minor.get("typeface")
    return fonts


def _find_logo(prs: Presentation) -> BrandLogo:
    """
    Heuristic: the first picture shape found on the slide master (not a
    layout) is treated as the logo. Real Prezlab masters may need a better
    heuristic (e.g. a naming convention); this is a starting point, not a
    claim of perfect detection.
    """
    master = prs.slide_masters[0]
    for shape in master.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return BrandLogo(
                image_path=None,  # the bytes live in shape.image.blob if needed later
                left=shape.left, top=shape.top, width=shape.width, height=shape.height,
            )
    return BrandLogo()


def _find_footer(prs: Presentation) -> BrandFooter:
    master = prs.slide_masters[0]
    for ph in master.placeholders:
        if ph.placeholder_format.type is not None and "FOOTER" in str(ph.placeholder_format.type):
            text = ph.text_frame.text if ph.has_text_frame else None
            return BrandFooter(text=text or None)
    return BrandFooter()


def extract_style_spec(master_pptx_path: str | Path) -> StyleSpec:
    """Stage 1 entry point: master .pptx path -> StyleSpec."""
    prs = Presentation(str(master_pptx_path))
    theme_root = get_theme_root(prs)

    spec = StyleSpec(
        theme_colors=_extract_theme_colors(theme_root),
        theme_fonts=_extract_theme_fonts(theme_root),
        brand_logo=_find_logo(prs),
        brand_footer=_find_footer(prs),
        slide_width=prs.slide_width,
        slide_height=prs.slide_height,
        meta={"source_file": str(master_pptx_path)},
    )
    return spec
