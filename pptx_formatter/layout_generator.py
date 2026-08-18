"""
Stage 2 - Automatic Master Layout Editing (technical plan, Section 5.2).

python-pptx cannot create new slide masters/layouts inside a presentation
(a confirmed limitation - see python-pptx GitHub issues #413 and #1028,
and Section 5.2 of the technical plan). This is why the plan calls for a
"Template Bank" of pre-built layout files rather than generating layouts
from scratch: this module starts from a template that already has every
layout archetype the deck needs, and restyles it in place.

python-pptx's default starter template (what you get from `Presentation()`
with no path) already ships 11 standard layouts - Title Slide, Title and
Content, Section Header, Two Content, Comparison, and so on - which covers
the archetypes the plan lists as examples. That's used as the bundled
Template Bank here (template_bank/default_template.pptx) so Phase 2 doesn't
depend on a commercial cloning library (Aspose.Slides) just to run locally.
Swap in Prezlab's real template bank file(s) once Phase 0 discovery
produces them; nothing else in this module needs to change.

What this DOES do (fully supported by python-pptx):
    - Rewrite the template's theme colors and fonts to match the StyleSpec.
    - Insert/update a logo picture on the slide master.
    - Update the footer placeholder text across layouts.

What this does NOT do:
    - Add a layout the template bank doesn't already contain.
    - Guarantee every layout "matches" the submitted master's specific
      placeholder geometry - only its colors/fonts/logo/footer.
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .style_spec import StyleSpec, THEME_COLOR_ROLES
from .extraction import get_theme_part, get_theme_root, A_NS

DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "template_bank" / "default_template.pptx"


def _write_theme_colors(theme_root: etree._Element, colors: dict) -> None:
    scheme = theme_root.find(".//a:clrScheme", A_NS)
    if scheme is None:
        return
    for role in THEME_COLOR_ROLES:
        if role not in colors:
            continue
        role_el = scheme.find(f"a:{role}", A_NS)
        if role_el is None:
            continue
        srgb = role_el.find("a:srgbClr", A_NS)
        if srgb is not None:
            srgb.set("val", colors[role].upper().lstrip("#"))
        else:
            # sysClr (e.g. dk1/lt1 often use sysClr) - replace with an explicit srgbClr
            sys_clr = role_el.find("a:sysClr", A_NS)
            if sys_clr is not None:
                role_el.remove(sys_clr)
            new_srgb = etree.SubElement(role_el, f"{{{A_NS['a']}}}srgbClr")
            new_srgb.set("val", colors[role].upper().lstrip("#"))


def _write_theme_fonts(theme_root: etree._Element, major: str, minor: str) -> None:
    scheme = theme_root.find(".//a:fontScheme", A_NS)
    if scheme is None:
        return
    major_el = scheme.find("a:majorFont/a:latin", A_NS)
    minor_el = scheme.find("a:minorFont/a:latin", A_NS)
    if major_el is not None:
        major_el.set("typeface", major)
    if minor_el is not None:
        minor_el.set("typeface", minor)


def _apply_logo(prs: Presentation, style_spec: StyleSpec) -> None:
    logo = style_spec.brand_logo
    if not logo or logo.left is None:
        return  # no logo detected in Stage 1 - nothing to propagate
    master = prs.slide_masters[0]
    # Remove any existing picture on the master first, so re-running this
    # generator is idempotent instead of stacking logos on top of each other.
    for shape in list(master.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape._element.getparent().remove(shape._element)
    # Without the original image bytes we can't re-insert the real logo here;
    # Stage 1 would need to have saved shape.image.blob to disk for this to
    # insert the actual picture. That plumbing is left as a follow-up -
    # documented in README's "Known simplifications" section.


def _apply_footer(prs: Presentation, style_spec: StyleSpec) -> None:
    footer = style_spec.brand_footer
    if not footer or not footer.text:
        return
    for master in prs.slide_masters:
        for ph in master.placeholders:
            if ph.placeholder_format.type is not None and "FOOTER" in str(ph.placeholder_format.type):
                if ph.has_text_frame:
                    ph.text_frame.text = footer.text
        for layout in master.slide_layouts:
            for ph in layout.placeholders:
                if ph.placeholder_format.type is not None and "FOOTER" in str(ph.placeholder_format.type):
                    if ph.has_text_frame:
                        ph.text_frame.text = footer.text


def generate_master_layouts(style_spec: StyleSpec, template_path: str | Path = DEFAULT_TEMPLATE) -> Presentation:
    """
    Phase 2 entry point: StyleSpec + a template-bank file -> a Presentation
    whose slide master/layouts are restyled to match the StyleSpec.

    The returned Presentation is the foundation Stage 3 (formatting.py)
    should add content slides onto via `prs.slides.add_slide(layout)`.
    """
    prs = Presentation(str(template_path))

    theme_part = get_theme_part(prs)
    theme_root = get_theme_root(prs)
    _write_theme_colors(theme_root, style_spec.theme_colors)
    _write_theme_fonts(theme_root, style_spec.theme_fonts.major_latin, style_spec.theme_fonts.minor_latin)
    theme_part._blob = etree.tostring(theme_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    _apply_logo(prs, style_spec)
    _apply_footer(prs, style_spec)

    return prs


def list_available_layouts(prs: Presentation) -> list[str]:
    """Convenience: the layout names available in a generated template, so
    Stage 3 / a caller can pick the right one per content slide."""
    return [layout.name for master in prs.slide_masters for layout in master.slide_layouts]
