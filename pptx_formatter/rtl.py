"""
Bilingual text handling: Arabic and other right-to-left scripts.

Prezlab decks run in both English and Arabic, often on the same slide, so
formatting can't assume a single direction. Three things have to be right
for Arabic to render correctly in PowerPoint:

1.  **Paragraph direction.** `<a:pPr rtl="1"/>` is what makes punctuation,
    digits and embedded Latin words order correctly. Without it a
    paragraph of Arabic still displays, but parentheses and numerals land
    on the wrong side.

2.  **Alignment.** RTL paragraphs read from the right, so a paragraph the
    designer left at "left" should become "right". A paragraph explicitly
    centered stays centered - that's a design decision, not a default.

3.  **The complex-script typeface.** Arabic resolves through `<a:cs>`, not
    `<a:latin>`. Setting only the Latin face leaves Arabic glyphs falling
    back to whatever the system picks, which is how decks end up with two
    different Arabic fonts on one slide.

Direction is decided per paragraph rather than per slide, so a bilingual
slide gets each paragraph handled on its own terms.
"""
from __future__ import annotations

from pptx.oxml.ns import qn

# Unicode blocks for right-to-left scripts. Arabic is the one that matters
# for Prezlab; Hebrew and the Arabic presentation forms are included so
# mixed or pre-composed text isn't misread as left-to-right.
_RTL_RANGES = (
    (0x0590, 0x05FF),   # Hebrew
    (0x0600, 0x06FF),   # Arabic
    (0x0700, 0x074F),   # Syriac
    (0x0750, 0x077F),   # Arabic Supplement
    (0x0780, 0x07BF),   # Thaana
    (0x08A0, 0x08FF),   # Arabic Extended-A
    (0xFB1D, 0xFDFF),   # Hebrew/Arabic presentation forms A
    (0xFE70, 0xFEFF),   # Arabic presentation forms B
)

_ARABIC_RANGES = (
    (0x0600, 0x06FF), (0x0750, 0x077F), (0x08A0, 0x08FF),
    (0xFB50, 0xFDFF), (0xFE70, 0xFEFF),
)


def _in_ranges(ch: str, ranges) -> bool:
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in ranges)


def is_rtl_char(ch: str) -> bool:
    return _in_ranges(ch, _RTL_RANGES)


def contains_arabic(text: str) -> bool:
    return any(_in_ranges(ch, _ARABIC_RANGES) for ch in text or "")


def rtl_ratio(text: str) -> float:
    """
    Share of strongly-directional characters that are RTL.

    Digits, spaces and punctuation are neutral - they take the direction of
    the surrounding text - so counting them would dilute the signal and
    misclassify a mostly-Arabic line that happens to contain a long number.
    """
    strong = [ch for ch in (text or "") if ch.isalpha()]
    if not strong:
        return 0.0
    return sum(1 for ch in strong if is_rtl_char(ch)) / len(strong)


def is_rtl_text(text: str, threshold: float = 0.4) -> bool:
    """
    Whether a string should be laid out right-to-left.

    The threshold sits below half deliberately: an Arabic sentence quoting
    an English product name is still an Arabic sentence, and should still
    read from the right.
    """
    return rtl_ratio(text) >= threshold


# --- applying direction ---------------------------------------------------

# `a:pPr` children that must follow the elements we insert, per the schema.
_RPR_SUCCESSORS = ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")


def set_complex_script_font(run, typeface: str) -> None:
    """
    Set a run's complex-script (`a:cs`) typeface.

    python-pptx's `font.name` only reaches `a:latin`, which leaves Arabic
    glyphs unstyled, so this writes the element directly - inserted in
    schema order, since PowerPoint rejects `a:cs` in the wrong position.
    """
    if not typeface:
        return
    rPr = run.font._rPr
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = rPr.makeelement(qn("a:cs"), {})
        rPr.insert_element_before(cs, *_RPR_SUCCESSORS)
    cs.set("typeface", typeface)


def set_paragraph_direction(paragraph, rtl: bool) -> None:
    """Mark a paragraph right-to-left (or explicitly left-to-right)."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1" if rtl else "0")


def align_for_direction(paragraph, rtl: bool) -> None:
    """
    Flip a paragraph's alignment to match its direction, leaving deliberate
    choices alone.

    Only left/right/unset alignments are touched: centered and justified
    text reads the same either way, and overriding those would undo the
    designer's intent.
    """
    pPr = paragraph._p.get_or_add_pPr()
    current = pPr.get("algn")
    if current in ("ctr", "just", "dist"):
        return
    pPr.set("algn", "r" if rtl else "l")


def apply_direction(paragraph, fonts, is_title: bool) -> bool:
    """
    Detect a paragraph's direction from its text and apply everything that
    follows from it: direction flag, alignment, and the right typeface on
    each run.

    Returns True when the paragraph was treated as right-to-left.
    """
    text = "".join(run.text for run in paragraph.runs)
    if not text.strip():
        return False

    rtl = is_rtl_text(text)
    set_paragraph_direction(paragraph, rtl)
    if rtl:
        align_for_direction(paragraph, True)

    # Set both faces on every run regardless of direction: a bilingual line
    # needs the Latin face for its Latin glyphs and the Arabic face for its
    # Arabic ones, and PowerPoint picks per character.
    latin = fonts.latin_for(is_title)
    cs = fonts.cs_for(is_title)
    for run in paragraph.runs:
        run.font.name = latin
        set_complex_script_font(run, cs)

    return rtl


def frame_is_rtl(text_frame) -> bool:
    """Whether a whole text frame reads predominantly right-to-left."""
    return is_rtl_text(text_frame.text or "")
