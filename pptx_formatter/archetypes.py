"""
The canonical layout vocabulary.

Everything in the pipeline - extraction, the Template Bank, the content
classifier - agrees on this fixed set of archetype names. The bank is keyed
by them, layouts are tagged with them, and content slides are classified
into them, so a slide can always be routed to a layout by name lookup.

Keep this list stable: bank JSON written by an earlier run refers to these
strings, so renaming one orphans banked layouts for that archetype.
"""
from __future__ import annotations

from dataclasses import dataclass

# --- canonical archetypes -------------------------------------------------

TITLE_SLIDE = "title_slide"
SECTION_HEADER = "section_header"
TITLE_ONLY = "title_only"
TITLE_AND_CONTENT = "title_and_content"
TWO_CONTENT = "two_content"
THREE_CONTENT = "three_content"
COMPARISON = "comparison"
QUOTE = "quote"
BIG_STATEMENT = "big_statement"
PICTURE_FULL = "picture_full"
PICTURE_CAPTION = "picture_caption"
TABLE = "table"
CHART = "chart"
BLANK = "blank"
CLOSING = "closing"

ALL_ARCHETYPES = [
    TITLE_SLIDE,
    SECTION_HEADER,
    TITLE_ONLY,
    TITLE_AND_CONTENT,
    TWO_CONTENT,
    THREE_CONTENT,
    COMPARISON,
    QUOTE,
    BIG_STATEMENT,
    PICTURE_FULL,
    PICTURE_CAPTION,
    TABLE,
    CHART,
    BLANK,
    CLOSING,
]

# Human-facing labels, used in reports and the layout names we generate.
ARCHETYPE_LABELS = {
    TITLE_SLIDE: "Title Slide",
    SECTION_HEADER: "Section Header",
    TITLE_ONLY: "Title Only",
    TITLE_AND_CONTENT: "Title and Content",
    TWO_CONTENT: "Two Content",
    THREE_CONTENT: "Three Content",
    COMPARISON: "Comparison",
    QUOTE: "Quote",
    BIG_STATEMENT: "Big Statement",
    PICTURE_FULL: "Full Bleed Picture",
    PICTURE_CAPTION: "Picture with Caption",
    TABLE: "Table",
    CHART: "Chart",
    BLANK: "Blank",
    CLOSING: "Closing",
}

# When the bank has no entry for an archetype and no exact structural match
# is available, fall back along this chain before giving up and generating
# one from the submission's own grid. Ordered nearest-first.
FALLBACK_CHAIN = {
    TITLE_SLIDE: [SECTION_HEADER, BIG_STATEMENT, TITLE_ONLY],
    SECTION_HEADER: [TITLE_SLIDE, BIG_STATEMENT, TITLE_ONLY],
    TITLE_ONLY: [TITLE_AND_CONTENT, SECTION_HEADER],
    TITLE_AND_CONTENT: [TITLE_ONLY, TWO_CONTENT],
    TWO_CONTENT: [COMPARISON, TITLE_AND_CONTENT],
    THREE_CONTENT: [TWO_CONTENT, TITLE_AND_CONTENT],
    COMPARISON: [TWO_CONTENT, TITLE_AND_CONTENT],
    QUOTE: [BIG_STATEMENT, SECTION_HEADER],
    BIG_STATEMENT: [QUOTE, SECTION_HEADER, TITLE_ONLY],
    PICTURE_FULL: [PICTURE_CAPTION, BLANK],
    PICTURE_CAPTION: [TITLE_AND_CONTENT, PICTURE_FULL],
    TABLE: [TITLE_AND_CONTENT, TITLE_ONLY],
    CHART: [TITLE_AND_CONTENT, TITLE_ONLY],
    BLANK: [TITLE_ONLY],
    CLOSING: [TITLE_SLIDE, SECTION_HEADER, BIG_STATEMENT],
}


# --- structural signatures ------------------------------------------------

@dataclass(frozen=True)
class Signature:
    """
    A compact structural fingerprint of a layout (or of a content slide).

    Used to pick the "structurally closest" banked layout when the bank has
    no exact archetype match, and to sanity-check a classification. Counts
    only - geometry is compared separately, since geometry is normalized to
    slide fractions and counts are not.
    """
    n_title: int = 0
    n_body: int = 0
    n_picture: int = 0
    n_table: int = 0
    n_chart: int = 0
    n_other: int = 0
    columns: int = 1          # how many side-by-side content regions
    body_text_len: int = 0    # total characters of body text (content slides only)

    def as_vector(self) -> tuple[float, ...]:
        # body_text_len is deliberately excluded: it says nothing about a
        # layout (which has no real text), so including it would make
        # layout-vs-slide comparisons meaningless.
        return (
            float(self.n_title),
            float(self.n_body),
            float(self.n_picture),
            float(self.n_table),
            float(self.n_chart),
            float(self.n_other),
            float(self.columns),
        )


# Relative importance of each signature dimension when scoring closeness.
# Pictures/tables/charts are weighted hardest because putting a table on a
# picture layout is a much worse miss than a body-count being off by one.
_WEIGHTS = (1.5, 1.0, 2.0, 2.5, 2.5, 0.5, 1.5)


def signature_distance(a: Signature, b: Signature) -> float:
    """Weighted Euclidean distance between two structural signatures."""
    va, vb = a.as_vector(), b.as_vector()
    return sum(w * (x - y) ** 2 for w, x, y in zip(_WEIGHTS, va, vb)) ** 0.5


def canonical_signature(archetype: str) -> Signature:
    """
    The signature an archetype is 'supposed' to have. Used to score a banked
    layout that carries a different archetype tag, and to seed generated
    fallback layouts.
    """
    return _CANONICAL.get(archetype, Signature(n_title=1, n_body=1))


_CANONICAL = {
    TITLE_SLIDE: Signature(n_title=1, n_body=1, columns=1),
    SECTION_HEADER: Signature(n_title=1, n_body=1, columns=1),
    TITLE_ONLY: Signature(n_title=1, columns=1),
    TITLE_AND_CONTENT: Signature(n_title=1, n_body=1, columns=1),
    TWO_CONTENT: Signature(n_title=1, n_body=2, columns=2),
    THREE_CONTENT: Signature(n_title=1, n_body=3, columns=3),
    COMPARISON: Signature(n_title=1, n_body=4, columns=2),
    QUOTE: Signature(n_title=1, n_body=1, columns=1),
    BIG_STATEMENT: Signature(n_title=1, columns=1),
    PICTURE_FULL: Signature(n_picture=1, columns=1),
    PICTURE_CAPTION: Signature(n_title=1, n_body=1, n_picture=1, columns=1),
    TABLE: Signature(n_title=1, n_table=1, columns=1),
    CHART: Signature(n_title=1, n_chart=1, columns=1),
    BLANK: Signature(columns=1),
    CLOSING: Signature(n_title=1, n_body=1, columns=1),
}


def label_for(archetype: str) -> str:
    return ARCHETYPE_LABELS.get(archetype, archetype.replace("_", " ").title())
