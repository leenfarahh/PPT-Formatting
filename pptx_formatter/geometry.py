"""
Rectangle arithmetic shared by the slide builder and the grid pass.

Both need to answer the same two questions before they move or copy a shape:
does this land inside something else, and does it land *on top of* something
else. Keeping the answers in one place means the builder's containment test
and the grid pass's collision test can't drift apart and start disagreeing
about whether a given pair of shapes overlaps.

Everything here works in EMU on a single coordinate space. Group children
carry their offsets in group space rather than slide space, so a rectangle
from inside a group is only ever comparable with its siblings - callers are
responsible for not mixing the two.
"""
from __future__ import annotations

# A shape is treated as living inside another when this much of its own area
# falls within it. Two thirds is deliberately loose: a rough deck's card and
# the text box sitting on it are usually flush, but hand-drawn slides leave
# the odd label hanging a few millimetres over an edge.
CONTAINMENT = 0.66

# New overlap is tolerated up to this share of the smaller shape's area.
# Nothing is ever exactly flush in a hand-built deck, and a hairline touch
# between neighbouring shapes is not what we're trying to catch.
COLLISION = 0.10


def rect(shape) -> tuple | None:
    """
    A shape's bounding box as (left, top, right, bottom), or None when its
    geometry is unresolvable.

    Placeholders inherit position from their layout and python-pptx resolves
    that for us, so an inheriting placeholder still reports a real box. A
    None here means genuinely unknown, which callers must treat as "don't
    reason about this shape" rather than as a box at the origin.
    """
    left, top = shape.left, shape.top
    width, height = shape.width, shape.height
    if None in (left, top, width, height):
        return None
    return (left, top, left + width, top + height)


def area(box: tuple) -> int:
    return max(0, box[2] - box[0]) * max(0, box[3] - box[1])


def intersection(a: tuple, b: tuple) -> int:
    """Area common to both boxes; 0 when they don't meet."""
    left, top = max(a[0], b[0]), max(a[1], b[1])
    right, bottom = min(a[2], b[2]), min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def contains(container: tuple, inner: tuple, threshold: float = CONTAINMENT) -> bool:
    """True when `threshold` of `inner`'s own area falls inside `container`."""
    inner_area = area(inner)
    if inner_area <= 0:
        return False
    return intersection(container, inner) / inner_area >= threshold


def collides(box: tuple, occupied, threshold: float = COLLISION) -> bool:
    """
    True when `box` covers more than `threshold` of any occupied box, or any
    occupied box covers more than `threshold` of it.

    The share is measured against the *smaller* of the two, so a small icon
    landing in the middle of a large panel counts as a collision even though
    it covers only a sliver of the panel.
    """
    box_area = area(box)
    for other in occupied:
        overlap = intersection(box, other)
        if not overlap:
            continue
        smaller = min(box_area, area(other))
        if smaller <= 0:
            continue
        if overlap / smaller > threshold:
            return True
    return False
