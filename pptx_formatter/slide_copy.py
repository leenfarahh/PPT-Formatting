"""
Copying a slide's shapes from a source presentation into a new slide in a
different (destination) presentation.

python-pptx has no built-in "copy this slide into another deck" operation,
and this is genuinely fiddly in OOXML: shapes that only reference inline
content (text boxes, autoshapes, tables) can be deep-copied directly, but
shapes that reference an external part - pictures (an image part) and
charts (a whole separate chart XML part + embedded workbook) - need that
related part carried over too, or the copy is a dangling reference.

What's handled here:
    - Text boxes, autoshapes, tables, grouped shapes: deep-copied directly.
    - Pictures: re-embedded via `add_picture` using the source image bytes
      (this is fully supported by python-pptx, not a workaround).

What's explicitly NOT handled (and skipped with a warning):
    - Charts. A chart is a graphicFrame pointing at a separate chart part
      (plus an embedded workbook) - copying that correctly is real work,
      and it's Phase 3 scope ("Charts, tables & icons", Section 5.3.4 /
      10.4 of the technical plan), not Phase 1-2. A placeholder text box
      is inserted instead so the slide doesn't just silently lose content.
"""
from __future__ import annotations

import copy
import io
import logging

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

logger = logging.getLogger(__name__)


def _is_chart_graphic_frame(shape) -> bool:
    return shape.shape_type == MSO_SHAPE_TYPE.CHART


def copy_slide_content(dest_slide, source_slide) -> list[str]:
    """
    Copy every shape from `source_slide` onto `dest_slide`.

    Returns a list of human-readable warnings (e.g. charts that were
    skipped) so a caller (CLI/API) can surface them instead of failing
    silently.
    """
    warnings: list[str] = []

    for shape in source_slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_blob = shape.image.blob
            dest_slide.shapes.add_picture(
                io.BytesIO(image_blob),
                Emu(shape.left), Emu(shape.top),
                Emu(shape.width), Emu(shape.height),
            )
            continue

        if _is_chart_graphic_frame(shape):
            warnings.append(
                f"Skipped chart shape '{shape.name}' - chart copying is Phase 3 scope, not implemented here."
            )
            tb = dest_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
            tb.text_frame.text = f"[Chart placeholder: '{shape.name}' - see Phase 3]"
            continue

        # Text boxes, autoshapes, tables, groups: deep-copy the XML element
        # and re-parent it onto the destination slide's shape tree.
        new_el = copy.deepcopy(shape._element)
        dest_slide.shapes._spTree.append(new_el)

    return warnings
