"""
Cloning a part and everything it depends on into another package.

Needed because some shapes are not self-contained. A picture points at an
image part; a chart points at a chart part which in turn points at an
embedded workbook, and sometimes at a color-style part, a chart-style part
and a theme override. Deep-copying only the shape's XML leaves those
references dangling and PowerPoint reports the file as corrupt.

The one subtlety is relationship ids. A chart's XML refers to its embedded
workbook by rId (`<c:externalData r:id="rId1"/>`), so the clone has to
carry the *same* rIds as the original or the reference points at the wrong
part. python-pptx assigns rIds automatically on `relate_to()`, so this
module writes relationships directly into the collection to preserve them.
"""
from __future__ import annotations

import posixpath
import re

from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.package import PartFactory, _Relationship
from pptx.opc.packuri import PackURI

# Parts that should be shared with, not copied into, the destination:
# cloning the theme or the slide master would fork the deck's identity,
# which is the opposite of what this tool is for.
SHARED_RELTYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
}


def _partname_template(partname: str) -> str:
    """
    Turn a concrete partname into a printf template for the next free slot.

    `/ppt/charts/chart1.xml` becomes `/ppt/charts/chart%d.xml`, and
    `/ppt/embeddings/Microsoft_Excel_Sheet1.xlsx` becomes
    `/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx`.
    """
    directory = posixpath.dirname(partname)
    base = posixpath.basename(partname)
    stem, dot, ext = base.partition(".")
    stem = re.sub(r"\d+$", "", stem) or "part"
    return f"{directory}/{stem}%d{dot}{ext}"


def clone_part_graph(source_part, dest_package, memo: dict | None = None):
    """
    Copy `source_part` and its dependency graph into `dest_package`.

    Returns the cloned part. `memo` maps already-cloned source parts to
    their clones, so a workbook referenced twice is copied once and a
    cyclic reference terminates.
    """
    memo = {} if memo is None else memo
    if source_part in memo:
        return memo[source_part]

    partname = dest_package.next_partname(_partname_template(str(source_part.partname)))
    new_part = PartFactory(partname, source_part.content_type, dest_package, source_part.blob)
    memo[source_part] = new_part

    for rId, rel in source_part.rels.items():
        if rel.is_external:
            target = rel.target_ref
        elif rel.reltype in SHARED_RELTYPES:
            # Point at the destination's own equivalent rather than forking it.
            target = _shared_target(rel, dest_package)
            if target is None:
                continue
        else:
            target = clone_part_graph(rel.target_part, dest_package, memo)

        _add_rel_with_id(new_part, rId, rel.reltype, target, rel.is_external)

    return new_part


def _shared_target(rel, dest_package):
    """Resolve a shared reltype against the destination package."""
    try:
        return dest_package.main_document_part.part_related_by(rel.reltype)
    except (KeyError, ValueError):
        return None


def _add_rel_with_id(part, rId: str, reltype: str, target, is_external: bool) -> None:
    """
    Register a relationship under a specific rId.

    `relate_to()` would allocate its own id; the XML we just copied already
    names this one, so it has to be preserved verbatim.
    """
    rels = part.rels
    rels._rels[rId] = _Relationship(
        rels._base_uri,
        rId,
        reltype,
        RTM.EXTERNAL if is_external else RTM.INTERNAL,
        target,
    )


def clone_graphic_frame(dest_slide, source_shape) -> str | None:
    """
    Copy a graphic-frame shape (chart, OLE object, embedded media) onto
    `dest_slide`, bringing its backing parts along.

    Returns the new relationship id the frame points at, or None when the
    shape carried no external reference and a plain XML copy was enough.
    """
    import copy

    from pptx.oxml.ns import qn

    new_el = copy.deepcopy(source_shape._element)

    # A graphic frame names its backing part by rId somewhere in its data
    # element - `c:chart/@r:id` for charts, `p:oleObj/@r:id` for OLE.
    r_id_attr = qn("r:id")
    holders = [el for el in new_el.iter() if el.get(r_id_attr)]
    if not holders:
        dest_slide.shapes._spTree.append(new_el)
        return None

    source_part = source_shape.part
    dest_part = dest_slide.part
    new_rId = None

    for holder in holders:
        old_rId = holder.get(r_id_attr)
        try:
            target = source_part.related_part(old_rId)
        except KeyError:
            continue
        cloned = clone_part_graph(target, dest_part.package)
        new_rId = dest_part.relate_to(cloned, source_part.rels[old_rId].reltype)
        holder.set(r_id_attr, new_rId)

    dest_slide.shapes._spTree.append(new_el)
    return new_rId
