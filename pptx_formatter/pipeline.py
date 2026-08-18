"""
End-to-end orchestration of Phases 1-2: Stage 1 -> Stage 2 -> Stage 3.

This is the "glue" a CLI command or an API endpoint calls; the actual logic
lives in extraction.py / layout_generator.py / formatting.py / slide_copy.py
so each stage stays independently testable (matching Section 3's "each
stage's output is independently renderable and diffable" design goal).
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from .style_spec import StyleSpec
from .extraction import extract_style_spec
from .layout_generator import generate_master_layouts, list_available_layouts, DEFAULT_TEMPLATE
from .slide_copy import copy_slide_content
from .formatting import format_slide
from . import qa


def run_pipeline(
    master_pptx_path: str | Path,
    rough_content_pptx_path: str | Path,
    output_pptx_path: str | Path,
    template_path: str | Path = DEFAULT_TEMPLATE,
    content_layout_name: str = "Title and Content",
) -> dict:
    """
    Full local pipeline:
        1. Extract a StyleSpec from the submitted master.
        2. Generate a restyled master/layout set from the Template Bank.
        3. Copy each rough content slide onto that restyled deck and run
           Stage 3 formatting (typography, color, grid) on it.
        4. Run the lightweight QA checks on the result.
        5. Save the finished deck.

    Returns a report dict - this is what the API returns as JSON, and what
    the CLI prints a summary of.
    """
    master_pptx_path = Path(master_pptx_path)
    rough_content_pptx_path = Path(rough_content_pptx_path)
    output_pptx_path = Path(output_pptx_path)

    # ---- Stage 1 ----
    style_spec = extract_style_spec(master_pptx_path)

    # ---- Stage 2 ----
    out_prs = generate_master_layouts(style_spec, template_path=template_path)
    available_layouts = list_available_layouts(out_prs)
    layout = _find_layout(out_prs, content_layout_name)

    # ---- Stage 3 ----
    rough_prs = Presentation(str(rough_content_pptx_path))
    per_slide_reports = []
    copy_warnings: list[str] = []
    qa_issues: list[str] = []

    background_hex = style_spec.theme_colors.get("lt1", "FFFFFF")

    for source_slide in rough_prs.slides:
        new_slide = out_prs.slides.add_slide(layout)
        warnings = copy_slide_content(new_slide, source_slide)
        copy_warnings.extend(warnings)

        report = format_slide(new_slide, style_spec)
        per_slide_reports.append(report)

        qa_issues.extend(qa.check_contrast(new_slide, background_hex))
        qa_issues.extend(qa.check_overflow(new_slide))

    output_pptx_path.parent.mkdir(parents=True, exist_ok=True)
    out_prs.save(str(output_pptx_path))

    return {
        "style_spec": style_spec.to_json(),
        "available_layouts": available_layouts,
        "layout_used": layout.name,
        "slides_processed": len(per_slide_reports),
        "per_slide_reports": per_slide_reports,
        "copy_warnings": copy_warnings,
        "qa_issues": qa_issues,
        "output_path": str(output_pptx_path),
    }


def _find_layout(prs: Presentation, name: str):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == name:
                return layout
    # Fall back to the second layout (index 1 is "Title and Content" in the
    # default python-pptx template) rather than failing outright.
    return prs.slide_masters[0].slide_layouts[1]
