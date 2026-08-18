"""
Lightweight validation checks (technical plan, Section 8.1).

These are heuristics, not a substitute for the plan's actual QA design
(which renders slides to images via LibreOffice and diffs them - Section
8.2). No rendering happens here; this only checks things that are readable
straight from the OOXML/python-pptx object model, which is cheap enough to
run on every slide without a rendering step.
"""
from __future__ import annotations

from pptx.util import Emu


def _relative_luminance(rgb: tuple[int, int, int]) -> float:
    """WCAG relative luminance, https://www.w3.org/TR/WCAG21/#dfn-relative-luminance"""
    def channel(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (channel(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(hex_a: str, hex_b: str) -> float:
    def to_rgb(h):
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    l1 = _relative_luminance(to_rgb(hex_a)) + 0.05
    l2 = _relative_luminance(to_rgb(hex_b)) + 0.05
    return max(l1, l2) / min(l1, l2)


def check_contrast(slide, background_hex: str, min_ratio: float = 4.5) -> list[str]:
    """Flag text runs whose literal RGB color fails WCAG AA (4.5:1) against
    the given background. Only checks runs with an explicit literal RGB
    color set - text using a theme color reference is assumed fine, since
    the theme itself is what Stage 2 controls."""
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if not run.text.strip():
                    continue
                try:
                    hex_color = str(run.font.color.rgb)
                except (AttributeError, TypeError):
                    continue
                ratio = contrast_ratio(hex_color, background_hex)
                if ratio < min_ratio:
                    issues.append(
                        f"Low contrast on '{shape.name}': text #{hex_color} vs background "
                        f"#{background_hex} = {ratio:.2f}:1 (need >= {min_ratio}:1)"
                    )
    return issues


def check_overflow(slide, chars_per_line_estimate: int = 40, line_height_emu: int = 228600) -> list[str]:
    """
    Rough heuristic overflow check: estimate how many lines a text frame's
    content would wrap to (using a fixed chars-per-line guess) and compare
    against the shape's height. This is deliberately conservative and will
    have false positives/negatives - a real implementation needs an actual
    text-layout pass (e.g. via the rendering step in Section 8.2), which is
    out of scope for this local Phase 1-2 harness.
    """
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if not text.strip():
            continue
        estimated_lines = max(1, sum(
            -(-len(line) // chars_per_line_estimate) if line else 1
            for line in text.split("\n")
        ))
        estimated_height = estimated_lines * line_height_emu
        if shape.height and estimated_height > shape.height * 1.15:  # 15% slack
            issues.append(
                f"'{shape.name}' may overflow: ~{estimated_lines} estimated lines "
                f"vs shape height {Emu(shape.height).inches:.2f}in"
            )
    return issues
