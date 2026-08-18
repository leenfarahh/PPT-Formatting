import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx import Presentation

from pptx_formatter.pipeline import run_pipeline
from pptx_formatter.layout_generator import DEFAULT_TEMPLATE


def _make_master(path):
    Presentation().save(str(path))  # default theme is fine for this test


def _make_content(path):
    from pptx.util import Inches, Pt
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    box.text_frame.text = "Some rough content"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    prs.save(str(path))


def test_pipeline_end_to_end_produces_valid_pptx(tmp_path):
    master_path = tmp_path / "master.pptx"
    content_path = tmp_path / "content.pptx"
    out_path = tmp_path / "out" / "formatted.pptx"

    _make_master(master_path)
    _make_content(content_path)

    assert DEFAULT_TEMPLATE.exists(), (
        "template_bank/default_template.pptx is missing - run "
        "template_bank/generate_starter_templates.py first"
    )

    report = run_pipeline(master_path, content_path, out_path)

    assert report["slides_processed"] == 1
    assert out_path.exists()

    # the output must be a valid, re-openable .pptx
    reopened = Presentation(str(out_path))
    assert len(reopened.slides) == 1
    assert "available_layouts" in report
    assert len(report["available_layouts"]) >= 10
