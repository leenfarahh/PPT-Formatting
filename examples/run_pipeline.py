"""
The fastest way to see Phases 1-2 work end to end, no server required.

Run: python examples/run_pipeline.py
(generates the sample master/content files first if they don't exist yet)
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_formatter.pipeline import run_pipeline  # noqa: E402

HERE = Path(__file__).resolve().parent
MASTER = HERE / "sample_master.pptx"
CONTENT = HERE / "sample_content.pptx"
OUT = HERE / "output" / "formatted_deck.pptx"


def main():
    if not MASTER.exists():
        print("Sample master not found - generating it now...")
        import make_sample_master
        make_sample_master.main()
    if not CONTENT.exists():
        print("Sample content not found - generating it now...")
        import make_sample_content
        make_sample_content.main()

    report = run_pipeline(MASTER, CONTENT, OUT)

    print("\n=== Style Spec extracted from the master (Stage 1) ===")
    print(report["style_spec"])

    print("\n=== Layouts available after Stage 2 restyling ===")
    print(report["available_layouts"])

    print(f"\n=== Stage 3: {report['slides_processed']} slide(s) formatted onto layout '{report['layout_used']}' ===")
    for i, slide_report in enumerate(report["per_slide_reports"], start=1):
        print(f"  Slide {i}: {json.dumps(slide_report)}")

    if report["copy_warnings"]:
        print("\nWarnings:")
        for w in report["copy_warnings"]:
            print(f"  - {w}")

    if report["qa_issues"]:
        print("\nQA flags:")
        for i in report["qa_issues"]:
            print(f"  - {i}")
    else:
        print("\nNo QA flags.")

    print(f"\nSaved formatted deck to: {report['output_path']}")


if __name__ == "__main__":
    main()
