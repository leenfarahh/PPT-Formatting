"""
End-to-end demo.

Generates a synthetic master and rough deck, runs all three stages, then
re-runs the deck a second time as a returning client to show Stage 1 being
skipped from the Template Bank.

Run: python examples/run_pipeline.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_formatter.bank import TemplateBank                # noqa: E402
from pptx_formatter.pipeline import (                       # noqa: E402
    run_pipeline, format_with_banked_spec,
)

import make_sample_content                                   # noqa: E402
import make_sample_master                                    # noqa: E402

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
BANK_DIR = OUT_DIR / "bank"
CLIENT = "Northwind Holdings"


def rule(title: str) -> None:
    print(f"\n{title}\n{'-' * len(title)}")


def main():
    rule("Building sample files")
    make_sample_master.main()
    make_sample_content.main()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    bank = TemplateBank(BANK_DIR)
    out_path = OUT_DIR / "formatted_deck.pptx"

    rule("Running stages 1-3")
    report = run_pipeline(
        master_pptx_path=HERE / "sample_master.pptx",
        rough_content_pptx_path=HERE / "sample_content.pptx",
        output_pptx_path=out_path,
        bank=bank,
        client=CLIENT,
        project="Q3 Board Deck",
    )

    master = report["master"]
    print(f"Master: {master['layouts_found']} layouts read, "
          f"{master['content_slides_ignored']} content slides ignored")

    layouts = report["layouts"]
    print(f"\nLayouts in the output deck ({len(layouts['available'])} total):")
    print(f"  kept from the designer : {len(layouts['designer'])}")
    for item in layouts["designer"]:
        print(f"      {item['name']:24s} {item['archetype']}")
    print(f"  filled from the bank   : {len(layouts['from_bank'])}")
    for item in layouts["from_bank"]:
        print(f"      {item['name']:24s} {item['archetype']}  (score {item.get('match_score')})")
    print(f"  generated from the grid: {len(layouts['generated'])}")
    for item in layouts["generated"]:
        print(f"      {item['name']:24s} {item['archetype']}")

    repaired = layouts["inheritance"]["repaired"]
    total_changes = sum(len(r["changes"]) for r in repaired)
    print(f"\nInheritance: {total_changes} hardcoded value(s) rewritten as theme "
          f"references across {len(repaired)} layout(s)")
    for orphan in layouts["inheritance"]["orphans"]:
        print(f"  ! {orphan}")

    print(f"\nSlides ({report['slides_processed']}):")
    for slide in report["slides"]:
        rtl = slide["formatting"]["rtl_paragraphs"]
        extras = []
        if rtl:
            extras.append(f"{rtl} RTL para")
        if slide["formatting"]["charts_formatted"]:
            extras.append("chart")
        if slide["formatting"]["tables_formatted"]:
            extras.append("table")
        suffix = f"  [{', '.join(extras)}]" if extras else ""
        print(f"  {slide['slide']:>3}. {slide['archetype']:<18} -> "
              f"{slide['layout']:<22}{suffix}")

    if report["warnings"]:
        print("\nWarnings:")
        for warning in report["warnings"]:
            print(f"  - {warning}")

    print(f"\nQA flags: {len(report['qa_issues'])}")
    for issue in report["qa_issues"][:10]:
        print(f"  - {issue}")
    if len(report["qa_issues"]) > 10:
        print(f"  ... and {len(report['qa_issues']) - 10} more")

    rule("Re-running as a returning client (Stage 1 skipped)")
    repeat_path = OUT_DIR / "formatted_deck_repeat.pptx"
    repeat = format_with_banked_spec(
        client=CLIENT,
        rough_content_pptx_path=HERE / "sample_content.pptx",
        output_pptx_path=repeat_path,
        bank=bank,
    )
    print(f"Used banked entry '{repeat['bank_entry']}' - no master submitted")
    print(f"Slides processed: {repeat['slides_processed']}")

    rule("Done")
    print(f"Formatted deck : {out_path}")
    print(f"Repeat run     : {repeat_path}")
    print(f"Template Bank  : {BANK_DIR}")
    print("\nBoth files open and edit normally in PowerPoint.")


if __name__ == "__main__":
    main()
