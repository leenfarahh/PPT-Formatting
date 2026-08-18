"""
Local command-line entry points - the fastest way to exercise Phases 1-2
without standing up the API. Run `python -m pptx_formatter.cli --help`.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .style_spec import StyleSpec
from .extraction import extract_style_spec
from .layout_generator import generate_master_layouts, list_available_layouts, DEFAULT_TEMPLATE
from .pipeline import run_pipeline


def cmd_extract(args):
    spec = extract_style_spec(args.master)
    spec.save(args.out)
    print(f"Wrote Style Spec to {args.out}")
    print(spec.to_json())


def cmd_generate_layouts(args):
    spec = StyleSpec.load(args.style_spec)
    prs = generate_master_layouts(spec, template_path=args.template)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.out)
    print(f"Wrote restyled template to {args.out}")
    print("Available layouts:", list_available_layouts(prs))


def cmd_pipeline(args):
    report = run_pipeline(
        master_pptx_path=args.master,
        rough_content_pptx_path=args.content,
        output_pptx_path=args.out,
        template_path=args.template,
        content_layout_name=args.layout,
    )
    print(f"\nSaved formatted deck to {report['output_path']}")
    print(f"Layout used: {report['layout_used']}")
    print(f"Available layouts in template: {report['available_layouts']}")
    print(f"Slides processed: {report['slides_processed']}")
    if report["copy_warnings"]:
        print("\nWarnings while copying slides:")
        for w in report["copy_warnings"]:
            print(f"  - {w}")
    if report["qa_issues"]:
        print("\nQA flags:")
        for i in report["qa_issues"]:
            print(f"  - {i}")
    else:
        print("\nNo QA flags.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pptx_formatter", description="Phases 1-2 local CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_extract = sub.add_parser("extract", help="Stage 1: master.pptx -> style_spec.json")
    p_extract.add_argument("master", help="Path to the master-slide .pptx")
    p_extract.add_argument("out", help="Path to write the Style Spec JSON to")
    p_extract.set_defaults(func=cmd_extract)

    p_gen = sub.add_parser("generate-layouts", help="Stage 2: style_spec.json -> restyled template.pptx")
    p_gen.add_argument("style_spec", help="Path to a Style Spec JSON file")
    p_gen.add_argument("out", help="Path to write the restyled template .pptx to")
    p_gen.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Template Bank file to restyle")
    p_gen.set_defaults(func=cmd_generate_layouts)

    p_pipe = sub.add_parser("pipeline", help="Stages 1-3 end to end: master + rough content -> formatted deck")
    p_pipe.add_argument("master", help="Path to the master-slide .pptx")
    p_pipe.add_argument("content", help="Path to the rough content .pptx")
    p_pipe.add_argument("out", help="Path to write the formatted deck to")
    p_pipe.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Template Bank file")
    p_pipe.add_argument("--layout", default="Title and Content", help="Layout name to place content slides on")
    p_pipe.set_defaults(func=cmd_pipeline)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
