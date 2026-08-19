"""
Command-line entry points. Run `python -m pptx_formatter.cli --help`.

Each stage is exposed separately as well as end to end, because the useful
thing during a rollout is being able to look at an intermediate result: the
Style Spec on its own, the layout set on its own, or just how the classifier
read a deck before committing to a reformat.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation

from . import archetypes
from .bank import TemplateBank, DEFAULT_BANK_ROOT
from .config import make_bank, settings_from_env
from .classifier import classify_deck
from .pipeline import (
    ingest_master, build_template, run_pipeline, format_with_banked_spec,
)
from .style_spec import StyleSpec


def _bank(args):
    """
    The Template Bank for this invocation.

    An explicit `--bank` always means a local directory - it's a path, so
    treating it as anything else would be surprising. Without one, the
    backend comes from the environment, so `PPTX_STORAGE_BACKEND=supabase`
    works the same here as it does for the web app.
    """
    if getattr(args, "no_bank", False):
        return None
    explicit = getattr(args, "bank", None)
    if explicit:
        return TemplateBank(explicit)
    return make_bank(settings_from_env())


def _parse_overrides(values) -> dict:
    """`--override 3=quote` becomes `{2: "quote"}` (slide numbers are 1-based)."""
    overrides = {}
    for item in values or []:
        if "=" not in item:
            raise SystemExit(f"--override expects SLIDE=archetype, got {item!r}")
        slide, archetype = item.split("=", 1)
        if archetype not in archetypes.ALL_ARCHETYPES:
            raise SystemExit(
                f"unknown archetype {archetype!r}; choose from: "
                + ", ".join(archetypes.ALL_ARCHETYPES)
            )
        overrides[int(slide) - 1] = archetype
    return overrides


# --- commands -------------------------------------------------------------

def cmd_extract(args):
    bank = _bank(args)
    spec, entry_id = ingest_master(
        args.master, bank=bank, client=args.client, project=args.project
    )
    if args.out:
        spec.save(args.out)
        print(f"Wrote Style Spec to {args.out}")
    if entry_id:
        print(f"Archived in the Template Bank as '{entry_id}'")
    print(f"Content slides ignored: {spec.meta.content_slides_ignored}")
    print(f"Layouts found: {spec.meta.layouts_found}")
    for layout in spec.layouts:
        print(f"  - {layout.name:32s} {layout.archetype}")
    missing = set(archetypes.ALL_ARCHETYPES) - spec.archetypes_present()
    if missing:
        print("Archetypes not authored by the designer (to be filled):")
        for m in sorted(missing):
            print(f"  - {m}")


def cmd_build_template(args):
    spec = StyleSpec.load(args.style_spec)
    prs, report = build_template(spec, args.base, bank=_bank(args))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.out)
    print(f"Wrote layout set to {args.out}")
    print(f"Content slides stripped: {report['slides_stripped']}")
    print(f"Designer layouts kept:   {len(report['designer_layouts'])}")
    print(f"Filled from bank:        {len(report['bank_layouts'])}")
    print(f"Generated from grid:     {len(report['generated_layouts'])}")
    _print_inheritance(report["inheritance"])


def cmd_classify(args):
    prs = Presentation(args.content)
    results = classify_deck(prs, prs.slide_width, prs.slide_height)
    for i, result in enumerate(results, start=1):
        print(f"Slide {i}: {result.archetype} (confidence {result.confidence:.2f})")
        for line in result.evidence:
            print(f"    - {line}")


def cmd_pipeline(args):
    report = run_pipeline(
        master_pptx_path=args.master,
        rough_content_pptx_path=args.content,
        output_pptx_path=args.out,
        bank=_bank(args),
        client=args.client,
        project=args.project,
        layout_overrides=_parse_overrides(args.override),
    )
    _print_report(report)
    if args.report:
        Path(args.report).write_text(
            json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print(f"\nFull JSON report written to {args.report}")


def cmd_format_existing(args):
    bank = _bank(args)
    report = format_with_banked_spec(
        client=args.client,
        rough_content_pptx_path=args.content,
        output_pptx_path=args.out,
        bank=bank,
        layout_overrides=_parse_overrides(args.override),
    )
    print(f"Stage 1 skipped; used banked entry '{report['bank_entry']}'")
    _print_report(report)


def cmd_bank_list(args):
    bank = _bank(args)
    entries = bank.list_entries()
    if not entries:
        print("Template Bank is empty. Run `extract` on a master to populate it.")
        return
    for entry in entries:
        master = "master.pptx" if bank.has_master(entry.entry_id) else "no master"
        print(f"{entry.entry_id}")
        print(f"    client:     {entry.client or '-'}   project: {entry.project or '-'}")
        print(f"    revision:   {entry.revision}   updated: {entry.updated_at}   ({master})")
        print(f"    archetypes: {', '.join(entry.archetypes) or '-'}")


def cmd_bank_show(args):
    print(_bank(args).load(args.entry_id).to_json())


def cmd_bank_refine(args):
    """Fold a designer's manual corrections back into a banked spec."""
    bank = _bank(args)
    spec = StyleSpec.load(args.style_spec)
    revision = bank.refine(args.entry_id, spec)
    print(f"Entry '{args.entry_id}' updated to revision {revision}")
    print(f"Previous version kept at {bank.entry_dir(args.entry_id) / 'revisions'}")


# --- reporting ------------------------------------------------------------

def _print_inheritance(inheritance: dict) -> None:
    repaired = inheritance.get("repaired", [])
    orphans = inheritance.get("orphans", [])
    if repaired:
        total = sum(len(r["changes"]) for r in repaired)
        print(f"Inheritance: rewrote {total} hardcoded value(s) as theme references "
              f"across {len(repaired)} layout(s)")
    else:
        print("Inheritance: no hardcoded theme values found")
    for orphan in orphans:
        print(f"  ! {orphan}")


def _print_report(report: dict) -> None:
    layouts = report["layouts"]
    print(f"\nSaved formatted deck to {report['output_path']}")
    if "master" in report:
        print(f"Content slides ignored in the master: "
              f"{report['master']['content_slides_ignored']}")
    print(f"Layouts: {len(layouts['designer'])} designer-authored, "
          f"{len(layouts['from_bank'])} from the bank, "
          f"{len(layouts['generated'])} generated")
    for item in layouts["from_bank"]:
        print(f"  bank  -> {item['name']} ({item['archetype']}, "
              f"score {item.get('match_score')})")
    for item in layouts["generated"]:
        print(f"  built -> {item['name']} ({item['archetype']})")
    _print_inheritance(layouts["inheritance"])

    print(f"\nSlides processed: {report['slides_processed']}")
    for slide in report["slides"]:
        note = ""
        if slide["resolved_archetype"] and slide["resolved_archetype"] != slide["archetype"]:
            note = f" (no {slide['archetype']} layout; used {slide['resolved_archetype']})"
        print(f"  {slide['slide']:>3}. {slide['archetype']:<18} -> "
              f"{slide['layout']}{note}")

    if report["warnings"]:
        print("\nWarnings:")
        for w in report["warnings"]:
            print(f"  - {w}")
    if report["qa_issues"]:
        print("\nQA flags:")
        for issue in report["qa_issues"]:
            print(f"  - {issue}")
    else:
        print("\nNo QA flags.")


# --- parser ---------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx_formatter",
        description="Format a rough deck to match a designer's master slide.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    def add_bank_args(p):
        p.add_argument("--bank", help=f"Template Bank directory (default: {DEFAULT_BANK_ROOT})")
        p.add_argument("--no-bank", action="store_true",
                       help="Don't read from or write to the Template Bank")

    p = sub.add_parser("extract", help="Stage 1: read a master into a Style Spec")
    p.add_argument("master", help="The designer's master .pptx (content slides are ignored)")
    p.add_argument("out", nargs="?", help="Optional path to write the Style Spec JSON to")
    p.add_argument("--client", help="Client identifier, used to key the bank")
    p.add_argument("--project", help="Project identifier, recorded in meta")
    add_bank_args(p)
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser("build-template", help="Stage 2: Style Spec -> full layout set")
    p.add_argument("style_spec", help="A Style Spec JSON file")
    p.add_argument("base", help="The master .pptx the spec came from")
    p.add_argument("out", help="Where to write the layout set .pptx")
    add_bank_args(p)
    p.set_defaults(func=cmd_build_template)

    p = sub.add_parser("classify", help="Show how each slide of a deck is classified")
    p.add_argument("content", help="The rough content .pptx")
    p.set_defaults(func=cmd_classify)

    p = sub.add_parser("pipeline", help="Stages 1-3: master + rough deck -> formatted deck")
    p.add_argument("master", help="The designer's master .pptx")
    p.add_argument("content", help="The rough content .pptx")
    p.add_argument("out", help="Where to write the formatted deck")
    p.add_argument("--client", help="Client identifier, used to key the bank")
    p.add_argument("--project", help="Project identifier")
    p.add_argument("--override", action="append", metavar="SLIDE=ARCHETYPE",
                   help="Force a slide onto an archetype, e.g. --override 4=quote")
    p.add_argument("--report", help="Write the full JSON report to this path")
    add_bank_args(p)
    p.set_defaults(func=cmd_pipeline)

    p = sub.add_parser(
        "format-existing",
        help="Format a deck for a returning client from the banked spec (skips Stage 1)",
    )
    p.add_argument("client", help="Client identifier to look up in the bank")
    p.add_argument("content", help="The rough content .pptx")
    p.add_argument("out", help="Where to write the formatted deck")
    p.add_argument("--override", action="append", metavar="SLIDE=ARCHETYPE")
    p.add_argument("--bank", help="Template Bank directory")
    p.set_defaults(func=cmd_format_existing)

    p = sub.add_parser("bank", help="Inspect and maintain the Template Bank")
    bank_sub = p.add_subparsers(dest="bank_command", required=True)

    q = bank_sub.add_parser("list", help="List archived entries")
    q.add_argument("--bank")
    q.set_defaults(func=cmd_bank_list)

    q = bank_sub.add_parser("show", help="Print an archived Style Spec")
    q.add_argument("entry_id")
    q.add_argument("--bank")
    q.set_defaults(func=cmd_bank_show)

    q = bank_sub.add_parser(
        "refine", help="Replace an entry's spec with a corrected one, keeping the old revision"
    )
    q.add_argument("entry_id")
    q.add_argument("style_spec", help="The corrected Style Spec JSON")
    q.add_argument("--bank")
    q.set_defaults(func=cmd_bank_refine)

    return parser


def _force_utf8_output() -> None:
    """
    Print UTF-8 regardless of the console's code page.

    Windows terminals default to cp1252, which cannot encode Arabic - and
    this CLI echoes slide text back in its classification evidence, so a
    bilingual deck would crash the command rather than report on it.
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):    # not a reconfigurable stream
            pass


def main(argv=None):
    _force_utf8_output()
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
    except (FileNotFoundError, LookupError, RuntimeError, ValueError) as exc:
        # Missing files, unknown clients and incomplete Supabase configuration
        # are all things the operator can fix; a traceback just buries the fix.
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
