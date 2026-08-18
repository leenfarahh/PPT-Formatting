"""
Generates the bundled starter Template Bank file (Section 5.2 of the
technical plan).

Real Prezlab layout archetypes should replace this once Phase 0 discovery
produces them - this script just captures python-pptx's own default
template (11 standard layouts: Title Slide, Title and Content, Section
Header, Two Content, Comparison, Title Only, Blank, Content with Caption,
Picture with Caption, Title and Vertical Text, Vertical Title and Text),
which is enough structural variety to demonstrate Stage 2 without depending
on a commercial cloning library.

Run: python template_bank/generate_starter_templates.py
"""
from pathlib import Path
from pptx import Presentation

OUT_PATH = Path(__file__).resolve().parent / "default_template.pptx"


def main():
    prs = Presentation()  # python-pptx's built-in default template
    prs.save(str(OUT_PATH))
    print(f"Wrote {OUT_PATH}")
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            print(f"  - layout: {layout.name}")


if __name__ == "__main__":
    main()
