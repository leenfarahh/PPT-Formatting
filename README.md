# PPTX Formatting Tool — Phases 1-2 Local Dev Build

Local, runnable implementation of **Phase 1** (MVP: Stage 1 extraction + Stage 3
typography/color) and **Phase 2** (Automatic Master Layout Editing: Stage 2 +
Stage 3 grid/spacing) from the technical plan. This is a dev/test harness, not
the production build — see "Known simplifications" below for what's
deliberately left out.

## What it does

1. **Stage 1 — Extraction** (`pptx_formatter/extraction.py`): reads a
   designer-submitted master `.pptx` and pulls out a `StyleSpec` (theme
   colors, fonts, logo, footer, grid/margins).
2. **Stage 2 — Master Layout Generation** (`pptx_formatter/layout_generator.py`):
   applies that `StyleSpec` to a bundled 11-layout template (python-pptx's own
   default template, used as a stand-in Template Bank), producing a restyled
   deck with the brand's colors/fonts across every layout.
3. **Stage 3 — Detailed Formatting** (`pptx_formatter/formatting.py`): for each
   rough content slide, copies it onto the restyled template and applies
   typography rules, remaps stray colors to the nearest theme accent, and
   snaps shapes to the grid.
4. **QA checks** (`pptx_formatter/qa.py`): lightweight contrast and overflow
   flags on the output.

## Setup

```bash
cd pptx-formatting-tool
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run the local demo (no API needed)

This generates two synthetic sample files (a fake brand master + a rough
content deck) and runs the full pipeline on them:

```bash
python examples/run_pipeline.py
```

Output lands at `examples/output/formatted_deck.pptx`. Open it in PowerPoint,
or preview it without PowerPoint via LibreOffice:

```bash
soffice --headless --convert-to pdf examples/output/formatted_deck.pptx
```

To try it against your own files instead, use the CLI directly:

```bash
# Stage 1 only: master.pptx -> style_spec.json
python -m pptx_formatter.cli extract path/to/master.pptx style_spec.json

# Stage 2 only: style_spec.json -> restyled template.pptx
python -m pptx_formatter.cli generate-layouts style_spec.json restyled.pptx

# Full pipeline: master + rough content -> formatted deck
python -m pptx_formatter.cli pipeline path/to/master.pptx path/to/content.pptx formatted.pptx
```

## Run the tests

```bash
pytest tests/ -v
```

11 tests covering extraction, layout generation, formatting rules, and an
end-to-end pipeline integration test.

## Run the local API

A simplified synchronous FastAPI server — this stands in for the production
job-queue architecture in the technical plan, for local testing only.

```bash
uvicorn api.main:app --reload
```

Then, in another terminal:

```bash
curl http://127.0.0.1:8000/health

curl -X POST http://127.0.0.1:8000/extract \
  -F "master=@examples/sample_master.pptx"

curl -X POST http://127.0.0.1:8000/pipeline \
  -F "master=@examples/sample_master.pptx" \
  -F "content=@examples/sample_content.pptx" \
  -o formatted.pptx
```

`/generate-layouts` takes the Style Spec JSON returned by `/extract` as a
form field named `style_spec` and returns a restyled template file.

## Project layout

```
pptx_formatter/       core library (the 3 stages + style spec + QA)
api/                  local FastAPI wrapper around the library
examples/             sample file generators + end-to-end demo script
template_bank/        bundled default template + generator script
tests/                pytest suite
```

## Known simplifications (vs. the full technical plan)

This build is scoped to prove out Phases 1-2 locally. It deliberately does
not include:

- **No job queue.** The plan's Celery/Redis async architecture is replaced
  with direct synchronous function calls — fine for one deck at a time on a
  laptop, not for concurrent production load.
- **No cloud storage or database.** Files are read/written to local disk (or
  a temp folder for the API); the plan's S3 + Postgres persistence layer
  isn't built here.
- **No Aspose.Slides fallback.** python-pptx cannot create brand-new slide
  masters/layouts inside a presentation, so Stage 2 restyles the colors,
  fonts, and footer of a bundled default 11-layout template rather than
  cloning the designer's own master layout structure. A production build
  would likely use Aspose.Slides (or similar) to actually clone the
  designer's layouts.
- **Template Bank is a placeholder.** `template_bank/default_template.pptx`
  is python-pptx's generic built-in template, not a real set of Prezlab
  layouts. Swap in real Prezlab template files here once available.
- **Logo re-insertion isn't implemented.** Stage 1 can detect a logo picture
  on the master; Stage 2's `_apply_logo()` is currently a documented no-op,
  since re-inserting it onto every layout needs image-blob handling not yet
  wired up.
- **Charts are skipped, not recolored.** Stage 3's content-copy step detects
  chart shapes and drops in a placeholder + warning instead of copying them,
  since chart copying needs a separate chart part + embedded workbook —
  that's Phase 3 (`5.3.4 Charts, tables & icons`) scope.
- **No auth.** The API is for `localhost` testing only.
- **Phases 3-4 are entirely out of scope** for this drop (icon libraries,
  chart/table restyling beyond the basic table bonus, and the external
  productization work).

## Git

This folder is a plain local git repo (see `git log`) with no remote
configured — push it to Prezlab's remote of choice when ready.
