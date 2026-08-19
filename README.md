# PPTX Formatting Tool

Takes a designer's master slide and a rough content deck, and returns a
formatted deck built on that master. Output is an ordinary `.pptx` that
opens and edits normally in PowerPoint.

## How it works

**1. The designer submits a master.** It can be empty or a fully populated
deck; only the slide master and layouts are read, and any content slides in
it are ignored (the count is reported back).

**2. The master is extracted into a Style Spec** — a versioned JSON
document holding theme colors, per-script fonts, per-layout placeholder
geometry, backgrounds, the logo and its placement rule, footer and
page-number rules, a grid inferred from the layouts, plus chart, table and
icon styling. The spec is archived in the **Template Bank**.

**3. A full layout set is built.** The output deck is built on the
designer's *own* master with its content slides stripped, so their authored
layouts survive natively. Any canonical archetype they didn't author is
filled from the Template Bank — the structurally closest banked layout
wins — or synthesized from their own grid when the bank has nothing
suitable. Every layout is then restyled from the spec, the logo and footer
are propagated, and placeholder inheritance is verified.

**4. The rough deck is reformatted onto it.** Each slide is classified
structurally (cover, two-column, comparison, chart, quote, closing…),
routed to the layout carrying that archetype, and its content mapped into
that layout's placeholders. Arabic text gets right-to-left direction and
the correct complex-script typeface.

Rough-deck decoration does not come with it. A rough slide draws a card as a
filled rectangle with separate text boxes laid on top; once those words are
mapped into a placeholder, the rectangle is an empty shell in coordinates
that mean nothing on the master's grid. Anything that framed harvested text
is dropped, along with whatever rode inside it, and anything still left is
admitted only if it lands clear of the mapped content. Both gates fail
closed and every removal is reported with its reason — see `dropped` in the
per-slide report.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate          # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

## Run the app

```bash
uvicorn api.main:app --reload
```

Open <http://127.0.0.1:8000>. The UI walks the whole flow: name the client,
submit the master, upload the rough deck, review how each slide was read and
correct anything that reads wrong, then format and download. It also browses
the Template Bank and every past run.

<http://127.0.0.1:8000/docs> has the interactive API reference.

## Try it without the UI

Generates a synthetic brand master and a rough deck, then runs everything:

```bash
python examples/run_pipeline.py
```

Output lands in `examples/output/`. It also re-runs the deck as a returning
client to show Stage 1 being skipped.

## Command line

```bash
# Stage 1 - read a master into a Style Spec and archive it
python -m pptx_formatter.cli extract master.pptx spec.json --client "Acme"

# See how a rough deck would be classified, before committing to a reformat
python -m pptx_formatter.cli classify content.pptx

# All three stages
python -m pptx_formatter.cli pipeline master.pptx content.pptx out.pptx \
    --client "Acme" --report report.json

# Disagree with a routing? Override it (slide numbers are 1-based)
python -m pptx_formatter.cli pipeline master.pptx content.pptx out.pptx \
    --override 4=quote --override 9=big_statement

# A repeat deck for a client already in the bank - no master needed
python -m pptx_formatter.cli format-existing "Acme" content.pptx out.pptx

# Inspect the bank
python -m pptx_formatter.cli bank list
python -m pptx_formatter.cli bank show acme
python -m pptx_formatter.cli bank refine acme corrected_spec.json
```

## The Style Spec

The contract between all three stages, and the durable artifact per client.
Plain JSON, tied to no library, so a future consumer (a PowerPoint add-in, a
rendering service) can read it independently.

| Field | Contents |
|---|---|
| `theme.colors` | `dk1`, `lt1`, `dk2`, `lt2`, `accent1`–`6`, `hlink`, `folHlink` as hex |
| `theme.fonts` | Major/minor family per script: Latin, East Asian, complex script |
| `layouts[]` | Per layout: archetype, placeholder geometry, source (designer vs. bank) |
| `brand.logo` | Image reference, position, size, and which layouts it appears on |
| `brand.footer` | Footer text/field rules and page-number behavior |
| `grid` | Margins, gutters, and column/row guide positions inferred from the master |
| `chart_style` | Series color rotation, font, gridline/axis styling |
| `table_style` | Header shading, border weights/colors, cell padding, font |
| `icon_palette` | Accent colors approved for monochrome icon recoloring |
| `meta` | Client/project ids, source master, spec version, extraction timestamp |

Two conventions run through it:

- **Geometry is stored as fractions of the slide, never EMU.** A layout
  banked from a 4:3 master has to be reusable on a 16:9 submission.
- **`None` means "inherits".** A placeholder with no font size inherits from
  the master. Recording that distinction is what lets the inheritance check
  tell a real override from an inherited value.

## The Template Bank

It does two jobs. It **fills gaps** — a submission with no quote layout gets
the closest quote layout any previous submission defined. And it lets a
repeat deck **skip Stage 1** entirely, starting from the archived spec.

The submitted master is archived next to the spec because skipping Stage 1
needs something to build on: the spec describes the brand, but the master
part carries the text styles, color map and theme that a rebuilt deck
inherits.

Corrections feed back in. Refining an entry snapshots the current spec as a
revision before writing the corrected one, so a designer's manual fix during
review is auditable, reversible, and inherited by that client's next deck.

**The bank is client data, not source.** It holds submitted masters and
extracted brand specs. Keep it on storage approved for client material.

## Storage

Two backends behind one interface, chosen by environment. Nothing upstream
changes: the pipeline takes a bank and doesn't care where it lives.

```bash
PPTX_STORAGE_BACKEND=local      # default
PPTX_STORAGE_BACKEND=supabase
```

Copy `.env.example` to `.env` to configure. Both the web app and the CLI
honour the same setting, so `bank list` reads from whichever backend is
active. Passing `--bank <path>` to the CLI always means a local directory.

### Local (default)

```
template_bank/
  index.json                    one row per entry
  entries/<entry_id>/
    style_spec.json             the archived spec
    master.pptx                 the submitted master, archived verbatim
    revisions/rev-<n>.json      prior versions
    assets/                     logo and background images
```

Local stays the default deliberately: the test suite, the demo, and anyone
trying the tool offline shouldn't need a Supabase project to exist.

### Supabase

Postgres is the system of record; a local directory is only a cache. Specs
and job records live in tables, masters and images in Storage buckets.
Anything the pipeline needs as a real file — the master to build on, a logo
to re-insert — is downloaded on demand, because python-pptx opens files, not
streams.

Apply the schema once:

```bash
supabase db execute --file supabase/schema.sql
```

That creates `bank_entries`, `bank_revisions` and `jobs`, plus three private
buckets (`masters`, `assets`, `outputs`), then set:

```bash
PPTX_STORAGE_BACKEND=supabase
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_SERVICE_KEY=...
```

**On access control.** The schema enables row-level security on every table
and grants no policies, so `anon` and `authenticated` can read nothing. The
server reaches the data with the service-role key, which bypasses RLS by
design — so that key stays server-side and is never sent to the browser. The
API itself has no authentication, which means anyone who can reach it can
read every client's banked material. Run it on localhost, or put an
authenticating proxy in front of it and add RLS policies scoped to your auth
model before it goes anywhere else.

## Layout archetypes

The fixed vocabulary shared by the bank, the extractor and the classifier:

`title_slide`, `section_header`, `title_only`, `title_and_content`,
`two_content`, `three_content`, `comparison`, `quote`, `big_statement`,
`picture_full`, `picture_caption`, `table`, `chart`, `blank`, `closing`

A submitted layout is tagged by the designer's own layout name first (a
layout called "Pull Quote" is a quote layout), then by its OOXML type, then
by its placeholder composition.

## Bilingual decks

Arabic is handled per paragraph, not per slide, so a bilingual slide gets
each paragraph on its own terms:

- `<a:pPr rtl="1"/>` so punctuation, digits and embedded Latin order correctly
- alignment flipped to the right, unless the paragraph was deliberately centered
- both the Latin (`+mn-lt`) and complex-script (`+mn-cs`) faces set on every
  run, since PowerPoint resolves the font per character

The complex-script entry is always written into the theme, falling back to
the Latin face when the brand declares none, so `+mn-cs` always resolves to
something deliberate rather than a system substitution.

## API

| Endpoint | Purpose |
|---|---|
| `GET /` | The UI |
| `GET /api/health` | Backend status and the archetype vocabulary |
| `POST /api/extract` | Stage 1: master → Style Spec, archived |
| `POST /api/classify` | Preview how a rough deck would be routed |
| `POST /api/jobs` | Run a formatting job (or `use_banked_spec` to skip Stage 1) |
| `GET /api/jobs` | Past runs |
| `GET /api/jobs/{id}` | One run, with its full report |
| `GET /api/jobs/{id}/download` | Download the formatted deck |
| `GET /api/bank` | Archived entries |
| `GET /api/bank/{id}` | One archived Style Spec |
| `PUT /api/bank/{id}` | Fold a designer's corrections back in |

`POST /api/jobs` returns the full report alongside the download URL:
per-slide classification, the evidence behind it, what was mapped where, and
every QA flag. That's the one to read during review.

The API is synchronous — no job queue — so a request holds open until the
deck is built. Fine for one deck at a time; a production deployment would
want the work queued.

## Tests

```bash
pytest tests/ -q
```

136 tests, no network required. The Supabase backend is exercised against an
in-memory client double (`tests/fake_supabase.py`), including a full pipeline
run and a check that its layout-selection scoring matches the local
backend's exactly.

`tests/test_overlap.py` covers the overlap guarantee end to end: a card-built
rough deck goes through the pipeline and every output slide is checked for
overlapping shapes. Its overlap arithmetic is written out longhand rather
than imported from `pptx_formatter.geometry`, so a bug in that module can't
also be the thing certifying its own output.

They also include a package validator (`tests/conftest.py`) that opens the
output as a zip and checks every part parses, every `rId` reference resolves,
and every part is covered by `[Content_Types].xml`. That catches the failure
mode that matters most here: a file python-pptx reopens happily but
PowerPoint refuses with "needs to be repaired".

## Known limits

- **QA is static, not rendered.** Overflow is estimated from character
  counts scaled by shape width and font size, so it misses some real
  overflows and occasionally flags text that fits. A precise answer needs a
  rendering pass (LibreOffice headless, say) that isn't wired up here.
- **Grid snapping is a snap, not a solver.** It moves unmapped shapes to the
  nearest guide, and it will not create an overlap doing so — a snap that
  would push a shape onto a placeholder or onto an already-snapped shape is
  abandoned and the shape left where it was. It still won't *resolve* an
  overlap the rough deck already had; that's reported, not fixed.
- **The margin inference needs full-width evidence.** Right and bottom
  margins are read off the widest and lowest placeholder, so a master whose
  layouts carry no full-width placeholder gives no evidence for them. A
  right margin inferred at more than twice the left is treated as an
  artefact and mirrored from the left instead of stranding a band of dead
  slide down the side of every generated layout.
- **Placeholders are authoritative.** Content mapped into a placeholder is
  not repositioned or resized — that's what keeps it inheriting from the
  master. Only shapes that couldn't be mapped get snapped.
- **Classification is rule-based.** The rules are inspectable and adjustable
  by whoever owns the deck standards, and every decision reports its
  evidence. There's no labelled corpus of Prezlab slides to learn from, and
  a `--override` exists for when the tool reads a slide wrong.
- **Gradient backgrounds are reproduced as a linear two-stop fill.** More
  complex fills fall back to their first stop.
- **There is no authentication.** Anyone who can reach the app can read
  every client's banked material and download any past deck. Localhost, or
  behind an authenticating proxy, until that changes.
- **Jobs run synchronously.** A request holds open for the length of the
  build, so a large deck ties up a worker. Real concurrency needs a queue.
