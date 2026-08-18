"""
Local dev version of the Submission API (technical plan, Section 4.1).

Deliberately simplified vs. the plan's production architecture:
    - Synchronous request handling, no job queue (Section 4.1 calls for
      Celery/Redis since real formatting is rendering-heavy; fine to skip
      for local testing with small sample files).
    - Local temp-directory storage instead of S3/Azure Blob + Postgres.
    - No auth - this is meant to run on localhost only.

Run locally:
    uvicorn api.main:app --reload --port 8000

Then see README.md for example curl commands, or open
http://127.0.0.1:8000/docs for the interactive Swagger UI.
"""
from __future__ import annotations

import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_formatter.style_spec import StyleSpec  # noqa: E402
from pptx_formatter.extraction import extract_style_spec  # noqa: E402
from pptx_formatter.layout_generator import generate_master_layouts, list_available_layouts, DEFAULT_TEMPLATE  # noqa: E402
from pptx_formatter.pipeline import run_pipeline  # noqa: E402

app = FastAPI(
    title="PPTX Formatting Tool - Local Dev API (Phases 1-2)",
    description="Local, synchronous stand-in for the Submission API in the technical plan. "
                 "Not the production architecture - see README.md.",
    version="0.1.0-phase1-2",
)

WORK_DIR = Path(tempfile.gettempdir()) / "pptx_formatter_dev"
WORK_DIR.mkdir(exist_ok=True)


def _save_upload(upload: UploadFile, suffix: str) -> Path:
    dest = WORK_DIR / f"{uuid.uuid4().hex}_{suffix}"
    with dest.open("wb") as f:
        shutil.copyfileobj(upload.file, f)
    return dest


@app.get("/health")
def health():
    return {"status": "ok", "scope": "Phase 1 (Stage 1 + typography/color) and Phase 2 (Stage 2 + grid)"}


@app.post("/extract")
async def extract(master: UploadFile = File(..., description="Master-slide .pptx")):
    """Stage 1: upload a master slide, get back its Style Spec as JSON."""
    if not master.filename.endswith(".pptx"):
        raise HTTPException(400, "Expected a .pptx file")
    path = _save_upload(master, "master.pptx")
    try:
        spec = extract_style_spec(path)
    except Exception as exc:  # pragma: no cover - surfaced to the caller, not swallowed
        raise HTTPException(422, f"Could not extract a Style Spec: {exc}") from exc
    return JSONResponse(content={"style_spec": _spec_to_dict(spec)})


@app.post("/generate-layouts")
async def generate_layouts(style_spec: str = Form(..., description="Style Spec JSON, as produced by /extract")):
    """Stage 2: submit a Style Spec, get back a restyled template .pptx (all layouts)."""
    import json
    spec = StyleSpec.from_dict(json.loads(style_spec))
    prs = generate_master_layouts(spec, template_path=DEFAULT_TEMPLATE)
    out_path = WORK_DIR / f"{uuid.uuid4().hex}_template.pptx"
    prs.save(str(out_path))
    return FileResponse(
        out_path, filename="restyled_template.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.post("/pipeline")
async def pipeline(
    master: UploadFile = File(..., description="Master-slide .pptx"),
    content: UploadFile = File(..., description="Rough content .pptx"),
    layout: str = Form("Title and Content", description="Layout name to place content slides on"),
):
    """Stages 1-3 end to end: submit a master + a rough content deck, get back the formatted deck."""
    master_path = _save_upload(master, "master.pptx")
    content_path = _save_upload(content, "content.pptx")
    out_path = WORK_DIR / f"{uuid.uuid4().hex}_formatted.pptx"
    try:
        report = run_pipeline(master_path, content_path, out_path, content_layout_name=layout)
    except Exception as exc:
        raise HTTPException(422, f"Pipeline failed: {exc}") from exc
    # Full report (QA flags, warnings, style spec) is available at /pipeline-report
    # if you need it; this endpoint returns the file directly for convenience.
    return FileResponse(
        out_path, filename="formatted_deck.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "X-Slides-Processed": str(report["slides_processed"]),
            "X-QA-Flag-Count": str(len(report["qa_issues"])),
        },
    )


def _spec_to_dict(spec: StyleSpec) -> dict:
    import json
    return json.loads(spec.to_json())
