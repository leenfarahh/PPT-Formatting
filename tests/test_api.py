"""The HTTP API and the UI it serves."""
from __future__ import annotations

import json
import zipfile

import pytest
from fastapi.testclient import TestClient

from api.main import app


@pytest.fixture
def client(tmp_path, monkeypatch):
    """A client wired to a throwaway local bank, so tests never share state."""
    monkeypatch.setenv("PPTX_STORAGE_BACKEND", "local")
    monkeypatch.setenv("PPTX_BANK_ROOT", str(tmp_path / "bank"))
    return TestClient(app)


def upload(path, field="content"):
    ctype = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    return {field: (path.name, path.read_bytes(), ctype)}


# --- meta -----------------------------------------------------------------

def test_health_reports_the_backend(client):
    body = client.get("/api/health").json()
    assert body["status"] == "ok"
    assert body["storage"]["backend"] == "local"
    values = [a["value"] for a in body["archetypes"]]
    assert "quote" in values and "title_and_content" in values


def test_health_never_leaks_key_material(client, monkeypatch):
    monkeypatch.setenv("PPTX_STORAGE_BACKEND", "supabase")
    monkeypatch.setenv("SUPABASE_URL", "https://example.supabase.co")
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "super-secret-service-key")

    body = client.get("/api/health").json()
    assert body["storage"]["configured"] is True
    assert "super-secret-service-key" not in json.dumps(body)


def test_ui_is_served(client):
    response = client.get("/")
    assert response.status_code == 200
    assert "PPTX Formatting Tool" in response.text
    # The page has to stand alone: no CDN, no build step.
    assert "<script" in response.text
    assert "https://" not in response.text.lower()


def test_ui_references_resolve():
    """
    A typo in an element id or an endpoint path fails silently in a browser -
    no error, just a control that does nothing. Cheap to check statically.
    """
    import re
    from pathlib import Path

    root = Path(__file__).resolve().parent.parent
    html = (root / "api" / "static" / "index.html").read_text(encoding="utf-8")
    routes = (root / "api" / "main.py").read_text(encoding="utf-8")

    ids = set(re.findall(r'\bid="([^"]+)"', html))
    referenced = set(re.findall(r'\$\("([^"]+)"\)', html))
    assert not referenced - ids, f"script refers to unknown ids: {referenced - ids}"

    for tab in set(re.findall(r'data-tab="([^"]+)"', html)):
        assert f'id="tab-{tab}"' in html, f"nav tab {tab!r} has no section"

    for path in set(re.findall(r'"(/api/[a-z/-]+)', html)):
        assert path.rstrip("/") in routes, f"UI calls undefined endpoint {path}"


# --- stage 1 --------------------------------------------------------------

def test_extract_returns_a_summary_and_the_full_spec(client, master_path):
    response = client.post(
        "/api/extract", files=upload(master_path, "master"),
        data={"client": "Acme Holdings", "project": "Board Deck"},
    )
    assert response.status_code == 200
    body = response.json()

    assert body["bank_entry"] == "acme-holdings"
    summary = body["summary"]
    assert summary["content_slides_ignored"] == 2
    assert summary["colors"]["accent1"] == "0F4C81"
    assert summary["fonts"]["minor_cs"] == "Dubai"
    assert "quote" in summary["archetypes_missing"]
    assert body["style_spec"]["meta"]["client"] == "Acme Holdings"


def test_non_pptx_upload_is_rejected(client, tmp_path):
    junk = tmp_path / "notes.txt"
    junk.write_bytes(b"not a deck")
    response = client.post(
        "/api/extract", files={"master": ("notes.txt", junk.read_bytes(), "text/plain")}
    )
    assert response.status_code == 400
    assert ".pptx" in response.json()["detail"]


# --- classification -------------------------------------------------------

def test_classify_previews_routing_with_evidence(client, content_path):
    body = client.post("/api/classify", files=upload(content_path)).json()
    slides = body["slides"]

    assert len(slides) == 6
    assert slides[0]["archetype"] == "title_slide"
    assert slides[5]["archetype"] == "closing"
    assert all(s["evidence"] for s in slides)
    assert all(0 <= s["confidence"] <= 1 for s in slides)


# --- jobs -----------------------------------------------------------------

def test_job_runs_the_pipeline_and_returns_a_download(client, master_path, content_path):
    response = client.post(
        "/api/jobs",
        files={**upload(content_path), **upload(master_path, "master")},
        data={"client": "Acme Holdings"},
    )
    assert response.status_code == 200
    body = response.json()

    assert body["job"]["status"] == "complete"
    assert body["report"]["slides_processed"] == 6
    assert body["download_url"].endswith("/download")

    download = client.get(body["download_url"])
    assert download.status_code == 200
    assert download.headers["content-type"].startswith(
        "application/vnd.openxmlformats"
    )
    # A .pptx is a zip; a truncated or empty response would not be.
    assert download.content[:2] == b"PK"


def test_overrides_are_honored(client, master_path, content_path):
    body = client.post(
        "/api/jobs",
        files={**upload(content_path), **upload(master_path, "master")},
        data={"client": "Acme", "overrides": json.dumps({"3": "section_header"})},
    ).json()

    slide = body["report"]["slides"][2]
    assert slide["archetype"] == "section_header"
    assert any("overridden" in line for line in slide["evidence"])


def test_unknown_archetype_in_overrides_is_rejected(client, master_path, content_path):
    response = client.post(
        "/api/jobs",
        files={**upload(content_path), **upload(master_path, "master")},
        data={"overrides": json.dumps({"1": "not_a_real_layout"})},
    )
    assert response.status_code == 400
    assert "unknown archetype" in response.json()["detail"]


def test_job_without_a_master_or_banked_brand_is_rejected(client, content_path):
    response = client.post("/api/jobs", files=upload(content_path))
    assert response.status_code == 400
    assert "master" in response.json()["detail"].lower()


def test_repeat_client_uses_the_banked_brand(client, master_path, content_path):
    client.post("/api/extract", files=upload(master_path, "master"),
                data={"client": "Acme Holdings"})

    body = client.post(
        "/api/jobs", files=upload(content_path),
        data={"client": "Acme Holdings", "use_banked_spec": "true"},
    ).json()

    assert body["report"]["stage_1_skipped"] is True
    assert body["job"]["slides_processed"] == 6


def test_unknown_client_for_a_banked_run_is_a_404(client, content_path):
    response = client.post(
        "/api/jobs", files=upload(content_path),
        data={"client": "Nobody At All", "use_banked_spec": "true"},
    )
    assert response.status_code == 404


def test_failed_jobs_are_recorded_rather_than_lost(client, content_path):
    """A failure should still leave a job row explaining itself."""
    client.post(
        "/api/jobs", files=upload(content_path),
        data={"client": "Ghost Client", "use_banked_spec": "true"},
    )
    jobs = client.get("/api/jobs").json()["jobs"]
    assert jobs and jobs[0]["status"] == "failed"
    assert jobs[0]["error"]


def test_jobs_are_listed_and_individually_retrievable(client, master_path, content_path):
    created = client.post(
        "/api/jobs",
        files={**upload(content_path), **upload(master_path, "master")},
        data={"client": "Acme"},
    ).json()["job"]

    listed = client.get("/api/jobs").json()["jobs"]
    assert any(job["job_id"] == created["job_id"] for job in listed)

    detail = client.get(f"/api/jobs/{created['job_id']}").json()
    assert detail["report"]["slides_processed"] == 6
    # The list view stays light; the detail view carries the report.
    assert "report" not in listed[0]


def test_download_of_an_unknown_job_is_a_404(client):
    assert client.get("/api/jobs/nope/download").status_code == 404


# --- bank -----------------------------------------------------------------

def test_bank_lists_and_shows_entries(client, master_path):
    client.post("/api/extract", files=upload(master_path, "master"),
                data={"client": "Acme Holdings"})

    entries = client.get("/api/bank").json()["entries"]
    assert len(entries) == 1
    assert entries[0]["client"] == "Acme Holdings"
    assert entries[0]["has_master"] is True

    detail = client.get(f"/api/bank/{entries[0]['entry_id']}").json()
    assert detail["summary"]["colors"]["accent1"] == "0F4C81"
    assert detail["style_spec"]["meta"]["spec_version"] == "1.0"


def test_refining_an_entry_keeps_a_revision(client, master_path):
    entry_id = client.post(
        "/api/extract", files=upload(master_path, "master"),
        data={"client": "Acme Holdings"},
    ).json()["bank_entry"]

    spec = client.get(f"/api/bank/{entry_id}").json()["style_spec"]
    spec["theme"]["colors"]["accent1"] = "FF0000"

    response = client.put(f"/api/bank/{entry_id}", data={"style_spec": json.dumps(spec)})
    assert response.status_code == 200
    assert response.json()["revision"] == 2

    updated = client.get(f"/api/bank/{entry_id}").json()
    assert updated["summary"]["colors"]["accent1"] == "FF0000"


def test_unknown_bank_entry_is_a_404(client):
    assert client.get("/api/bank/nope").status_code == 404


# --- configuration --------------------------------------------------------

def test_misconfigured_supabase_reads_as_a_setup_problem(client, monkeypatch):
    """Not a 500: the deployment is incomplete, and the message should say so."""
    monkeypatch.setenv("PPTX_STORAGE_BACKEND", "supabase")
    monkeypatch.delenv("SUPABASE_URL", raising=False)
    monkeypatch.delenv("SUPABASE_SERVICE_KEY", raising=False)
    monkeypatch.delenv("SUPABASE_ANON_KEY", raising=False)

    response = client.get("/api/bank")
    assert response.status_code == 503
    assert "SUPABASE_URL" in response.json()["detail"]
