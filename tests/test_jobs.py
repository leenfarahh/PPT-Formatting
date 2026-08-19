"""Job records, against both storage backends."""
from __future__ import annotations

import pytest

from pptx_formatter.jobs import (
    STATUS_COMPLETE, STATUS_FAILED, LocalJobStore, SupabaseJobStore,
)

from fake_supabase import FakeSupabaseClient

REPORT = {
    "slides_processed": 6,
    "qa_issues": ["slide 2: something"],
    "warnings": ["slide 3: something else", "slide 4: another"],
    "bank_entry": "acme",
}


@pytest.fixture(params=["local", "supabase"])
def store(request, tmp_path):
    """Both backends, so parity is checked rather than assumed."""
    if request.param == "local":
        return LocalJobStore(tmp_path / "jobs")
    return SupabaseJobStore(FakeSupabaseClient(), cache_dir=tmp_path / "cache")


@pytest.fixture
def deck(tmp_path):
    path = tmp_path / "formatted.pptx"
    path.write_bytes(b"PK\x03\x04 pretend deck")
    return path


def test_create_starts_a_running_job(store):
    record = store.create(client="Acme", content_filename="rough.pptx")
    assert record.status == "running"
    assert record.created_at is not None
    assert store.get(record.job_id).client == "Acme"


def test_complete_records_counts_from_the_report(store, deck):
    record = store.create(client="Acme", content_filename="rough.pptx")
    store.complete(record.job_id, deck, REPORT)

    stored = store.get(record.job_id)
    assert stored.status == STATUS_COMPLETE
    assert stored.slides_processed == 6
    assert stored.qa_flag_count == 1
    assert stored.warning_count == 2
    assert stored.entry_id == "acme"
    assert stored.completed_at is not None


def test_output_is_retrievable_after_completion(store, deck):
    record = store.create(client="Acme", content_filename="rough.pptx")
    store.complete(record.job_id, deck, REPORT)

    path = store.output_path(record.job_id)
    assert path is not None and path.exists()
    assert path.read_bytes() == deck.read_bytes()


def test_a_running_job_has_nothing_to_download(store):
    record = store.create(client="Acme", content_filename="rough.pptx")
    assert store.output_path(record.job_id) is None


def test_failure_is_recorded_with_its_reason(store):
    record = store.create(client="Acme", content_filename="rough.pptx")
    store.fail(record.job_id, "master was not a .pptx")

    stored = store.get(record.job_id)
    assert stored.status == STATUS_FAILED
    assert "not a .pptx" in stored.error


def test_listing_is_newest_first(store, deck):
    for i in range(3):
        record = store.create(client=f"Client {i}", content_filename=f"{i}.pptx")
        store.complete(record.job_id, deck, REPORT)

    listed = store.list()
    assert len(listed) == 3
    timestamps = [job.created_at for job in listed]
    assert timestamps == sorted(timestamps, reverse=True)


def test_summary_omits_the_full_report(store, deck):
    """The list view shouldn't carry every slide's evidence."""
    record = store.create(client="Acme", content_filename="rough.pptx")
    store.complete(record.job_id, deck, REPORT)

    summary = store.get(record.job_id).summary()
    assert "report" not in summary
    assert summary["slides_processed"] == 6


def test_unknown_job(store, deck):
    assert store.get("nope") is None
    assert store.output_path("nope") is None
    with pytest.raises(LookupError):
        store.fail("nope", "boom")
    with pytest.raises(LookupError):
        store.complete("nope", deck, REPORT)


def test_supabase_job_uploads_the_deck_to_storage(tmp_path, deck):
    client = FakeSupabaseClient()
    store = SupabaseJobStore(client, cache_dir=tmp_path / "cache")

    record = store.create(client="Acme", content_filename="rough.pptx")
    store.complete(record.job_id, deck, REPORT)

    assert f"outputs/{record.job_id}/formatted_deck.pptx" in client.object_paths()
    assert client.rows("jobs")[0]["report"]["slides_processed"] == 6


def test_supabase_download_is_cached_locally(tmp_path, deck):
    client = FakeSupabaseClient()
    store = SupabaseJobStore(client, cache_dir=tmp_path / "cache")
    record = store.create(client="Acme", content_filename="rough.pptx")
    store.complete(record.job_id, deck, REPORT)

    first = store.output_path(record.job_id)
    assert first.exists()
    # Second call must not need the object store again.
    client.storage.objects.clear()
    assert store.output_path(record.job_id) == first
