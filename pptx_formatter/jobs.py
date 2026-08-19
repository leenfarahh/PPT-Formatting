"""
Job records - one per formatting run.

A run produces a deck plus a report worth keeping: which layout each slide
was routed to, why, and what QA flagged. Recording that makes the output
re-downloadable, makes a disputed routing reviewable after the fact, and
gives the UI something to list.

Two backends behind one interface, matching the Template Bank: local JSON
files, or a Postgres table plus a Storage bucket.
"""
from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from .config import BUCKET_OUTPUTS

TABLE_JOBS = "jobs"

STATUS_RUNNING = "running"
STATUS_COMPLETE = "complete"
STATUS_FAILED = "failed"

PPTX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


@dataclass
class JobRecord:
    job_id: str
    client: str | None = None
    project: str | None = None
    entry_id: str | None = None
    master_filename: str | None = None
    content_filename: str | None = None
    output_name: str = "formatted_deck.pptx"
    status: str = STATUS_RUNNING
    error: str | None = None
    slides_processed: int = 0
    qa_flag_count: int = 0
    warning_count: int = 0
    stage_1_skipped: bool = False
    report: dict = field(default_factory=dict)
    created_at: str | None = None
    completed_at: str | None = None

    def summary(self) -> dict:
        """The row shape the UI lists, without the full report payload."""
        data = asdict(self)
        data.pop("report", None)
        return data


class BaseJobStore:
    """Shared helpers; subclasses provide the storage."""

    def new_id(self) -> str:
        return uuid.uuid4().hex

    def create(self, **kwargs) -> JobRecord:
        raise NotImplementedError

    def complete(self, job_id: str, output_file: Path, report: dict) -> JobRecord:
        raise NotImplementedError

    def fail(self, job_id: str, error: str) -> JobRecord:
        raise NotImplementedError

    def get(self, job_id: str) -> JobRecord | None:
        raise NotImplementedError

    def list(self, limit: int = 50) -> list:
        raise NotImplementedError

    def output_path(self, job_id: str) -> Path | None:
        raise NotImplementedError


# --- local ----------------------------------------------------------------

class LocalJobStore(BaseJobStore):
    """Jobs as JSON files, output decks alongside them."""

    def __init__(self, root: str | Path):
        self.root = Path(root)

    def _dir(self, job_id: str) -> Path:
        return self.root / job_id

    def _record_path(self, job_id: str) -> Path:
        return self._dir(job_id) / "job.json"

    def _write(self, record: JobRecord) -> JobRecord:
        path = self._record_path(record.job_id)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(asdict(record), indent=2, ensure_ascii=False), encoding="utf-8"
        )
        return record

    def create(self, **kwargs) -> JobRecord:
        record = JobRecord(job_id=self.new_id(), created_at=_now(), **kwargs)
        return self._write(record)

    def complete(self, job_id: str, output_file: Path, report: dict) -> JobRecord:
        record = self.get(job_id)
        if record is None:
            raise LookupError(f"No such job '{job_id}'")

        destination = self._dir(job_id) / record.output_name
        destination.parent.mkdir(parents=True, exist_ok=True)
        source = Path(output_file)
        if source.resolve() != destination.resolve():
            destination.write_bytes(source.read_bytes())

        record.status = STATUS_COMPLETE
        record.completed_at = _now()
        record.report = report
        _apply_report_counts(record, report)
        return self._write(record)

    def fail(self, job_id: str, error: str) -> JobRecord:
        record = self.get(job_id)
        if record is None:
            raise LookupError(f"No such job '{job_id}'")
        record.status = STATUS_FAILED
        record.error = error
        record.completed_at = _now()
        return self._write(record)

    def get(self, job_id: str) -> JobRecord | None:
        path = self._record_path(job_id)
        if not path.exists():
            return None
        data = json.loads(path.read_text(encoding="utf-8"))
        known = set(JobRecord.__dataclass_fields__)
        return JobRecord(**{k: v for k, v in data.items() if k in known})

    def list(self, limit: int = 50) -> list:
        if not self.root.exists():
            return []
        records = [
            record
            for path in self.root.iterdir() if path.is_dir()
            for record in [self.get(path.name)] if record is not None
        ]
        records.sort(key=lambda r: r.created_at or "", reverse=True)
        return records[:limit]

    def output_path(self, job_id: str) -> Path | None:
        record = self.get(job_id)
        if record is None:
            return None
        path = self._dir(job_id) / record.output_name
        return path if path.exists() else None


# --- supabase -------------------------------------------------------------

class SupabaseJobStore(BaseJobStore):
    """Jobs as Postgres rows, output decks in a Storage bucket."""

    def __init__(self, client, cache_dir: str | Path, bucket: str = BUCKET_OUTPUTS):
        self.client = client
        self.cache_dir = Path(cache_dir) / "jobs"
        self.bucket = bucket

    def _row(self, job_id: str) -> dict | None:
        response = (
            self.client.table(TABLE_JOBS)
            .select("*").eq("job_id", job_id).limit(1).execute()
        )
        rows = response.data or []
        return rows[0] if rows else None

    @staticmethod
    def _to_record(row: dict) -> JobRecord:
        known = set(JobRecord.__dataclass_fields__)
        data = {k: v for k, v in row.items() if k in known}
        data["report"] = row.get("report") or {}
        return JobRecord(**data)

    def create(self, **kwargs) -> JobRecord:
        record = JobRecord(job_id=self.new_id(), created_at=_now(), **kwargs)
        payload = asdict(record)
        payload.pop("report")
        self.client.table(TABLE_JOBS).insert(payload).execute()
        return record

    def complete(self, job_id: str, output_file: Path, report: dict) -> JobRecord:
        row = self._row(job_id)
        if row is None:
            raise LookupError(f"No such job '{job_id}'")
        record = self._to_record(row)

        object_name = f"{job_id}/{record.output_name}"
        self.client.storage.from_(self.bucket).upload(
            object_name,
            Path(output_file).read_bytes(),
            {"content-type": PPTX_CONTENT_TYPE, "upsert": "true"},
        )

        record.status = STATUS_COMPLETE
        record.completed_at = _now()
        record.report = report
        _apply_report_counts(record, report)

        self.client.table(TABLE_JOBS).update({
            "status": record.status,
            "completed_at": record.completed_at,
            "slides_processed": record.slides_processed,
            "qa_flag_count": record.qa_flag_count,
            "warning_count": record.warning_count,
            "entry_id": record.entry_id,
            "output_object": object_name,
            "report": report,
        }).eq("job_id", job_id).execute()
        return record

    def fail(self, job_id: str, error: str) -> JobRecord:
        row = self._row(job_id)
        if row is None:
            raise LookupError(f"No such job '{job_id}'")
        record = self._to_record(row)
        record.status = STATUS_FAILED
        record.error = error
        record.completed_at = _now()
        self.client.table(TABLE_JOBS).update({
            "status": record.status, "error": error, "completed_at": record.completed_at,
        }).eq("job_id", job_id).execute()
        return record

    def get(self, job_id: str) -> JobRecord | None:
        row = self._row(job_id)
        return self._to_record(row) if row else None

    def list(self, limit: int = 50) -> list:
        response = (
            self.client.table(TABLE_JOBS)
            .select("*").order("created_at", desc=True).limit(limit).execute()
        )
        return [self._to_record(row) for row in (response.data or [])]

    def output_path(self, job_id: str) -> Path | None:
        """Download the output to a local file so it can be served."""
        row = self._row(job_id)
        if not row or not row.get("output_object"):
            return None
        record = self._to_record(row)
        local = self.cache_dir / job_id / record.output_name
        if local.exists():
            return local
        local.parent.mkdir(parents=True, exist_ok=True)
        local.write_bytes(
            self.client.storage.from_(self.bucket).download(row["output_object"])
        )
        return local


def _apply_report_counts(record: JobRecord, report: dict) -> None:
    record.slides_processed = report.get("slides_processed", 0)
    record.qa_flag_count = len(report.get("qa_issues", []))
    record.warning_count = len(report.get("warnings", []))
    record.stage_1_skipped = bool(report.get("stage_1_skipped"))
    if report.get("bank_entry"):
        record.entry_id = report["bank_entry"]


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
