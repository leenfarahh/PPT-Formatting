"""
Supabase-backed Template Bank.

Same interface as the local `TemplateBank`, so nothing upstream changes -
`pipeline.py` takes a bank and doesn't care where it stores things.

The split is: **Postgres is the system of record, the local directory is a
cache.** Specs and index rows live in tables; masters and image assets live
in Storage buckets. Anything the pipeline needs as a real file on disk - the
master to build on, a logo to re-insert - is downloaded into a cache
directory on demand, which is what lets `master_path()` and `asset_dir()`
keep returning `Path` objects.

Selection scoring and client lookup come from `BaseBank`, so they can't
drift from the local backend's behavior.
"""
from __future__ import annotations

import json
from pathlib import Path

from .bank import (
    BankEntry, BaseBank, _absolutize_assets, _asset_holders, _now,
    _relativize_assets, _slug,
)
from .config import BUCKET_ASSETS, BUCKET_MASTERS
from .style_spec import StyleSpec

TABLE_ENTRIES = "bank_entries"
TABLE_REVISIONS = "bank_revisions"


class SupabaseBank(BaseBank):
    """Template Bank stored in Supabase Postgres plus Storage."""

    def __init__(self, client, cache_dir: str | Path, bucket_masters: str = BUCKET_MASTERS,
                 bucket_assets: str = BUCKET_ASSETS):
        self.client = client
        self.cache_dir = Path(cache_dir)
        self.bucket_masters = bucket_masters
        self.bucket_assets = bucket_assets

    # -- local cache -------------------------------------------------------

    def entry_dir(self, entry_id: str) -> Path:
        return self.cache_dir / entry_id

    def asset_dir(self, entry_id: str) -> Path:
        path = self.entry_dir(entry_id) / "assets"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def master_path(self, entry_id: str) -> Path:
        """
        Local path to the archived master, downloading it if not cached.

        Stage 2 opens this file to build on, so it has to be a real file
        rather than a stream.
        """
        path = self.entry_dir(entry_id) / "master.pptx"
        if path.exists():
            return path
        row = self._row(entry_id)
        object_name = (row or {}).get("master_object")
        if not object_name:
            return path         # caller checks has_master() first
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self._download(self.bucket_masters, object_name))
        return path

    def has_master(self, entry_id: str) -> bool:
        row = self._row(entry_id)
        return bool(row and row.get("master_object"))

    # -- storage helpers ---------------------------------------------------

    def _upload(self, bucket: str, object_name: str, data: bytes, content_type: str) -> None:
        """
        Upload, replacing anything already at that path.

        Re-running an extraction for the same entry should overwrite its
        logo rather than fail on a name collision, so upsert is always on.
        """
        self.client.storage.from_(bucket).upload(
            object_name,
            data,
            {"content-type": content_type, "upsert": "true"},
        )

    def _download(self, bucket: str, object_name: str) -> bytes:
        return self.client.storage.from_(bucket).download(object_name)

    # -- rows --------------------------------------------------------------

    def _row(self, entry_id: str) -> dict | None:
        response = (
            self.client.table(TABLE_ENTRIES)
            .select("*")
            .eq("entry_id", entry_id)
            .limit(1)
            .execute()
        )
        rows = response.data or []
        return rows[0] if rows else None

    def _rows(self) -> list:
        response = (
            self.client.table(TABLE_ENTRIES)
            .select("*")
            .order("updated_at", desc=False)
            .execute()
        )
        return response.data or []

    @staticmethod
    def _to_entry(row: dict) -> BankEntry:
        return BankEntry(
            entry_id=row["entry_id"],
            client=row.get("client"),
            project=row.get("project"),
            source_name=row.get("source_name"),
            spec_version=row.get("spec_version") or "1.0",
            revision=row.get("revision") or 1,
            created_at=row.get("created_at"),
            updated_at=row.get("updated_at"),
            slide_width=row.get("slide_width") or 0,
            slide_height=row.get("slide_height") or 0,
            archetypes=list(row.get("archetypes") or []),
        )

    def list_entries(self) -> list:
        return [self._to_entry(row) for row in self._rows()]

    def entry(self, entry_id: str) -> BankEntry | None:
        row = self._row(entry_id)
        return self._to_entry(row) if row else None

    # -- writing -----------------------------------------------------------

    def allocate(self, client: str | None, source_name: str | None) -> str:
        """Reserve an entry id, so extraction can write assets under it."""
        stem = _slug(client or source_name or "master")
        taken = {row["entry_id"] for row in self._rows()}
        entry_id = stem
        n = 2
        while entry_id in taken:
            entry_id = f"{stem}-{n}"
            n += 1
        self.asset_dir(entry_id)
        return entry_id

    def save(
        self,
        spec: StyleSpec,
        entry_id: str | None = None,
        master_pptx: str | Path | None = None,
    ) -> str:
        """Archive a spec, uploading its master and assets to Storage."""
        if entry_id is None:
            entry_id = self.allocate(spec.meta.client, spec.meta.source_name)

        entry_dir = self.entry_dir(entry_id)
        entry_dir.mkdir(parents=True, exist_ok=True)

        master_object = None
        if master_pptx is not None:
            source = Path(master_pptx)
            if source.exists():
                master_object = f"{entry_id}/master.pptx"
                self._upload(
                    self.bucket_masters, master_object, source.read_bytes(),
                    "application/vnd.openxmlformats-officedocument.presentationml"
                    ".presentation",
                )
                # Keep a cached copy so an immediate Stage 2 doesn't round-trip.
                cached = entry_dir / "master.pptx"
                if source.resolve() != cached.resolve():
                    cached.write_bytes(source.read_bytes())

        # Normalize asset paths to be entry-relative, then push the bytes up.
        _relativize_assets(spec, entry_dir)
        self._upload_assets(spec, entry_id, entry_dir)

        existing = self._row(entry_id)
        record = {
            "entry_id": entry_id,
            "client": spec.meta.client or (existing or {}).get("client"),
            "project": spec.meta.project or (existing or {}).get("project"),
            "source_name": spec.meta.source_name,
            "spec_version": spec.meta.spec_version,
            "revision": (existing or {}).get("revision") or 1,
            "slide_width": spec.slide_width,
            "slide_height": spec.slide_height,
            "archetypes": sorted(spec.archetypes_present()),
            "style_spec": json.loads(spec.to_json()),
            "updated_at": _now(),
        }
        if master_object:
            record["master_object"] = master_object
        elif existing and existing.get("master_object"):
            record["master_object"] = existing["master_object"]
        if not existing:
            record["created_at"] = _now()

        self.client.table(TABLE_ENTRIES).upsert(record, on_conflict="entry_id").execute()
        return entry_id

    def _upload_assets(self, spec: StyleSpec, entry_id: str, entry_dir: Path) -> None:
        for holder in _asset_holders(spec):
            relative = getattr(holder, "asset_path", None)
            if not relative:
                continue
            local = entry_dir / relative
            if not local.exists():
                continue
            self._upload(
                self.bucket_assets, f"{entry_id}/{relative}",
                local.read_bytes(), _content_type(local),
            )

    def refine(self, entry_id: str, spec: StyleSpec) -> int:
        """
        Replace an entry's spec, keeping the previous revision.

        The old spec goes into `bank_revisions` before the new one lands, so
        a designer's correction stays auditable and reversible - the same
        guarantee the local backend gives with its `revisions/` directory.
        """
        existing = self._row(entry_id)
        if not existing:
            raise FileNotFoundError(f"No banked spec for entry '{entry_id}'")

        revision = (existing.get("revision") or 1) + 1
        self.client.table(TABLE_REVISIONS).insert({
            "entry_id": entry_id,
            "revision": revision - 1,
            "style_spec": existing["style_spec"],
            "created_at": _now(),
        }).execute()

        entry_dir = self.entry_dir(entry_id)
        entry_dir.mkdir(parents=True, exist_ok=True)
        _relativize_assets(spec, entry_dir)
        self._upload_assets(spec, entry_id, entry_dir)
        spec.meta.notes.append(f"revision {revision} saved {_now()}")

        self.client.table(TABLE_ENTRIES).update({
            "style_spec": json.loads(spec.to_json()),
            "revision": revision,
            "archetypes": sorted(spec.archetypes_present()),
            "updated_at": _now(),
        }).eq("entry_id", entry_id).execute()
        return revision

    # -- reading -----------------------------------------------------------

    def load(self, entry_id: str) -> StyleSpec:
        """Load a spec, caching any assets it references to local files."""
        row = self._row(entry_id)
        if not row:
            raise FileNotFoundError(f"No banked spec for entry '{entry_id}'")

        spec = StyleSpec.from_dict(row["style_spec"])
        entry_dir = self.entry_dir(entry_id)
        self._cache_assets(spec, entry_id, entry_dir)
        _absolutize_assets(spec, entry_dir)
        return spec

    def _cache_assets(self, spec: StyleSpec, entry_id: str, entry_dir: Path) -> None:
        """
        Materialize referenced assets on disk.

        Stage 2 inserts the logo by path, so the bytes have to exist locally.
        A missing object is cleared from the spec rather than raised: a
        deleted logo shouldn't stop a deck being formatted.
        """
        for holder in _asset_holders(spec):
            relative = getattr(holder, "asset_path", None)
            if not relative or Path(relative).is_absolute():
                continue
            local = entry_dir / relative
            if local.exists():
                continue
            local.parent.mkdir(parents=True, exist_ok=True)
            try:
                local.write_bytes(
                    self._download(self.bucket_assets, f"{entry_id}/{relative}")
                )
            except Exception:               # noqa: BLE001 - client-specific errors
                holder.asset_path = None

    def revisions(self, entry_id: str) -> list:
        """Prior revisions of an entry, oldest first."""
        response = (
            self.client.table(TABLE_REVISIONS)
            .select("revision, created_at")
            .eq("entry_id", entry_id)
            .order("revision", desc=False)
            .execute()
        )
        return response.data or []


def _content_type(path: Path) -> str:
    return {
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".gif": "image/gif", ".svg": "image/svg+xml", ".webp": "image/webp",
        ".emf": "image/x-emf", ".wmf": "image/x-wmf",
    }.get(path.suffix.lower(), "application/octet-stream")
