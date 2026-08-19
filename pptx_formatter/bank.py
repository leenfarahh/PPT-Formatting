"""
The Template Bank - a durable, versioned archive of Style Specs.

Every extraction deposits its spec here, so the bank accumulates canonical
layout archetypes drawn from real Prezlab masters rather than from a
hand-authored set of generic templates. That gives it two jobs:

1.  **Gap filling.** When a submitted master doesn't define, say, a quote
    layout, the bank supplies the structurally closest quote layout any
    previous submission did define. Geometry is stored as slide fractions,
    so a layout banked from a 4:3 deck still applies cleanly to a 16:9 one.

2.  **Skipping Stage 1.** A repeat deck for the same client can start from
    the archived spec, no master re-submission needed. `latest_for_client()`
    is that entry point.

Layout on disk:

    template_bank/
      index.json                       one row per entry, newest last
      entries/<entry_id>/
        style_spec.json                the archived Style Spec
        master.pptx                    the submitted master, archived verbatim
        revisions/rev-<n>.json         prior versions, kept on refine()
        assets/                        logo and background images

The master file is archived next to the spec because "skip Stage 1" needs
something to build on: the spec describes the brand, but the master part
itself carries the text styles, color map and theme part that a rebuilt
deck inherits from. Keeping both means a repeat deck reproduces the
original exactly rather than approximately.

Entries are never overwritten in place: `refine()` snapshots the current
spec into `revisions/` before writing the new one, so a designer's manual
corrections are auditable and reversible.
"""
from __future__ import annotations

import json
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from . import archetypes
from .style_spec import StyleSpec, LayoutSpec

DEFAULT_BANK_ROOT = Path(__file__).resolve().parent.parent / "template_bank"
INDEX_NAME = "index.json"


# --- scoring weights ------------------------------------------------------
# An exact archetype match should nearly always win, so it is worth more
# than any structural or aspect penalty can plausibly recover.
EXACT_ARCHETYPE_BONUS = 100.0
FALLBACK_STEP_PENALTY = 12.0     # per position down the fallback chain
ASPECT_PENALTY = 15.0            # per unit of aspect-ratio difference
SAME_CLIENT_BONUS = 8.0
STRUCTURE_WEIGHT = 3.0

# Below this, a banked layout is too far from what was asked for to be worth
# using: a purpose-built layout generated from the submission's own grid
# will fit better than a distant cousin borrowed from another deck.
MIN_SELECTION_SCORE = 60.0


@dataclass
class BankEntry:
    """One archived master, as recorded in the index."""
    entry_id: str
    client: str | None = None
    project: str | None = None
    source_name: str | None = None
    spec_version: str = "1.0"
    revision: int = 1
    created_at: str | None = None
    updated_at: str | None = None
    slide_width: int = 0
    slide_height: int = 0
    archetypes: list = field(default_factory=list)

    @property
    def aspect_ratio(self) -> float:
        return self.slide_width / self.slide_height if self.slide_height else 1.7778


class BaseBank:
    """
    Storage-agnostic bank behavior.

    Everything here is expressed in terms of `list_entries()` and `load()`,
    so the layout-selection scoring - the part with actual judgment in it -
    is written once and shared by the local and Supabase backends rather
    than drifting between them.

    Subclasses supply the storage: `list_entries`, `entry`, `load`, `save`,
    `refine`, `allocate`, `asset_dir`, `master_path` and `has_master`.
    """

    # -- to be provided by the backend ------------------------------------

    def list_entries(self) -> list:
        raise NotImplementedError

    def load(self, entry_id: str) -> StyleSpec:
        raise NotImplementedError

    # -- shared behavior ---------------------------------------------------

    def latest_entry_for_client(self, client: str) -> BankEntry | None:
        """
        The most recently updated entry for a client, matched on a slug so
        "Acme Bank" and "acme-bank" resolve to the same account.
        """
        target = _slug(client)
        if not target:
            return None
        matches = [e for e in self.list_entries() if _slug(e.client or "") == target]
        if not matches:
            return None
        return max(matches, key=lambda e: e.updated_at or e.created_at or "")

    def latest_for_client(self, client: str) -> StyleSpec | None:
        """
        The most recently archived spec for a client, or None.

        This is what lets a repeat deck skip Stage 1 entirely.
        """
        entry = self.latest_entry_for_client(client)
        return self.load(entry.entry_id) if entry else None

    def select_layout(
        self,
        archetype: str,
        target: StyleSpec,
        exclude_entry: str | None = None,
        min_score: float = MIN_SELECTION_SCORE,
    ) -> tuple[LayoutSpec, float] | None:
        """
        Find the banked layout that best fills a gap for `archetype`.

        Scoring, highest wins:
          + an exact archetype match dominates everything else
          + partial credit for archetypes near it in the fallback chain
          - structural distance from what this archetype should look like
          - difference in slide aspect ratio between bank entry and target
          + a nudge toward the same client's own house style

        Returns the chosen layout (tagged `bank:<entry_id>`) and its score,
        or None when the bank holds nothing scoring above `min_score` -
        in which case the caller should generate a layout instead.
        """
        wanted = archetypes.canonical_signature(archetype)
        chain = archetypes.FALLBACK_CHAIN.get(archetype, [])
        target_client = _slug(target.meta.client or "")

        best: tuple[LayoutSpec, float] | None = None
        for entry in self.list_entries():
            if exclude_entry and entry.entry_id == exclude_entry:
                continue
            # Cheap gate: skip entries carrying nothing relevant at all.
            if archetype not in entry.archetypes and not any(
                c in entry.archetypes for c in chain
            ):
                continue
            try:
                spec = self.load(entry.entry_id)
            except (FileNotFoundError, json.JSONDecodeError):
                continue

            for layout in spec.layouts:
                score = 0.0
                if layout.archetype == archetype:
                    score += EXACT_ARCHETYPE_BONUS
                elif layout.archetype in chain:
                    score += EXACT_ARCHETYPE_BONUS - FALLBACK_STEP_PENALTY * (
                        chain.index(layout.archetype) + 1
                    )
                else:
                    continue    # unrelated archetype, not a candidate

                score -= STRUCTURE_WEIGHT * archetypes.signature_distance(
                    layout.signature(), wanted
                )
                score -= ASPECT_PENALTY * abs(entry.aspect_ratio - target.aspect_ratio)
                if target_client and _slug(entry.client or "") == target_client:
                    score += SAME_CLIENT_BONUS

                if best is None or score > best[1]:
                    best = (_clone_layout(layout, f"bank:{entry.entry_id}"), score)

        if best is None or best[1] < min_score:
            return None
        return best


class TemplateBank(BaseBank):
    """File-backed store of archived Style Specs."""

    def __init__(self, root: str | Path = DEFAULT_BANK_ROOT):
        self.root = Path(root)
        self.entries_dir = self.root / "entries"
        self.index_path = self.root / INDEX_NAME

    # -- index -------------------------------------------------------------

    def _read_index(self) -> list:
        if not self.index_path.exists():
            return []
        try:
            raw = json.loads(self.index_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
        return [BankEntry(**row) for row in raw.get("entries", [])]

    def _write_index(self, entries: list) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        payload = {
            "bank_version": "1.0",
            "updated_at": _now(),
            "entries": [vars(e) for e in entries],
        }
        self.index_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    def list_entries(self) -> list:
        return self._read_index()

    def entry(self, entry_id: str) -> BankEntry | None:
        for e in self._read_index():
            if e.entry_id == entry_id:
                return e
        return None

    # -- paths -------------------------------------------------------------

    def entry_dir(self, entry_id: str) -> Path:
        return self.entries_dir / entry_id

    def spec_path(self, entry_id: str) -> Path:
        return self.entry_dir(entry_id) / "style_spec.json"

    def asset_dir(self, entry_id: str) -> Path:
        return self.entry_dir(entry_id) / "assets"

    def master_path(self, entry_id: str) -> Path:
        """The archived master `.pptx`, whether or not it exists yet."""
        return self.entry_dir(entry_id) / "master.pptx"

    def has_master(self, entry_id: str) -> bool:
        return self.master_path(entry_id).exists()

    # -- writing -----------------------------------------------------------

    def allocate(self, client: str | None, source_name: str | None) -> str:
        """
        Reserve an entry id and its asset directory *before* extraction, so
        the extractor can write the logo straight into the bank rather than
        into a temp dir that then has to be copied.
        """
        stem = _slug(client or source_name or "master")
        existing = {e.entry_id for e in self._read_index()}
        entry_id = stem
        n = 2
        while entry_id in existing or self.entry_dir(entry_id).exists():
            entry_id = f"{stem}-{n}"
            n += 1
        self.asset_dir(entry_id).mkdir(parents=True, exist_ok=True)
        return entry_id

    def save(
        self,
        spec: StyleSpec,
        entry_id: str | None = None,
        master_pptx: str | Path | None = None,
    ) -> str:
        """
        Archive `spec` (and optionally the master it came from), returning
        the entry id.

        Asset paths inside the spec are rewritten to be relative to the
        entry directory, so the bank stays portable: it can be zipped, moved
        between machines, or committed without absolute paths leaking in.
        """
        if entry_id is None:
            entry_id = self.allocate(spec.meta.client, spec.meta.source_name)

        entry_dir = self.entry_dir(entry_id)
        entry_dir.mkdir(parents=True, exist_ok=True)
        self.asset_dir(entry_id).mkdir(parents=True, exist_ok=True)

        if master_pptx is not None:
            source = Path(master_pptx)
            dest = self.master_path(entry_id)
            if source.exists() and source.resolve() != dest.resolve():
                shutil.copy2(source, dest)

        _relativize_assets(spec, entry_dir)
        self.spec_path(entry_id).write_text(spec.to_json(), encoding="utf-8")

        entries = self._read_index()
        now = _now()
        existing = next((e for e in entries if e.entry_id == entry_id), None)
        if existing is None:
            entries.append(BankEntry(
                entry_id=entry_id,
                client=spec.meta.client,
                project=spec.meta.project,
                source_name=spec.meta.source_name,
                spec_version=spec.meta.spec_version,
                revision=1,
                created_at=now,
                updated_at=now,
                slide_width=spec.slide_width,
                slide_height=spec.slide_height,
                archetypes=sorted(spec.archetypes_present()),
            ))
        else:
            existing.updated_at = now
            existing.client = spec.meta.client or existing.client
            existing.project = spec.meta.project or existing.project
            existing.slide_width = spec.slide_width
            existing.slide_height = spec.slide_height
            existing.archetypes = sorted(spec.archetypes_present())
        self._write_index(entries)
        return entry_id

    def refine(self, entry_id: str, spec: StyleSpec) -> int:
        """
        Replace an entry's spec with a corrected one, snapshotting the
        previous revision first.

        This is the feedback path: a designer fixes something during review,
        the correction is folded back into the archived spec, and the next
        deck for that client inherits the fix. Returns the new revision.
        """
        current = self.spec_path(entry_id)
        if not current.exists():
            raise FileNotFoundError(f"No banked spec for entry '{entry_id}'")

        entries = self._read_index()
        row = next((e for e in entries if e.entry_id == entry_id), None)
        revision = (row.revision if row else 1) + 1

        rev_dir = self.entry_dir(entry_id) / "revisions"
        rev_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(current, rev_dir / f"rev-{revision - 1}.json")

        _relativize_assets(spec, self.entry_dir(entry_id))
        spec.meta.notes.append(f"revision {revision} saved {_now()}")
        current.write_text(spec.to_json(), encoding="utf-8")

        if row:
            row.revision = revision
            row.updated_at = _now()
            row.archetypes = sorted(spec.archetypes_present())
            self._write_index(entries)
        return revision

    # -- reading -----------------------------------------------------------

    def load(self, entry_id: str) -> StyleSpec:
        """Load an archived spec, resolving asset paths back to absolute."""
        path = self.spec_path(entry_id)
        if not path.exists():
            raise FileNotFoundError(f"No banked spec for entry '{entry_id}'")
        spec = StyleSpec.load(path)
        _absolutize_assets(spec, self.entry_dir(entry_id))
        return spec


# --- asset path handling --------------------------------------------------

def _asset_holders(spec: StyleSpec):
    """Every object in a spec that carries an `asset_path`."""
    yield spec.brand.logo
    yield spec.master_background
    for layout in spec.layouts:
        yield layout.background


def _relativize_assets(spec: StyleSpec, entry_dir: Path) -> None:
    """
    Rewrite absolute asset paths to be relative to the entry directory,
    copying any file that lives outside it into `assets/` on the way.
    """
    asset_dir = entry_dir / "assets"
    for holder in _asset_holders(spec):
        path_str = getattr(holder, "asset_path", None)
        if not path_str:
            continue
        path = Path(path_str)
        if not path.is_absolute():
            continue        # already relative; nothing to do
        if not path.exists():
            holder.asset_path = None
            continue
        try:
            rel = path.relative_to(entry_dir)
        except ValueError:
            asset_dir.mkdir(parents=True, exist_ok=True)
            dest = asset_dir / path.name
            if path.resolve() != dest.resolve():
                shutil.copy2(path, dest)
            rel = dest.relative_to(entry_dir)
        holder.asset_path = rel.as_posix()


def _absolutize_assets(spec: StyleSpec, entry_dir: Path) -> None:
    for holder in _asset_holders(spec):
        path_str = getattr(holder, "asset_path", None)
        if not path_str:
            continue
        path = Path(path_str)
        if path.is_absolute():
            continue
        resolved = entry_dir / path
        holder.asset_path = str(resolved) if resolved.exists() else None


def _clone_layout(layout: LayoutSpec, source: str) -> LayoutSpec:
    """A copy of a banked layout, retagged with where it came from."""
    from copy import deepcopy
    clone = deepcopy(layout)
    clone.source = source
    return clone


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
