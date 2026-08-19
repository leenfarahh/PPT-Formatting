"""
An in-memory stand-in for the Supabase client.

The Supabase backend is worth testing properly - it's the system of record
for client material - but pointing the suite at a real project would make
it slow, networked, and dependent on credentials nobody has in CI. This
implements the slice of the client the code actually uses: chained
select/insert/upsert/update queries, and a storage bucket that holds bytes.

It's a test double, so it is deliberately strict about the calls it
supports. If the backend starts using a query form this doesn't implement,
the tests should fail loudly rather than silently pass against a stub.
"""
from __future__ import annotations

import copy


class FakeResponse:
    def __init__(self, data):
        self.data = data


class FakeQuery:
    """A chainable query over one in-memory table."""

    def __init__(self, table: "FakeTable"):
        self.table = table
        self.op = None
        self.payload = None
        self.filters: list = []
        self.limit_to: int | None = None
        self.order_by: tuple | None = None
        self.on_conflict: str | None = None

    # -- builders ---------------------------------------------------------

    def select(self, *_columns, **_kwargs):
        self.op = "select"
        return self

    def insert(self, row):
        self.op = "insert"
        self.payload = copy.deepcopy(row)
        return self

    def upsert(self, row, on_conflict: str | None = None):
        self.op = "upsert"
        self.payload = copy.deepcopy(row)
        self.on_conflict = on_conflict
        return self

    def update(self, patch):
        self.op = "update"
        self.payload = copy.deepcopy(patch)
        return self

    def delete(self):
        self.op = "delete"
        return self

    def eq(self, column, value):
        self.filters.append((column, value))
        return self

    def limit(self, count):
        self.limit_to = count
        return self

    def order(self, column, desc: bool = False):
        self.order_by = (column, desc)
        return self

    # -- execution --------------------------------------------------------

    def _matches(self, row) -> bool:
        return all(row.get(column) == value for column, value in self.filters)

    def execute(self) -> FakeResponse:
        if self.op == "select":
            rows = [r for r in self.table.rows if self._matches(r)]
            if self.order_by:
                column, desc = self.order_by
                rows.sort(key=lambda r: r.get(column) or "", reverse=desc)
            if self.limit_to is not None:
                rows = rows[: self.limit_to]
            return FakeResponse(copy.deepcopy(rows))

        if self.op == "insert":
            self.table.rows.append(self.payload)
            return FakeResponse([copy.deepcopy(self.payload)])

        if self.op == "upsert":
            key = self.on_conflict or self.table.primary_key
            existing = next(
                (r for r in self.table.rows if r.get(key) == self.payload.get(key)), None
            )
            if existing is None:
                self.table.rows.append(self.payload)
                return FakeResponse([copy.deepcopy(self.payload)])
            # Postgres updates only the columns supplied, so the double does too.
            existing.update(self.payload)
            return FakeResponse([copy.deepcopy(existing)])

        if self.op == "update":
            updated = []
            for row in self.table.rows:
                if self._matches(row):
                    row.update(self.payload)
                    updated.append(copy.deepcopy(row))
            return FakeResponse(updated)

        if self.op == "delete":
            kept, removed = [], []
            for row in self.table.rows:
                (removed if self._matches(row) else kept).append(row)
            self.table.rows = kept
            return FakeResponse(removed)

        raise NotImplementedError(f"FakeQuery does not implement {self.op!r}")


class FakeTable:
    def __init__(self, name: str, primary_key: str):
        self.name = name
        self.primary_key = primary_key
        self.rows: list = []


class FakeStorageBucket:
    def __init__(self, store: dict, bucket: str):
        self.store = store
        self.bucket = bucket

    def upload(self, path: str, data: bytes, file_options: dict | None = None):
        key = (self.bucket, path)
        upsert = str((file_options or {}).get("upsert", "false")).lower() == "true"
        if key in self.store and not upsert:
            raise RuntimeError(f"object already exists: {path}")
        self.store[key] = bytes(data)
        return {"path": path}

    def download(self, path: str) -> bytes:
        try:
            return self.store[(self.bucket, path)]
        except KeyError:
            raise FileNotFoundError(f"no object at {self.bucket}/{path}") from None

    def remove(self, paths):
        for path in paths:
            self.store.pop((self.bucket, path), None)

    def list(self, prefix: str = ""):
        return [
            {"name": path}
            for (bucket, path) in self.store
            if bucket == self.bucket and path.startswith(prefix)
        ]


class FakeStorage:
    def __init__(self):
        self.objects: dict = {}

    def from_(self, bucket: str) -> FakeStorageBucket:
        return FakeStorageBucket(self.objects, bucket)


class FakeSupabaseClient:
    """Mimics `supabase.Client` for the operations this codebase performs."""

    PRIMARY_KEYS = {
        "bank_entries": "entry_id",
        "bank_revisions": "id",
        "jobs": "job_id",
    }

    def __init__(self):
        self.tables = {
            name: FakeTable(name, key) for name, key in self.PRIMARY_KEYS.items()
        }
        self.storage = FakeStorage()

    def table(self, name: str) -> FakeQuery:
        if name not in self.tables:
            raise KeyError(f"unknown table {name!r}")
        return FakeQuery(self.tables[name])

    # -- test helpers -----------------------------------------------------

    def rows(self, name: str) -> list:
        return self.tables[name].rows

    def object_paths(self) -> list:
        return sorted(f"{bucket}/{path}" for bucket, path in self.storage.objects)
