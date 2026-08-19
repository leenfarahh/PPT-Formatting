"""
Runtime configuration and backend selection.

The tool runs against either the local filesystem or Supabase. Which one is
a deployment decision, not a code one, so it's read from the environment and
everything downstream takes a bank and a job store rather than knowing where
they live.

    PPTX_STORAGE_BACKEND   "local" (default) or "supabase"
    PPTX_BANK_ROOT         local bank directory
    PPTX_CACHE_DIR         where Supabase objects are cached locally
    SUPABASE_URL           project URL
    SUPABASE_SERVICE_KEY   service-role key, server-side only
    SUPABASE_ANON_KEY      fallback if no service key is set

Local stays the default deliberately: the test suite, the demo and anyone
trying the tool offline shouldn't need a Supabase project to exist.
"""
from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from .bank import DEFAULT_BANK_ROOT

BACKEND_LOCAL = "local"
BACKEND_SUPABASE = "supabase"

# Storage buckets. Create these in the Supabase dashboard (or via
# supabase/schema.sql) and keep all three private.
BUCKET_MASTERS = "masters"
BUCKET_ASSETS = "assets"
BUCKET_OUTPUTS = "outputs"


@dataclass
class Settings:
    backend: str = BACKEND_LOCAL
    bank_root: Path = DEFAULT_BANK_ROOT
    cache_dir: Path = field(
        default_factory=lambda: Path(tempfile.gettempdir()) / "pptx_formatter_cache"
    )
    supabase_url: str | None = None
    supabase_key: str | None = None

    @property
    def uses_supabase(self) -> bool:
        return self.backend == BACKEND_SUPABASE

    def describe(self) -> dict:
        """Backend summary safe to send to a browser - no key material."""
        info = {"backend": self.backend}
        if self.uses_supabase:
            info["supabase_url"] = self.supabase_url
            info["configured"] = bool(self.supabase_url and self.supabase_key)
        else:
            info["bank_root"] = str(self.bank_root)
        return info


def settings_from_env(env: dict | None = None) -> Settings:
    env = os.environ if env is None else env
    backend = (env.get("PPTX_STORAGE_BACKEND") or BACKEND_LOCAL).strip().lower()
    if backend not in (BACKEND_LOCAL, BACKEND_SUPABASE):
        raise ValueError(
            f"PPTX_STORAGE_BACKEND must be '{BACKEND_LOCAL}' or '{BACKEND_SUPABASE}', "
            f"got {backend!r}"
        )

    bank_root = Path(env.get("PPTX_BANK_ROOT") or DEFAULT_BANK_ROOT)
    cache_dir = Path(
        env.get("PPTX_CACHE_DIR") or Path(tempfile.gettempdir()) / "pptx_formatter_cache"
    )
    return Settings(
        backend=backend,
        bank_root=bank_root,
        cache_dir=cache_dir,
        supabase_url=env.get("SUPABASE_URL"),
        # Prefer the service-role key: the server needs to read and write
        # rows that row-level security hides from anonymous callers. It must
        # never be handed to a browser.
        supabase_key=env.get("SUPABASE_SERVICE_KEY") or env.get("SUPABASE_ANON_KEY"),
    )


def make_client(settings: Settings):
    """Build a Supabase client, with a clear error when it can't be."""
    if not settings.supabase_url or not settings.supabase_key:
        raise RuntimeError(
            "Supabase backend selected but SUPABASE_URL and SUPABASE_SERVICE_KEY "
            "are not both set."
        )
    try:
        from supabase import create_client
    except ImportError as exc:      # pragma: no cover - depends on install
        raise RuntimeError(
            "The supabase package is not installed. Run "
            "`pip install -r requirements.txt`, or set "
            "PPTX_STORAGE_BACKEND=local to use the filesystem."
        ) from exc
    return create_client(settings.supabase_url, settings.supabase_key)


def make_bank(settings: Settings | None = None, client=None):
    """The Template Bank for the configured backend."""
    settings = settings or settings_from_env()
    if settings.uses_supabase:
        from .supabase_bank import SupabaseBank
        return SupabaseBank(
            client or make_client(settings), cache_dir=settings.cache_dir
        )
    from .bank import TemplateBank
    return TemplateBank(settings.bank_root)


def make_job_store(settings: Settings | None = None, client=None):
    """The job store for the configured backend."""
    settings = settings or settings_from_env()
    if settings.uses_supabase:
        from .jobs import SupabaseJobStore
        return SupabaseJobStore(
            client or make_client(settings), cache_dir=settings.cache_dir
        )
    from .jobs import LocalJobStore
    return LocalJobStore(settings.bank_root.parent / "jobs")
