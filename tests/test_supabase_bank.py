"""
The Supabase backend, exercised against an in-memory client.

The important property is **parity**: the pipeline takes a bank and doesn't
care where it stores things, so a Supabase-backed run has to produce the
same result as a local one. Several of these mirror `test_bank.py`
deliberately.
"""
from __future__ import annotations

import pytest

from pptx_formatter import archetypes
from pptx_formatter.extraction import extract_style_spec
from pptx_formatter.layout_builder import generated_layout_spec
from pptx_formatter.supabase_bank import SupabaseBank

from conftest import assert_valid_pptx
from fake_supabase import FakeSupabaseClient


@pytest.fixture
def sb_client():
    return FakeSupabaseClient()


@pytest.fixture
def sb_bank(sb_client, tmp_path):
    return SupabaseBank(sb_client, cache_dir=tmp_path / "cache")


def test_save_and_load_round_trip(sb_bank, master_path):
    spec = extract_style_spec(master_path, client="Acme")
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    reloaded = sb_bank.load(entry_id)
    assert reloaded.theme.colors == spec.theme.colors
    assert reloaded.theme.fonts.minor_cs == spec.theme.fonts.minor_cs
    assert len(reloaded.layouts) == len(spec.layouts)


def test_spec_is_stored_as_a_row_and_master_as_an_object(sb_bank, sb_client, master_path):
    spec = extract_style_spec(master_path, client="Acme", project="Board")
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    rows = sb_client.rows("bank_entries")
    assert len(rows) == 1
    assert rows[0]["entry_id"] == entry_id
    assert rows[0]["client"] == "Acme"
    assert rows[0]["style_spec"]["theme"]["colors"]["accent1"] == "0F4C81"
    # Denormalized so gap-filling can skip irrelevant entries cheaply.
    assert archetypes.TITLE_SLIDE in rows[0]["archetypes"]
    assert f"masters/{entry_id}/master.pptx" in sb_client.object_paths()


def test_master_is_downloaded_on_demand(sb_bank, master_path, tmp_path):
    """Stage 2 opens the master as a file, so it must materialize locally."""
    spec = extract_style_spec(master_path, client="Acme")
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    # Drop the cache to force a fetch from storage.
    cached = sb_bank.entry_dir(entry_id) / "master.pptx"
    cached.unlink()
    assert sb_bank.has_master(entry_id)

    fetched = sb_bank.master_path(entry_id)
    assert fetched.exists()
    assert fetched.read_bytes() == master_path.read_bytes()


def test_logo_bytes_survive_a_round_trip(sb_bank, sb_client, master_path, tmp_path):
    """Without the bytes, Stage 2 knows where a logo was but can't place it."""
    from examples.make_sample_master import write_solid_png

    logo = tmp_path / "logo.png"
    write_solid_png(logo, 60, 20, (15, 76, 129))

    spec = extract_style_spec(master_path, client="Acme")
    spec.brand.logo.asset_path = str(logo)
    spec.brand.logo.left_frac = 0.85
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    assert f"assets/{entry_id}/assets/logo.png" in sb_client.object_paths()

    # Clear the cache; loading should pull the object back down to a file.
    (sb_bank.entry_dir(entry_id) / "assets" / "logo.png").unlink()
    reloaded = sb_bank.load(entry_id)
    from pathlib import Path
    assert reloaded.brand.logo.present
    assert Path(reloaded.brand.logo.asset_path).read_bytes() == logo.read_bytes()


def test_missing_asset_object_does_not_break_loading(sb_bank, master_path, tmp_path):
    """A deleted logo should cost the logo, not the whole deck."""
    from examples.make_sample_master import write_solid_png

    logo = tmp_path / "logo.png"
    write_solid_png(logo, 60, 20, (1, 2, 3))
    spec = extract_style_spec(master_path, client="Acme")
    spec.brand.logo.asset_path = str(logo)
    spec.brand.logo.left_frac = 0.85
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    sb_bank.client.storage.from_("assets").remove([f"{entry_id}/assets/logo.png"])
    (sb_bank.entry_dir(entry_id) / "assets" / "logo.png").unlink()

    reloaded = sb_bank.load(entry_id)
    assert reloaded.brand.logo.asset_path is None


def test_entry_ids_do_not_collide(sb_bank, master_path):
    first = sb_bank.save(extract_style_spec(master_path, client="Acme"))
    second = sb_bank.save(extract_style_spec(master_path, client="Acme"))
    assert first != second


def test_latest_for_client_matches_on_a_slug(sb_bank, master_path):
    sb_bank.save(extract_style_spec(master_path, client="Acme Holdings"),
                 master_pptx=master_path)
    assert sb_bank.latest_for_client("acme-holdings") is not None
    assert sb_bank.latest_for_client("Someone Else") is None


def test_selection_scoring_matches_the_local_backend(sb_bank, bank, master_path):
    """Same inputs, same choice - the scoring lives in the shared base class."""
    donor = extract_style_spec(master_path, client="Donor")
    donor.layouts.append(generated_layout_spec(archetypes.QUOTE, donor))

    sb_bank.save(donor, "donor", master_pptx=master_path)
    bank.save(extract_style_spec(master_path, client="Donor"), "donor",
              master_pptx=master_path)
    # Give the local bank the same extra layout.
    local_spec = bank.load("donor")
    local_spec.layouts.append(generated_layout_spec(archetypes.QUOTE, local_spec))
    bank.save(local_spec, "donor")

    target = extract_style_spec(master_path, client="Target")
    remote_hit = sb_bank.select_layout(archetypes.QUOTE, target)
    local_hit = bank.select_layout(archetypes.QUOTE, target)

    assert remote_hit is not None and local_hit is not None
    assert remote_hit[0].archetype == local_hit[0].archetype == archetypes.QUOTE
    assert round(remote_hit[1], 6) == round(local_hit[1], 6)


def test_refine_keeps_the_previous_revision(sb_bank, sb_client, master_path):
    spec = extract_style_spec(master_path, client="Acme")
    entry_id = sb_bank.save(spec, master_pptx=master_path)

    corrected = sb_bank.load(entry_id)
    corrected.theme.colors["accent1"] = "FF0000"
    revision = sb_bank.refine(entry_id, corrected)

    assert revision == 2
    assert sb_bank.entry(entry_id).revision == 2
    assert sb_bank.load(entry_id).theme.colors["accent1"] == "FF0000"

    archived = sb_client.rows("bank_revisions")
    assert len(archived) == 1
    assert archived[0]["revision"] == 1
    assert archived[0]["style_spec"]["theme"]["colors"]["accent1"] == "0F4C81"


def test_refine_on_a_missing_entry_is_a_clear_error(sb_bank, master_path):
    with pytest.raises(FileNotFoundError, match="No banked spec"):
        sb_bank.refine("nope", extract_style_spec(master_path))


def test_pipeline_runs_end_to_end_on_supabase(sb_bank, master_path, content_path, tmp_path):
    """The whole point: the pipeline shouldn't know where the bank lives."""
    from pptx_formatter.pipeline import run_pipeline

    out = tmp_path / "formatted.pptx"
    report = run_pipeline(
        master_path, content_path, out, bank=sb_bank,
        client="Acme Holdings", project="Board Deck",
    )

    assert report["slides_processed"] == 6
    assert report["bank_entry"] is not None
    assert_valid_pptx(out)


def test_repeat_client_skips_stage_one_on_supabase(sb_bank, master_path, content_path, tmp_path):
    from pptx_formatter.pipeline import format_with_banked_spec, ingest_master

    ingest_master(master_path, bank=sb_bank, client="Acme Holdings")

    out = tmp_path / "repeat.pptx"
    report = format_with_banked_spec("Acme Holdings", content_path, out, sb_bank)

    assert report["stage_1_skipped"] is True
    assert report["slides_processed"] == 6
    assert_valid_pptx(out)
