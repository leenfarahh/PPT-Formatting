"""The Template Bank: archiving, selection, and feeding corrections back."""
from __future__ import annotations

import json

from pptx_formatter import archetypes
from pptx_formatter.extraction import extract_style_spec
from pptx_formatter.layout_builder import generated_layout_spec
from pptx_formatter.style_spec import StyleSpec


def test_save_and_load_round_trip(bank, master_path):
    spec = extract_style_spec(master_path, client="Acme")
    entry_id = bank.save(spec, master_pptx=master_path)

    reloaded = bank.load(entry_id)
    assert reloaded.theme.colors == spec.theme.colors
    assert len(reloaded.layouts) == len(spec.layouts)
    assert bank.has_master(entry_id)


def test_index_records_available_archetypes(bank, master_path):
    spec = extract_style_spec(master_path, client="Acme", project="Board")
    entry_id = bank.save(spec, master_pptx=master_path)

    entry = bank.entry(entry_id)
    assert entry.client == "Acme"
    assert entry.project == "Board"
    assert archetypes.TITLE_SLIDE in entry.archetypes
    assert entry.revision == 1


def test_asset_paths_are_stored_relative(bank, master_path, tmp_path):
    """The bank has to stay portable, so absolute paths can't leak into it."""
    from examples.make_sample_master import write_solid_png

    logo = tmp_path / "logo.png"
    write_solid_png(logo, 60, 20, (1, 2, 3))

    spec = extract_style_spec(master_path, client="Acme")
    spec.brand.logo.asset_path = str(logo)
    spec.brand.logo.left_frac = 0.8
    entry_id = bank.save(spec, master_pptx=master_path)

    raw = json.loads(bank.spec_path(entry_id).read_text(encoding="utf-8"))
    stored = raw["brand"]["logo"]["asset_path"]
    assert not stored.startswith(str(tmp_path))
    assert stored == "assets/logo.png"

    # Loading resolves it back to something openable.
    from pathlib import Path
    assert Path(bank.load(entry_id).brand.logo.asset_path).exists()


def test_latest_for_client_matches_on_a_slug(bank, master_path):
    """"Acme Holdings" and "acme-holdings" are the same account."""
    spec = extract_style_spec(master_path, client="Acme Holdings")
    bank.save(spec, master_pptx=master_path)

    assert bank.latest_for_client("acme-holdings") is not None
    assert bank.latest_for_client("ACME HOLDINGS") is not None
    assert bank.latest_for_client("Someone Else") is None


def test_exact_archetype_match_wins_selection(bank, master_path):
    spec = extract_style_spec(master_path, client="Donor")
    quote = generated_layout_spec(archetypes.QUOTE, spec)
    quote.name = "Donor Quote"
    spec.layouts.append(quote)
    bank.save(spec, "donor", master_pptx=master_path)

    target = extract_style_spec(master_path, client="Target")
    chosen, score = bank.select_layout(archetypes.QUOTE, target)

    assert chosen.archetype == archetypes.QUOTE
    assert chosen.source == "bank:donor"
    assert score > 50


def test_selection_falls_back_to_a_near_archetype(bank, master_path):
    """
    A bank with no quote layout but a section header should offer the
    section header rather than nothing - it's the nearest structural match.
    """
    spec = extract_style_spec(master_path, client="Donor")
    spec.layouts = [l for l in spec.layouts if l.archetype == archetypes.SECTION_HEADER]
    bank.save(spec, "donor", master_pptx=master_path)

    target = extract_style_spec(master_path, client="Target")
    hit = bank.select_layout(archetypes.QUOTE, target)
    assert hit is not None
    assert hit[0].archetype == archetypes.SECTION_HEADER


def test_selection_ignores_unrelated_archetypes(bank, master_path):
    spec = extract_style_spec(master_path, client="Donor")
    spec.layouts = [l for l in spec.layouts if l.archetype == archetypes.TITLE_SLIDE]
    bank.save(spec, "donor", master_pptx=master_path)

    target = extract_style_spec(master_path, client="Target")
    assert bank.select_layout(archetypes.TABLE, target) is None


def test_same_client_house_style_is_preferred(bank, master_path):
    """Between two equal candidates, the client's own layout should win."""
    other = extract_style_spec(master_path, client="Other Co")
    other.layouts.append(generated_layout_spec(archetypes.QUOTE, other))
    bank.save(other, "other", master_pptx=master_path)

    theirs = extract_style_spec(master_path, client="Acme")
    theirs.layouts.append(generated_layout_spec(archetypes.QUOTE, theirs))
    bank.save(theirs, "acme", master_pptx=master_path)

    target = extract_style_spec(master_path, client="Acme")
    chosen, _ = bank.select_layout(archetypes.QUOTE, target)
    assert chosen.source == "bank:acme"


def test_refine_keeps_the_previous_revision(bank, master_path):
    """A designer's correction has to be auditable and reversible."""
    spec = extract_style_spec(master_path, client="Acme")
    entry_id = bank.save(spec, master_pptx=master_path)

    corrected = bank.load(entry_id)
    corrected.theme.colors["accent1"] = "FF0000"
    revision = bank.refine(entry_id, corrected)

    assert revision == 2
    assert bank.entry(entry_id).revision == 2
    assert bank.load(entry_id).theme.colors["accent1"] == "FF0000"

    archived = json.loads(
        (bank.entry_dir(entry_id) / "revisions" / "rev-1.json").read_text(encoding="utf-8")
    )
    assert archived["theme"]["colors"]["accent1"] == "0F4C81"


def test_entry_ids_do_not_collide(bank, master_path):
    spec = extract_style_spec(master_path, client="Acme")
    first = bank.save(spec, master_pptx=master_path)
    second = bank.save(extract_style_spec(master_path, client="Acme"))
    assert first != second
