"""End-to-end test: a complete run_check() on a synthetic folder.

Builds a throw-away folder with a fake Excel list and real (tiny) photos,
runs the whole pipeline, and checks every outcome the app can produce:
matched, index-only, duplicate name, reversed name, unknown photo,
missing photo - plus the reports, decision log, and runs/ archive.
"""
import csv
import json

import pytest

import passport_check as pc
from conftest import make_excel, make_photo


@pytest.fixture
def workspace(tmp_path, monkeypatch, test_config):
    """A fake app folder with Excel + images, HERE pointed at it."""
    monkeypatch.setattr(pc, "HERE", tmp_path)
    make_excel(tmp_path / "list.xlsx", [
        (1, "MARIJA", "STANKOVIC", "123456"),   # photo present -> stamped
        (2, "SASA", "STANKOVIC", "654321"),     # no photo -> missing
        (3, "PETAR", "PETROVIC", "333"),        # duplicate name ->
        (4, "PETAR", "PETROVIC", "444"),        #   manual check
        (5, "ANA", "KOVAC", ""),                # photo, no ID -> index only
        (6, "KRSTA", "KOVACEVIC", "555"),       # photo has reversed name
    ])
    imgs = tmp_path / "images"
    make_photo(imgs / "MARIJA STANKOVIC.jpg")
    make_photo(imgs / "PETAR PETROVIC.jpg")
    make_photo(imgs / "ANA KOVAC.png")
    make_photo(imgs / "KOVACEVIC KRSTA.jpg")    # back-to-front
    make_photo(imgs / "NOBODY UNKNOWN.jpg")     # not in the list at all
    (imgs / "SOME SCAN.pdf").write_bytes(b"%PDF-1.4 fake")   # skipped
    return tmp_path


def run(cfg):
    return pc.run_check(cfg, log=lambda *a, **k: None)


def test_full_run_outcomes(workspace, test_config):
    results = run(test_config)

    # matched + stamped
    assert [(n, i) for _, n, i in results["stamped"]] == \
        [("MARIJA STANKOVIC", "123456")]
    out = workspace / "output"
    assert (out / "001 - MARIJA STANKOVIC.jpg").exists()

    # photo present but no passport number -> stamped with list number only
    assert results["index_only"] == [("5", "ANA KOVAC")]
    assert (out / "005 - ANA KOVAC.png").exists()

    # duplicate name -> flagged manual, nothing stamped
    assert results["manual"][0][0] == "PETAR PETROVIC.jpg"
    assert sorted(results["manual"][0][1]) == ["333", "444"]
    assert not any(f.name.endswith("PETAR PETROVIC.jpg")
                   for f in out.iterdir())

    # people with no usable photo (KRSTA's photo has a reversed name,
    # so it is correctly NOT counted as received)
    assert sorted(results["missing_image"]) == \
        ["KRSTA KOVACEVIC", "SASA STANKOVIC"]

    # unmatched photos, with the right diagnosis
    diags = dict(results["not_found"])
    assert diags["KOVACEVIC KRSTA.jpg"]["reason"] == "reversed"
    assert diags["KOVACEVIC KRSTA.jpg"]["suggestion"] == "KRSTA KOVACEVIC.jpg"
    assert diags["NOBODY UNKNOWN.jpg"]["reason"] == "unknown"

    # PDF skipped, stats consistent
    assert results["pdf_skipped"] == ["SOME SCAN.pdf"]
    assert results["stats"]["total"] == 5          # 4 unique + 1 duplicated name
    assert results["stats"]["stamped"] == 1
    assert results["stats"]["needs_manual"] == 3   # 2 unmatched + 1 duplicate


def test_full_run_reports_and_archive(workspace, test_config):
    results = run(test_config)

    # reports next to the app
    assert (workspace / "report.txt").exists()
    assert (workspace / "report.xlsx").exists()
    report_text = (workspace / "report.txt").read_text(encoding="utf-8")
    assert pc.__version__ in report_text
    assert "MARIJA STANKOVIC" in report_text

    # decision log: one line per photo + per missing person
    with open(workspace / "decisions.csv", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f, delimiter=";"))
    assert rows[0] == pc.DECISION_COLUMNS
    by_file = {r[0]: r for r in rows[1:]}
    assert by_file["MARIJA STANKOVIC.jpg"][2] == "matched - number added"
    assert by_file["MARIJA STANKOVIC.jpg"][4] == "123456"
    assert by_file["NOBODY UNKNOWN.jpg"][2] == "no match (unknown)"
    assert by_file["SOME SCAN.pdf"][2] == "pdf skipped"
    missing_rows = [r for r in rows[1:] if r[2] == "no photo received"]
    assert sorted(r[3] for r in missing_rows) == \
        ["KRSTA KOVACEVIC", "SASA STANKOVIC"]

    # dated archive with everything needed to compare runs later
    run_dirs = list((workspace / "runs").iterdir())
    assert len(run_dirs) == 1
    archived = {p.name for p in run_dirs[0].iterdir()}
    assert {"report.txt", "report.xlsx", "decisions.csv",
            "summary.json"} <= archived
    summary = json.loads(
        (run_dirs[0] / "summary.json").read_text(encoding="utf-8"))
    assert summary["version"] == pc.__version__
    assert summary["stats"] == results["stats"]


def test_second_run_gets_its_own_archive_folder(workspace, test_config):
    run(test_config)
    # a second run must never overwrite the first archive
    import time
    time.sleep(1.1)   # archive folders are named per second
    run(test_config)
    assert len(list((workspace / "runs").iterdir())) == 2


def test_originals_are_never_modified(workspace, test_config):
    img = workspace / "images" / "MARIJA STANKOVIC.jpg"
    before = img.read_bytes()
    run(test_config)
    assert img.read_bytes() == before
