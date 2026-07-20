"""Unit tests for the name-matching brain of passport_check.

These protect the risky parts: if a change breaks how names are compared
or how mistakes are diagnosed, a test here goes red immediately.
"""
import pytest

import passport_check as pc
from conftest import make_excel


# --------------------------------------------------------------------------- #
# normalize / clean_stem
# --------------------------------------------------------------------------- #
def test_normalize_lowercase_and_spaces():
    assert pc.normalize("  MARIJA   STANKOVIC ") == "marija stankovic"


def test_normalize_strips_accents():
    assert pc.normalize("Đurđa Jovanović") == pc.normalize("Đurđa Jovanović")
    assert pc.normalize("Élise Müller") == "elise muller"


def test_normalize_none_and_empty():
    assert pc.normalize(None) == ""
    assert pc.normalize("   ") == ""


def test_clean_stem_strips_trailing_number_suffix():
    assert pc.clean_stem("ANDREJA HASANOVIC.101") == "andreja hasanovic"


def test_clean_stem_keeps_normal_names():
    assert pc.clean_stem("MILA LUCIC") == "mila lucic"


# --------------------------------------------------------------------------- #
# diagnose_unmatched - one test per reason the app can report
# --------------------------------------------------------------------------- #
PEOPLE = {
    "marija stankovic": {"name": "MARIJA STANKOVIC", "id": "111", "index": "1"},
    "ana kovac": {"name": "ANA KOVAC", "id": "222", "index": "2"},
}


def diag(key, fname="X.jpg"):
    return pc.diagnose_unmatched(key, fname, PEOPLE, {})


def test_diagnose_reversed_name():
    d = diag("stankovic marija", "STANKOVIC MARIJA.jpg")
    assert d["reason"] == "reversed"
    assert d["suggestion"] == "MARIJA STANKOVIC.jpg"


def test_diagnose_junk_words_in_filename():
    d = diag("marija stankovic passport scan 2")
    assert d["reason"] == "junk"
    assert d["suggestion"].startswith("MARIJA STANKOVIC")


def test_diagnose_two_people_in_one_file():
    d = diag("ana kovac i marija stankovic")
    assert d["reason"] == "two_people"
    assert "ANA KOVAC" in d["why"] and "MARIJA STANKOVIC" in d["why"]


def test_diagnose_spelling_mistake():
    d = diag("marija stankovik", "MARIJA STANKOVIK.jpg")
    assert d["reason"] == "spelling"
    assert d["suggestion"] == "MARIJA STANKOVIC.jpg"


def test_diagnose_partial_match_on_surname():
    d = diag("bob stankovic")
    assert d["reason"] == "partial"
    assert "MARIJA STANKOVIC" in d["why"]


def test_diagnose_completely_unknown():
    d = diag("xxxx yyyy")
    assert d["reason"] == "unknown"
    assert d["suggestion"] == ""


def test_diagnose_keeps_file_extension_in_suggestion():
    d = pc.diagnose_unmatched("stankovic marija", "STANKOVIC MARIJA.png",
                              PEOPLE, {})
    assert d["suggestion"].endswith(".png")


# --------------------------------------------------------------------------- #
# _pick_column
# --------------------------------------------------------------------------- #
HEADERS = ["Index", "Name", "Surname", "Passport no"]


def test_pick_column_exact_header():
    assert pc._pick_column(HEADERS, "Name", ("zzz",)) == 1


def test_pick_column_numeric_index():
    assert pc._pick_column(HEADERS, "3", ("zzz",)) == 3


def test_pick_column_auto_by_keyword():
    assert pc._pick_column(HEADERS, "auto", ("passport",)) == 3
    assert pc._pick_column(HEADERS, "auto", ("surname",)) == 2


def test_pick_column_not_found_raises():
    with pytest.raises(pc.ColumnNotFound):
        pc._pick_column(HEADERS, "auto", ("telephone",))


# --------------------------------------------------------------------------- #
# find_excel / read_excel (on real temporary .xlsx files)
# --------------------------------------------------------------------------- #
def test_find_excel_single_file(tmp_path, monkeypatch):
    monkeypatch.setattr(pc, "HERE", tmp_path)
    make_excel(tmp_path / "list.xlsx", [])
    assert pc.find_excel("auto").name == "list.xlsx"


def test_find_excel_ignores_our_own_report(tmp_path, monkeypatch):
    monkeypatch.setattr(pc, "HERE", tmp_path)
    make_excel(tmp_path / "list.xlsx", [])
    make_excel(tmp_path / "report.xlsx", [])
    assert pc.find_excel("auto").name == "list.xlsx"


def test_find_excel_none_raises(tmp_path, monkeypatch):
    monkeypatch.setattr(pc, "HERE", tmp_path)
    with pytest.raises(pc.ExcelNotFound):
        pc.find_excel("auto")


def test_find_excel_several_raises(tmp_path, monkeypatch):
    monkeypatch.setattr(pc, "HERE", tmp_path)
    make_excel(tmp_path / "a.xlsx", [])
    make_excel(tmp_path / "b.xlsx", [])
    with pytest.raises(pc.ExcelNotFound):
        pc.find_excel("auto")


def test_read_excel_people_and_duplicates(tmp_path, monkeypatch, test_config):
    monkeypatch.setattr(pc, "HERE", tmp_path)
    make_excel(tmp_path / "list.xlsx", [
        (1, "MARIJA", "STANKOVIC", 123456),      # numeric id -> must lose '.0'
        (2, "PETAR", "PETROVIC", "333"),         # same name twice ->
        (3, "PETAR", "PETROVIC", "444"),         #   goes to duplicates
        (4, "ANA", "KOVAC", ""),                 # no passport number yet
    ])
    people, duplicates = pc.read_excel(test_config)

    assert people["marija stankovic"]["id"] == "123456"
    assert people["marija stankovic"]["index"] == "1"
    assert people["ana kovac"]["id"] == ""
    assert "petar petrovic" not in people
    assert sorted(duplicates["petar petrovic"]["ids"]) == ["333", "444"]
