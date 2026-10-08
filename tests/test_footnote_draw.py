"""Footnotes against what Word drew, offline (ROADMAP.md, "Phase 4 -- measured", 4.13).

``tests/fixtures/footnote-draw-observations.json`` holds, for
``tools/make_footnote_draw_probe.py``, Word's text objects and filled rectangles and every
face number the renderer asked for (``tools/read_footnote_draw_probe.py``).  Each document
is laid out here from those numbers alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_footnote_draw_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


def test_every_document_is_recorded():
    assert sorted(DATA["documents"]) == sorted(name for name, _ in reader.documents())


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_footnotes_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["separators"] == recorded["separators"]
    assert result["pages"] == recorded["pages"]
    assert result["warnings"] == recorded["warnings"]
    # Every line, every number and every separator where Word drew it, on Word's pages.
    assert result["lines"][0] == result["lines"][1]
    assert result["separators"] == [0, 0]
    assert result["pages"][0] == result["pages"][1]
    assert result["result"].extra == 0
    assert result["result"].not_drawn == 0


def test_a_last_line_keeps_its_space_after_above_the_notes_in_mode_15():
    """``make_note_after_probe.py``: a candidate line with 240 twips after leaves the page
    240 twips sooner in mode 15, and where it would without in no settings and mode 14
    (ROADMAP.md, F.24)."""
    import read_note_after_probe as after

    recorded = json.loads(after.OBSERVATIONS.read_text(encoding="utf-8"))
    fonts = render_record.RecordedFonts(recorded["faces"])
    for name, data in after.documents():
        assert after.model_side(data, fonts) == recorded["documents"][name], name
    moved = {name: sides.index(1) for name, sides in recorded["documents"].items()}
    assert moved["noteafter-after240-15"] == moved["noteafter-after0-15"] - 12
    assert moved["noteafter-after240-14"] == moved["noteafter-after0-14"] == moved["noteafter-after240-none"]


def _numbers(name: str, pages: dict | None = None) -> list[str]:
    import make_footnote_draw_probe as probe

    from docx2svg import parse_package
    from docx2svg.notes import footnote_numbers

    return footnote_numbers(parse_package(probe.build(name)), pages)


def test_a_section_numbers_its_footnotes_by_their_place_in_the_document():
    # decimal from 5; upperRoman restarting; lowerLetter from 3, restarting; chicago going
    # on from the whole document's count.
    assert _numbers("fndraw-format-none") == ["5", "6", "7", "I", "II", "III", "c", "d", "e",
                                              "†††", "‡‡‡", "§§§"]
    assert _numbers("fndraw-chain-15") == ["1", "2", "9", "10", "1", "2", "7", "8"]


def test_the_documents_footnote_properties_are_not_taken():
    assert _numbers("fndraw-docpr-15") == ["1", "2", "3"]
    assert _numbers("fndraw-docpage-14") == [str(k) for k in range(1, 17)]


def test_each_page_restarts_from_the_pages_found():
    pages = {note: page for note, page in zip(range(1, 17), [0] * 2 + [1] * 3 + [2] * 9 + [3] * 2)}
    assert _numbers("fndraw-eachpage-none", pages) == ["1", "2", "1", "2", "3"] + [str(k) for k in range(1, 10)] + [
        "1", "2"]


def test_mode_15_opens_a_continued_page_with_the_separator():
    import make_footnote_draw_probe as probe

    from docx2svg import parse_package
    from docx2svg.notes import footnote_separator_kind, separator

    for name, kind in (("fndraw-separators-14", "continuationSeparator"), ("fndraw-separators-15", "separator")):
        document = parse_package(probe.build(name))
        assert footnote_separator_kind(document, "continuationSeparator") == kind
        assert separator(document, kind, "footnote")[0].path.startswith("w:footnote[")
    # Not named by the settings: Word's own.
    document = parse_package(probe.build("fndraw-unnamed-15"))
    assert separator(document, "separator", "footnote")[0].path == "w:footnotes/separator"
