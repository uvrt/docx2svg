"""Endnotes against what Word drew, offline (ROADMAP.md, "Phase 4 -- measured", 4.12).

``tests/fixtures/endnote-observations.json`` holds, for ``tools/make_endnote_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_endnote_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_endnote_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_endnotes_against_word(name, data, fonts):
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


def _numbered(name: str):
    import make_endnote_probe as probe

    from docx2svg import parse_package
    from docx2svg.notes import with_endnotes

    document, placed = with_endnotes(parse_package(probe.build(name)))
    return document, placed


def test_references_are_numbered_by_their_sections_endnote_properties():
    document, _ = _numbered("endnote-format-none")
    numbers = [run.text for paragraph in document.body if paragraph.note is None for run in paragraph.runs
               if run.note_number]
    # decimal from 5; upperRoman restarting; lowerLetter from 3, restarting.
    assert numbers == ["5", "6", "I", "II", "c", "d"]
    notes = [run.text for paragraph in document.body if paragraph.note == "endnote" for run in paragraph.runs
             if run.note_number]
    assert notes == numbers


def test_the_documents_endnote_format_is_not_taken():
    document, _ = _numbered("endnote-docpr-15")
    numbers = [run.text for paragraph in document.body if paragraph.note is None for run in paragraph.runs
               if run.note_number]
    assert numbers == ["1", "2", "iii"]


def test_notes_go_after_their_section_under_sectend_and_at_the_end_otherwise():
    document, placed = _numbered("endnote-sectend-15")
    kinds = [paragraph.note for paragraph in document.body]
    assert kinds == [None, None, "separator", "endnote", "endnote"] * 3
    assert [(s.first_paragraph, s.last_paragraph) for s in document.sections] == [(0, 5), (5, 10), (10, 15)]
    assert len(placed) == 9
    document, _ = _numbered("endnote-format-none")
    kinds = [paragraph.note for paragraph in document.body]
    assert kinds == [None] * 6 + ["separator"] + ["endnote"] * 6


def test_word_draws_its_own_separator_unless_the_settings_name_the_parts():
    import make_endnote_probe as probe

    from docx2svg import parse_package
    from docx2svg.notes import separator

    for name, own in (("endnote-separators-none", False), ("endnote-unnamed-15", False),
                      ("endnote-separators-15", True)):
        paragraph = separator(parse_package(probe.build(name)), "separator")[0]
        assert (paragraph.path.startswith("w:endnote[")) is own, name


def test_a_note_numbers_advance_is_whole_device_pixels():
    from fractions import Fraction

    from docx2svg.linebreak import note_number_units

    # Calibri "i" (470 / 2048) at 7 pt, drawn at 29 px: 6.66 px, laid out 7.
    assert note_number_units((470, 2048), 14) == 7 * Fraction(72 * 4096, 300)
