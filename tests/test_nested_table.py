"""A table nested in a table cell against what Word drew, offline (ROADMAP.md, "Tables --
measured", "A table nested in a cell").

``tests/fixtures/nested-table-observations.json`` holds, for
``tools/make_nested_table_probe.py``, Word's text objects and filled rectangles and every
face number the renderer asked for (``tools/read_nested_table_probe.py``).  Each document
is laid out here from those numbers alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_nested_table_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # A nested table of 4000 pct centred in its cell: its cells' text a pixel right of
    # Word's (the centring's rounding), over cells in percent and in dxa (6 lines).
    "nested-table-none": {"sample": 4, "width": 2},
    "nested-table-14": {"sample": 4, "width": 2},
    # Mode 15: a nested table aligned right, a pixel right of Word's (2 lines).
    "nested-table-15": {"place": 2},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_nested_tables_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["borders"] == recorded["borders"]
    assert result["warnings"] == recorded["warnings"]
    misses = {family: total - agree for family, (agree, total) in result["families"].items() if agree != total}
    assert misses == KNOWN[name]
    assert result["result"].extra == 0


def _layout(case: int):
    import make_nested_table_probe as probe

    from docx2svg import ConvertOptions, _lay_out

    cases = probe.CASES
    try:
        probe.CASES = (cases[case],)
        layout, _, _ = _lay_out(probe.build("none"), ConvertOptions())
    finally:
        probe.CASES = cases
    return layout


@pytest.mark.faces
def test_an_empty_paragraph_after_a_nested_table_takes_no_room():
    """The cell ends at the nested table's bottom border: ``After`` follows it directly."""
    import make_nested_table_probe as probe

    notes = [case.note for case in probe.CASES]
    full, empty = (_layout(notes.index(note)) for note in ("first in its cell", "the paragraph after it empty"))
    after = {name: next(line.baseline for line in layout.pages[0].text_lines() if line.path == "w:body/w:p[2]")
             for name, layout in (("full", full), ("empty", empty))}
    assert after == {"full": 591, "empty": 535}
