"""A floating table (``w:tblpPr``) and the text beside it against what Word drew, offline
(ROADMAP.md, "Floating drawings -- measured", F.18).

``tests/fixtures/float-table-observations.json`` holds, for ``tools/make_float_table_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_float_table_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_float_table_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # A table 340 twips from the column's left edge: its second cell's text on a tie
    # (722.5 px), a pixel left of Word's (3 lines); a table centred in the margin: its
    # rows a pixel below Word's (1 line, the others' baselines agree).
    "float-table-none": {"room": 3, "v": 1},
    "float-table-14": {"room": 3, "v": 1},
    # Mode 15: a centred table's second cell on a tie (1,277.5 px), a pixel right of
    # Word's (9 lines); a table with no borders aligned right, which Word draws 15 twips
    # further left, as if it had a border there (6 lines); the centred table again, and
    # two tables whose top is on a half pixel (312.5 px), a pixel below Word's (7 lines);
    # and the last case, a table past the bottom margin, which Word splits across two
    # pages in mode 15: the layout stops there (10 lines).
    "float-table-15": {"edge": 6, "h": 9, "no case": 10, "v": 7},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_floating_tables_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["borders"] == recorded["borders"]
    assert result["warnings"] == recorded["warnings"]
    misses = {family: total - agree for family, (agree, total) in result["families"].items() if agree != total}
    assert misses == KNOWN[name]
    # Every glyph the model draws is one Word drew.
    assert result["result"].extra == 0


def _floating(**attributes) -> dict:
    return {"tblpPr": {key: str(value) for key, value in attributes.items()}}


def test_a_table_at_zero_zero_against_the_margin_is_in_the_flow():
    """Every position 0 and vertAnchor not text: Word lays the table out in the flow."""
    from docx2svg.model import Document, Table
    from docx2svg.table import floating_of

    document, table = Document(), Table()
    assert floating_of(document, table, _floating(vertAnchor="margin")) is None
    assert floating_of(document, table, _floating(vertAnchor="page", tblpY=0)) is None
    assert floating_of(document, table, _floating(vertAnchor="text")).vertical == "text"


def test_below_mode_15_a_table_at_tblpY_0_against_the_margin_is_placed_against_the_text():
    from docx2svg.model import Document, Table
    from docx2svg.table import floating_of

    table = Table()
    assert floating_of(Document(), table, _floating(vertAnchor="margin", tblpX=1000)).vertical == "text"
    assert floating_of(Document(compatibility_mode=15), table,
                       _floating(vertAnchor="margin", tblpX=1000)).vertical == "margin"
    assert floating_of(Document(), table, _floating(vertAnchor="margin", tblpY=1)).vertical == "margin"


def test_what_the_probe_did_not_measure_stops_the_table():
    from docx2svg.model import Document, Table
    from docx2svg.table import Unsupported, floating_of

    for attributes in ({"vertAnchor": "text", "tblpXSpec": "inside"}, {"vertAnchor": "text", "tblpYSpec": "top"},
                       {"vertAnchor": "margin", "horzAnchor": "page"}):
        with pytest.raises(Unsupported):
            floating_of(Document(), Table(), _floating(**attributes))


def test_text_keeps_at_least_ten_twips_off_and_fifteen_past_a_borderless_edge():
    from docx2svg.table import MIN_FROM_TEXT, UNBORDERED_EDGE

    assert (MIN_FROM_TEXT, UNBORDERED_EDGE) == (10, 15)
