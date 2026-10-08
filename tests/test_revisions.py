"""Tracked changes, content controls and comments against Word's final view, offline
(ROADMAP.md, "Revisions -- measured").

``tests/fixtures/revision-observations.json`` holds, for ``tools/make_revision_probe.py``,
Word's text objects and fills in its *No Markup* view, each case's lines as Word drew
them, the cases Word draws otherwise with every revision accepted, and every face number
the renderer asked for (``tools/read_revision_probe.py``).  Each document is laid out here
from those numbers alone: every case, every line, every glyph and every rule pixel as
Word has them.  The rules the probe refuted are held to their scores too.
"""

from __future__ import annotations

import dataclasses
import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import make_revision_probe as probe  # noqa: E402
import read_revision_probe as reader  # noqa: E402
import render_record  # noqa: E402

from docx2svg import paginate  # noqa: E402
from docx2svg.parse import document as parse_document  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
FAMILIES = {"move": [6, 6], "mark": [32, 32], "sdt": [11, 11], "props": [5, 5], "rows": [9, 9],
            "comment": [3, 3], "section": [4, 4]}
#: Where Word's final view is not what accepting every revision makes (every mode).
ACCEPT_DIFFERS = {"mark at a cell's end", "mark at a cell's end, next centred", "mark at a cell's end, styled",
                  "mark before a table", "mark before a table, first cell right", "mark before a table, spaced",
                  "mark before a table, styled", "rows cellMerge alone"}
#: ... and below mode 15 also where a joined paragraph's text begins a line (R.4).
DROPPED = {"mark break, second spaced", "mark break, second spaced 480", "mark chain of four, breaks, last spaced",
           "mark chain, breaks, last spaced", "mark content kept", "mark list first", "mark styled first",
           "mark two lines and a break, second spaced", "mark before a block control"}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_the_final_view_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["word"] == recorded["word"]
    assert result["disagree"] == recorded["disagree"] == []
    assert result["families"] == recorded["families"] == FAMILIES
    assert result["lines"] == recorded["lines"] == [291, 291]
    assert result["glyphs"] == recorded["glyphs"]
    assert result["rules"] == recorded["rules"] == [344556, 344556, 0, 0]
    assert result["warnings"] == recorded["warnings"] == []
    assert result["pages"] == recorded["pages"] == len(probe.CASES)


def test_accepting_every_revision_draws_otherwise_where_recorded():
    for name, _ in reader.documents():
        below = name != "revision-15"
        assert set(DATA["documents"][name]["accept_differs"]) == ACCEPT_DIFFERS | (DROPPED if below else set())


def test_word_draws_mode_14_as_it_draws_no_mode():
    assert DATA["documents"]["revision-14"]["word"] == DATA["documents"]["revision-none"]["word"]


def _first_properties(first, following):
    joined = _JOINED(first, following)
    return dataclasses.replace(joined, properties=first.properties)


_JOINED = parse_document._joined


@pytest.mark.parametrize("name,patch,families", [
    # The joined paragraph under the first paragraph's properties, not the last's.
    ("revision-15", (parse_document, "_joined", _first_properties),
     {"move": [4, 6], "mark": [11, 32], "sdt": [11, 11], "props": [5, 5], "rows": [9, 9], "comment": [3, 3],
      "section": [2, 4]}),
    # No paragraph joined at all (the model before), the rest of the final view as now.
    ("revision-15", (parse_document, "_join", lambda blocks: blocks),
     {"move": [4, 6], "mark": [8, 32], "sdt": [11, 11], "props": [4, 5], "rows": [9, 9], "comment": [3, 3],
      "section": [2, 4]}),
    # Below mode 15, no line drawn lower where a joined paragraph's text begins it.
    ("revision-none", (paginate, "join_starts", lambda paragraph, pieces, broken: []),
     {"move": [6, 6], "mark": [23, 32], "sdt": [11, 11], "props": [5, 5], "rows": [9, 9], "comment": [3, 3],
      "section": [4, 4]}),
])
def test_refuted_rules_keep_their_scores(name, patch, families, fonts):
    data = dict(reader.documents())[name]
    with mock.patch.object(*patch):
        result = reader.score(name, data, fonts, DATA["documents"][name])
    assert result["families"] == families
