"""Tab leaders against what Word drew, offline (ROADMAP.md, "Phase 5 -- measured", 5.19).

``tests/fixtures/tab-leader-observations.json`` holds, for ``tools/make_tab_leader_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_tab_leader_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_tab_leader_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


def test_every_document_is_recorded():
    assert sorted(DATA["documents"]) == sorted(name for name, _ in reader.documents())


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_tab_leaders_against_word(name, data, fonts):
    pytest.importorskip("numpy")
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["leaders"] == recorded["leaders"]
    assert result["rules"] == recorded["rules"]
    assert result["pages"] == recorded["pages"]
    assert result["warnings"] == recorded["warnings"]
    # Every line and every leader glyph where Word drew it, every rule Word's.
    assert result["lines"][0] == result["lines"][1]
    assert result["leaders"][0] == result["leaders"][1] > 0
    assert result["rules"] == [0, 0]
    assert result["result"].extra == 0 and result["result"].not_drawn == 0


def test_a_leader_is_the_stops_own():
    from docx2svg.linebreak import Geometry, TabStop, tab_leader

    geometry = Geometry(0, 0, 10 ** 7, tabs=(TabStop(1000, "left", "dot"), TabStop(5000, "right", "none")))
    assert tab_leader(0, geometry, True) == "dot"
    assert tab_leader(2000, geometry, True) is None  # a leader of "none"
    assert tab_leader(6000, geometry, True) is None  # a default stop


def test_leader_glyphs():
    from docx2svg.layout import LEADER_GLYPHS

    assert LEADER_GLYPHS == {"dot": ".", "hyphen": "-", "underscore": "_", "heavy": "_", "middleDot": "·"}
