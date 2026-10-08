"""Autofit tables narrower than their narrowest content against what Word drew, offline
(ROADMAP.md, "Tables -- measured", stage 7c).

``tests/fixtures/autofit-over-observations.json`` holds, for
``tools/make_autofit_over_probe.py``, Word's text objects and filled rectangles and every
face number the renderer asked for (``tools/read_autofit_over_probe.py``).  Each document is
laid out here from those numbers alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from test_autofit_width import _resolved

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_autofit_over_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, and why: below mode 15 a table
#: in percent narrower than its content is drawn past the margin with its words broken by
#: a rule not settled, so the layout stops at it (the last case).
KNOWN = {"autofit-over-none": 3, "autofit-over-14": 3, "autofit-over-15": 0}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_tables_narrower_than_their_content_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["borders"] == recorded["borders"]
    assert result["warnings"] == recorded["warnings"]
    agree, total = result["lines"]
    assert total - agree == KNOWN[name]
    assert result["result"].extra == 0


def test_the_columns_share_the_room_by_their_narrowest_content():
    """Nine columns in dxa whose widest words need 12,332 twips with their margins: Word
    keeps the table to the room (9,272 below mode 15), each column its 216 twips of margins
    and the rest in proportion to its widest word."""
    from docx2svg.table import autofit_widths

    words = [1334, 1048, 1144, 1048, 953, 1144, 1048, 1430, 1239]
    stated = [1400, 1080, 940, 940, 880, 940, 880, 1040, 920]
    widths = autofit_widths(*_resolved([(w, "dxa") for w in stated], [(w + 216, w + 216) for w in words]), 9164)
    assert sum(widths) == 9272
    assert widths[0] == 1157 and widths[4] == 888 and widths[7] == 1225


def test_a_stated_width_in_dxa_gives_way_to_the_room():
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([(2000, "dxa")] * 3, [(3552, 3552)] * 3, (6000, "dxa")), 9164)
    assert sum(widths) == 9272


def test_in_mode_15_the_room_is_the_column_less_the_outer_borders():
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([None] * 3, [(3552, 3552)] * 3, mode=15), 9164)
    assert widths == (3051, 3051, 3052)


def test_a_table_in_percent_narrower_than_its_content_still_stops_below_mode_15():
    from docx2svg.table import Unsupported, autofit_widths

    with pytest.raises(Unsupported):
        autofit_widths(*_resolved([None] * 3, [(3552, 3552)] * 3, (5000, "pct")), 9164)
