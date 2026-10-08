"""Autofit tables sized from their content against what Word drew, offline (ROADMAP.md,
"Tables -- measured", stage 7b).

``tests/fixtures/autofit-width-observations.json`` holds, for
``tools/make_autofit_width_probe.py``, Word's text objects and filled rectangles and every
face number the renderer asked for (``tools/read_autofit_width_probe.py``).  Each document
is laid out here from those numbers alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_autofit_width_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # The last family, which the layout stops at: a cell across columns wider than they
    # are, whose extra Word shares unevenly by a rule not settled (two rounds, nine
    # tables: ROADMAP.md, Tables -- measured, stage 7b).
    "autofit-width-none": {"span": 37},
    "autofit-width-14": {"span": 37},
    "autofit-width-15": {"span": 37},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_autofit_tables_against_word(name, data, fonts):
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


def _resolved(cells, extents, width=(0, "auto"), mode=None):
    """A one-row table of cells with ``cells`` (``w:tcW`` or ``None``), borders of w:sz 4
    and margins of 108, resolved, with each cell's ``(narrowest, widest)`` content."""
    from docx2svg.model import Document, Table
    from docx2svg.table import resolve_table

    table = Table(rows=(tuple(() for _ in cells),), grid=tuple(1000 for _ in cells),
                  properties={"tblW": width, "tblCellMar.left": (108, "dxa"), "tblCellMar.right": (108, "dxa"),
                              "tblBorders.left": ("single", 4, 0, "000000"),
                              "tblBorders.right": ("single", 4, 0, "000000")},
                  row_properties=({},), cell_properties=(tuple({"tcW": w} if w else {} for w in cells),))
    document = Document(compatibility_mode=mode)
    resolved = resolve_table(document, table)
    resolved.extents = {(0, k): (Fraction(low), Fraction(high)) for k, (low, high) in enumerate(extents)}
    return document, resolved


def test_columns_with_no_width_are_as_wide_as_their_content():
    from docx2svg.table import autofit_widths

    assert autofit_widths(*_resolved([None, None], [(600, 744.9), (1359, 1803.3)]), 9164) == (745, 1804)


def test_a_word_wider_than_its_width_widens_the_column_and_the_table():
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([(1440, "dxa")] * 3, [(711, 952), (2121.7, 2677), (699, 939)]), 9164)
    assert widths == (1440, 2122, 1440)


def test_content_wider_than_the_room_is_shared_by_widest_less_narrowest():
    """60 words beside 10: Word drew the first column 7,430 twips of 9,272 (9,164 and the
    last cell's right margin); the rule, 7,430."""
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([None, None], [(711.4, 26563.6), (721.6, 5030.3)]), 9164)
    assert sum(widths) == 9272 and widths[0] == 7430


def test_columns_in_twips_keep_their_width_while_the_others_give_way():
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([(3000, "dxa"), None], [(711.4, 26563.6), (721.6, 26573.8)]), 9164)
    assert widths == (3000, 6272)


def test_in_mode_15_the_room_is_the_column_less_the_outer_borders():
    from docx2svg.table import autofit_widths

    widths = autofit_widths(*_resolved([None, None], [(711.4, 26563.6), (721.6, 26573.8)], mode=15), 9164)
    assert sum(widths) == 9154


def test_a_stated_table_width_widens_the_columns_by_their_content_or_their_width():
    from docx2svg.table import autofit_widths

    assert autofit_widths(*_resolved([None, None], [(711.4, 951.7), (978.3, 1533.6)], (6000, "dxa")),
                          9164) == (2298, 3702)
    assert autofit_widths(*_resolved([(1000, "dxa"), (3000, "dxa")], [(711, 952), (722, 962)], (6000, "dxa")),
                          9164) == (1500, 4500)
    assert autofit_widths(*_resolved([(1000, "dxa"), None], [(711, 952), (722, 962)], (6000, "dxa")),
                          9164) == (1000, 5000)


def test_what_the_probe_did_not_settle_stops_the_table():
    from docx2svg.table import Unsupported, autofit_widths

    with pytest.raises(Unsupported):  # columns in twips giving way beside ones at their narrowest
        autofit_widths(*_resolved([(8000, "dxa"), None], [(711, 952), (5000, 26000)]), 9164)
    with pytest.raises(Unsupported):  # in percent, narrower than its content, below mode 15
        autofit_widths(*_resolved([None, None], [(5000, 26000), (5000, 26000)], (5000, "pct")), 9164)
