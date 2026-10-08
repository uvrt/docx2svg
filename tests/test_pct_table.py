"""A table whose width is stated in percent against what Word drew, offline (ROADMAP.md,
"Tables -- measured", stage 7a).

``tests/fixtures/pct-table-observations.json`` holds, for ``tools/make_pct_table_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_pct_table_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_pct_table_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # One cell's text a pixel from Word's in three tables (six unequal shares, twice, and
    # three columns of 2000 pct): the rounding of a grid line not settled at the pixel.
    # And the last family, which the layout stops at: a column narrower than the word in
    # it, which Word widens (autofit), and what the probe did not measure.
    "pct-table-none": {"grid": 2, "not modelled": 23, "share": 2, "width": 2},
    "pct-table-14": {"grid": 2, "not modelled": 23, "share": 2, "width": 2},
    # Mode 15: one cell's text a pixel from Word's in two tables of even shares and the
    # two with margins of 300 (a tie of the rounding at 1,277.5 px among them).
    "pct-table-15": {"margin": 4, "not modelled": 23, "width": 4},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_percent_tables_against_word(name, data, fonts):
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


def _resolved(shares, width=(5000, "pct"), mode=None):
    """A one-row table of cells with ``shares`` (``(w, type)``), borders of w:sz 4 and
    margins of 108, resolved."""
    from docx2svg.model import Document, Table
    from docx2svg.table import resolve_table

    table = Table(rows=(tuple(() for _ in shares),), grid=tuple(1000 for _ in shares),
                  properties={"tblW": width, "tblCellMar.left": (108, "dxa"), "tblCellMar.right": (108, "dxa"),
                              "tblBorders.left": ("single", 4, 0, "000000"),
                              "tblBorders.right": ("single", 4, 0, "000000")},
                  row_properties=({},), cell_properties=(tuple({"tcW": share} for share in shares),))
    document = Document(compatibility_mode=mode)
    return document, resolve_table(document, table)


def test_the_grid_word_writes_for_a_table_stretched_to_the_width():
    """818 / 849 / 849 / 850 / 851 / 783 pct of a 9,360-twip column and two margins of 108:
    the grid Word wrote for them, each line at the floor of its share of 9,576."""
    from docx2svg.table import percent_widths

    shares = [(w, "pct") for w in (818, 849, 849, 850, 851, 783)]
    assert percent_widths(*_resolved(shares), 9360) == (1566, 1626, 1626, 1628, 1630, 1500)


def test_shares_under_the_width_are_scaled_and_over_it_the_last_takes_the_rest():
    from docx2svg.table import percent_widths

    assert percent_widths(*_resolved([(1000, "pct"), (2000, "pct")]), 9164) == (3126, 6254)
    assert percent_widths(*_resolved([(1000, "pct"), (4500, "pct")]), 9164) == (1876, 7504)


def test_in_mode_15_the_whole_is_the_column_less_the_outer_borders():
    from docx2svg.table import percent_widths

    assert sum(percent_widths(*_resolved([(2500, "pct"), (2500, "pct")], mode=15), 9164)) == 9154


def test_what_the_probe_did_not_measure_stops_the_table():
    from docx2svg.table import Unsupported, percent_widths

    for shares, width in (([(2500, "pct"), (3000, "dxa")], (5000, "pct")),
                          ([(4788, "dxa"), (4788, "dxa")], (3500, "pct"))):  # no minimums given
        with pytest.raises(Unsupported):
            percent_widths(*_resolved(shares, width), 9164)


def test_cells_in_twips_share_a_table_in_percent_by_their_width_and_their_narrowest_content():
    """Under the width, in proportion to their widths; over it, each gives up the excess in
    proportion to its width less its narrowest content: 4,788 / 4,788 twips under 6,566,
    with a 16-letter word (1,524.5 twips, and margins) in one and a 671-twip label in the
    other -- Word drew 3,470 / 3,096."""
    from fractions import Fraction

    from docx2svg.table import dxa_shares

    assert dxa_shares([1000, 3000], 6566, (0, 0)) == (1641, 4925)
    minimums = (Fraction(15245, 10) + 216, Fraction(6713, 10) + 216)
    assert dxa_shares([4788, 4788], 6566, minimums) == (3467, 3099)
