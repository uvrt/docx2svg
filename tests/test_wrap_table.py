"""A table beside a floating drawing, and a drawing anchored in a table cell, against what
Word drew, offline (ROADMAP.md, "Floating drawings -- measured", F.15 to F.17).

``tests/fixtures/wrap-table-observations.json`` holds, for ``tools/make_wrap_table_probe.py``
and ``tools/make_cell_anchor_probe.py``, Word's text objects and pictures and every face
number the renderer asked for (``tools/read_wrap_table_probe.py``).  Each document is laid
out here from those numbers alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_wrap_table_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # A table centred in the room beside a drawing, 2,999 twips wide: its second cell's
    # text a pixel right of Word's (a tie of the centring, 3 lines).  (The floating
    # tables of the last family are laid out since make_float_table_probe.py: 18 / 18.)
    "wrap-table-none": {"jc": 3},
    "wrap-table-14": {"jc": 3},
    "wrap-table-15": {"jc": 3},
    # Below mode 15 a drawing positioned against the page that text wraps around stops the
    # table (Word moves the rows clear of it): the wrap family's layoutInCell-off pages,
    # last, are not drawn (136 lines).  And a drawing at the top of a cell's second
    # paragraph is 0.05 px higher in Word than four stacked lines put it, so the first
    # paragraph's last line is beside it, or below it (11 lines).  In mode 15 a centred
    # table's second cell draws its text a pixel right of Word's (8 lines: the table's own
    # centring, not the drawing).
    "cell-anchor-none": {"wrap": 147},
    "cell-anchor-14": {"wrap": 147},
    "cell-anchor-15": {"cell": 8},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_tables_and_drawings_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["images"] == recorded["images_score"]
    assert result["warnings"] == recorded["warnings"]
    misses = {family: total - agree for family, (agree, total) in result["families"].items() if agree != total}
    assert misses == KNOWN[name]
    # Every glyph the model draws is one Word drew.
    assert result["result"].extra == 0


def test_a_drawing_in_a_cell_is_positioned_in_it_where_word_positions_it():
    """In the cell with layoutInCell, in mode 15 whatever it says, and below it without
    layoutInCell where an axis is against the character or the line."""
    from docx2svg.floating import in_cell
    from docx2svg.model import Anchor, AnchorPosition

    def anchor(h="column", v="paragraph", cell=False):
        return Anchor((100, 100), h=AnchorPosition(h, offset=0), v=AnchorPosition(v, offset=0), layout_in_cell=cell)

    assert in_cell(anchor(cell=True), False) and in_cell(anchor(), True)
    assert not in_cell(anchor(), False)
    assert in_cell(anchor(h="character"), False) and in_cell(anchor(v="line"), False)


def test_every_alignment_against_a_cells_margin_is_its_top():
    from fractions import Fraction

    from docx2svg.floating import Frames, vertical
    from docx2svg.model import Anchor, AnchorPosition

    frames = Frames(11906, 16838, 2000, 5000, 1800, 16838 - 1800, 700, 650, 1, False, Fraction(1800),
                    Fraction(1800), Fraction(2069), Fraction(2000), Fraction(1990), Fraction(1800), cell=True)
    for align in ("top", "center", "bottom"):
        assert vertical(Anchor((635000, 317500), v=AnchorPosition("margin", align=align)), frames) == 1800
