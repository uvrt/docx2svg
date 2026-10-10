"""A floating drawing text wraps around in a table cell aligned vertically or merged
(ROADMAP.md, "Floating drawings -- measured", F.25), and the layout's approximations.

The documents are generated here, as a production report built them (a two-by-two table,
the first row merged across or not, ``w:vAlign`` on its first cell, a picture anchored in
it with ``layoutInCell`` and ``allowOverlap``, positioned against the column and the
paragraph, wrapped per case; a cover page's logo far left of its cell).  Laid out with
uniform advances and Calibri's line, no font needed.  Word's side is in
``tests/test_wrap_table.py`` (``cell-valign-14`` / ``-15``).
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_anchor_probe as anchor_probe  # noqa: E402
import make_cell_valign_probe as valign_probe  # noqa: E402
import read_anchor_probe  # noqa: E402
import render_record  # noqa: E402
import wml  # noqa: E402
from make_anchor_probe import Anchor  # noqa: E402

import docx2svg  # noqa: E402
from docx2svg import fonts  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402


class Uniform:
    """Half an em per character, Calibri's line."""

    def advance(self, face, bold, italic, char):
        return 1024, 2048

    def kern(self, face, bold, italic, left, right):
        return 0

    def metrics(self, face, bold=False, italic=False):
        return FaceMetrics(2048, 1950, 550, 0, 1331, 1331, 976, 286)

    def decorations(self, face, bold=False, italic=False):
        return fonts.Decorations(2048, 976, 286, 512, 134, -232, 134, 1950, 550)

    def drawing_name(self, face, bold=False, italic=False):
        return None


SPACING = valign_probe.SPACING
#: The report's cover page: a logo far left of a cell merged across, a little up.
TEMPLATE = dict(h=("column", "offset", -6536055), v=("paragraph", "offset", -5715), cx=6092190, cy=2708910)


def _p(text: str, runs: str = "") -> str:
    return wml.paragraph(wml.run(text) + runs, mark={}, spacing=SPACING)


def _cell(paragraphs: str, *, span: int = 1, valign: str | None = None) -> str:
    props = f'<w:tcW w:w="{4000 * span}" w:type="dxa"/>'
    props += f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
    props += f'<w:vAlign w:val="{valign}"/>' if valign else ""
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{paragraphs}</w:tc>"


def document(valign: str | None, wrap: str, merged: bool, *, setting: str = "15", h=("column", "offset", 0),
             v=("paragraph", "offset", 0), cx: int = 1200000, cy: int = 500000, simple=None,
             neighbour: int = 1) -> bytes:
    """The report's document: a paragraph, the table, six paragraphs after it
    (``neighbour`` paragraphs in the first row's second cell)."""
    anchor = Anchor(h, v, cx, cy, layout_in_cell=True, wrap=valign_probe.WRAPS[wrap], simple=simple).xml(1)
    first = _cell(_p("cell 00 ", anchor), span=2 if merged else 1, valign=valign)
    row1 = first if merged else first + _cell("".join(_p(f"cell 01.{k}") for k in range(neighbour)))
    row2 = _cell(_p("cell 10")) + _cell(_p("cell 11"))
    props = (f'<w:tblW w:w="8000" w:type="dxa"/><w:tblBorders>{valign_probe.BORDERS}</w:tblBorders>'
             '<w:tblLayout w:type="fixed"/><w:tblLook w:val="0000"/>')
    table = (f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{'<w:gridCol w:w=\"4000\"/>' * 2}</w:tblGrid>"
             f"<w:tr>{row1}</w:tr><w:tr>{row2}</w:tr></w:tbl>")
    body = _p("Before the table.") + table + "".join(_p(f"After the table, paragraph {k}.") for k in range(6))
    return anchor_probe.package(body, setting)


def lay_out(data: bytes):
    options = render_record.options(Uniform())
    return docx2svg.convert_docx_to_layout(data, options), options


REPORTED = [  # (vAlign, wrap, merged across): the report's isolation, every one complete in Word
    ("center", "through", True), ("center", "through", False), ("center", "square", False),
    ("center", "none", False), (None, "through", False), (None, "through", True),
]


@pytest.mark.parametrize("setting", ["none", "15"])
@pytest.mark.parametrize("valign,wrap,merged", REPORTED, ids=["-".join(map(str, c)) for c in REPORTED])
def test_the_reported_cases_are_laid_out_to_the_end(valign, wrap, merged, setting):
    layout, _options = lay_out(document(valign, wrap, merged, setting=setting))
    coverage = layout.coverage
    assert coverage.stop is None, coverage.summary()
    assert coverage.complete and coverage.status == "complete" and coverage.approximations == []
    assert (coverage.blocks_laid_out, coverage.blocks) == (8, 8)


@pytest.mark.parametrize("setting", ["none", "15"])
def test_the_reported_cover_page_is_laid_out_to_the_end(setting):
    """vAlign center, merged across, wrapThrough, the logo 6,536,055 EMU left of its column:
    it stays in the cell (its left edge on the inside of the cell's left border), and the
    cell's text goes below it, as Word lays it out."""
    layout, _options = lay_out(document("center", "through", True, setting=setting, **TEMPLATE))
    coverage = layout.coverage
    assert coverage.status == "complete" and (coverage.blocks_laid_out, coverage.blocks) == (8, 8)
    [picture] = read_anchor_probe.model_images(layout)
    # Not 6,536,055 EMU (10,292 twips) left of the cell: on its inside-left edge, right of
    # the page's left margin (1,442 twips, 300.4 device px).
    assert 300 < picture[1] < 305
    assert _first_line(layout, "cell 00").baseline > picture[4]  # below the logo, wider than the cell


def _text(line) -> str:
    return "".join(char for span in line.spans for char in span.chars)


def _first_line(layout, prefix: str):
    return next(line for line in layout.pages[0].text_lines() if _text(line).strip().startswith(prefix))


@pytest.mark.parametrize("wrap", ["square", "through", "topAndBottom", "none"])
def test_a_centred_cell_moves_its_text_and_its_drawing_together(wrap):
    """Laid out from the cell's top, then moved down by half the room left beside the
    taller of the text and the drawing (``make_cell_valign_probe.py``, ``align``)."""
    top, _ = lay_out(document("top", wrap, False, neighbour=8))
    centred, _ = lay_out(document("center", wrap, False, neighbour=8))
    [top_picture] = read_anchor_probe.model_images(top)
    [centred_picture] = read_anchor_probe.model_images(centred)
    moved = centred_picture[2] - top_picture[2]
    assert moved > 0
    assert _first_line(centred, "cell 00").baseline - _first_line(top, "cell 00").baseline == pytest.approx(moved,
                                                                                                         abs=0.6)
    # The rows below did not move: the drawing is shorter than the row.
    assert _first_line(centred, "cell 10").baseline == _first_line(top, "cell 10").baseline


def test_a_drawing_placed_as_no_probe_measured_is_approximated_not_stopped():
    """A drawing text wraps around in a cell, aligned against the character (not measured
    in a cell): the table and everything after it are laid out, and the coverage says
    where the layout approximated."""
    data = document("center", "square", False, h=("character", "align", "left"))
    layout, options = lay_out(data)
    coverage = layout.coverage
    assert coverage.stop is None and coverage.complete
    assert (coverage.blocks_laid_out, coverage.blocks) == (8, 8)
    assert coverage.status == "approximate"
    [approximation] = coverage.approximations
    assert approximation.code == "layout-approximate:cell-drawing" and approximation.page == 1
    assert approximation.path == "w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]/w:r[2]/wp:anchor[1]"
    assert "character left" in approximation.message
    assert coverage.as_dict()["status"] == "approximate"
    assert coverage.summary().startswith("complete with approximations: 1 page(s), 8 block(s); 1 approximated")
    assert "layout-approximate:cell-drawing" in {w.code for w in options.warnings}
    # Placed by its anchor, and drawn.
    assert len(read_anchor_probe.model_images(layout)) == 1


def test_a_partial_layout_is_partial_whatever_it_approximated():
    from docx2svg.coverage import Coverage, StopInfo

    approximation = StopInfo("layout-approximate:cell-drawing", "cell-drawing", "", 1, "w:body/w:tbl[1]")
    assert Coverage(True, 1, 1, "layout", 2, 2, 0).status == "complete"
    assert Coverage(True, 1, 1, "layout", 2, 2, 0, approximations=[approximation]).status == "approximate"
    assert Coverage(False, 1, None, None, 2, 1, 1, approximations=[approximation]).status == "partial"
