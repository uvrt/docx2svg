"""Rows off their table's grid (round 53, R107; ROADMAP.md, "Tables -- measured", "A row off
its grid").

A row whose cells reach past ``w:tblGrid`` -- a row copied from a table with another grid, a
cell added without a grid column, a merge tracked as Word tracks one -- used to stop every
layout with an ``IndexError``.  Word lays it out on a grid widened to the row, and a fixed
table at its own width: ``tests/fixtures/grid-mismatch-observations.json`` holds what Word
16.106 drew for ``tools/make_grid_mismatch_probe.py`` (``tools/read_grid_mismatch_probe.py
--record``), and every glyph is held to it here, offline.  Made-up documents only.
"""

from __future__ import annotations

import functools
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import probe_docx  # noqa: E402
import read_grid_mismatch_probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

import docx2svg  # noqa: E402
from docx2svg import table as table_model  # noqa: E402
from docx2svg.parse import parse_package  # noqa: E402

DATA = json.loads(read_grid_mismatch_probe.OBSERVATIONS.read_text(encoding="utf-8"))

#: Word's grid is within a twip of the rule's (``table.fixed_widths``): a glyph at a cell's
#: start may be a fifth of a pixel off at the probe's 300 dpi (0.07 px at 96 dpi).
TWIP_AT_300_DPI = 300 / 1440


@functools.cache
def _documents() -> dict[str, bytes]:
    return dict(read_grid_mismatch_probe.documents())


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name", sorted(DATA["documents"]))
def test_every_glyph_where_word_drew_it(name, fonts):
    recorded = DATA["documents"][name]
    result, rects, warnings = read_render.score(name, _documents()[name], fonts, recorded["objects"],
                                                recorded["fills"])
    assert read_render.row(result) == recorded["glyphs"]
    assert warnings == recorded["warnings"] == ["table-grid-mismatch"]
    assert rects == recorded["rects"]
    assert result.extra == 0 and result.not_drawn == 0 and result.matched == result.model
    assert result.matched == min(result.y_agree, result.face_agree, result.size_agree)
    assert result.worst_dx <= TWIP_AT_300_DPI + 1e-3, result.problems[:5]


# -- the rules, one by one -------------------------------------------------------------------------


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _cell(text: str, width: int | None = 2000, span: int = 1, extra: str = "") -> str:
    props = (f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else "") + (
        f'<w:gridSpan w:val="{span}"/>' if span > 1 else "") + extra
    return f"<w:tc><w:tcPr>{props}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def _docx(rows: list[str], grid: tuple = (2000, 2000, 2000), width: str = '<w:tblW w:w="6000" w:type="dxa"/>',
          layout: str = '<w:tblLayout w:type="fixed"/>', row_properties: tuple = ()) -> bytes:
    columns = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    body = "".join(f"<w:tr>{row_properties[k] if k < len(row_properties) else ''}{row}</w:tr>"
                   for k, row in enumerate(rows))
    table = (f"<w:tbl><w:tblPr>{width}{layout}</w:tblPr><w:tblGrid>{columns}</w:tblGrid>{body}</w:tbl>"
             "<w:p><w:r><w:t>After.</w:t></w:r></w:p>")
    return probe_docx.package(table)


def _resolved(data: bytes):
    document = parse_package(data)
    table = next(item for item in document.body if isinstance(item, docx2svg.model.Table))
    return document, table_model.resolve_table(document, table)


THREE = _cell("a") + _cell("b") + _cell("c")


def test_a_cell_past_the_grid_adds_a_column_its_width():
    _, resolved = _resolved(_docx([THREE, THREE + _cell("d", 1200), THREE]))
    assert resolved.grid == (2000, 2000, 2000, 1200)
    assert resolved.mismatched == ((1, 4),)


def test_a_cell_with_no_width_past_the_grid_adds_one_of_360():
    _, resolved = _resolved(_docx([THREE, THREE + _cell("d", None)]))
    assert resolved.grid == (2000, 2000, 2000, table_model.UNSIZED_COLUMN)


def test_a_span_past_the_grid_adds_what_its_width_leaves():
    _, resolved = _resolved(_docx([THREE, _cell("a") + _cell("b", 4000, span=2) + _cell("c")]))
    assert resolved.grid == (2000, 2000, 2000, 2000)


def test_a_shorter_row_leaves_the_grid_and_is_reported():
    _, resolved = _resolved(_docx([THREE, _cell("a") + _cell("b")]))
    assert resolved.grid == (2000, 2000, 2000)
    assert resolved.mismatched == ((1, 2),)


def test_a_fixed_table_is_as_wide_as_it_says():
    # Measured: a grid of four 2,000s in a 6,000-twip fixed table is four of 1,500; 2,000 x 3 and
    # 1,200 there, margins of 108: Word 1,661 / 1,662 / 1,662 / 1,015.
    document, resolved = _resolved(_docx([THREE + _cell("d")] * 2, grid=(2000,) * 4))
    assert table_model.table_widths(document, resolved, 9026) == (1500, 1500, 1500, 1500)
    margins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
    document, resolved = _resolved(_docx([THREE, THREE + _cell("d", 1200)],
                                         width='<w:tblW w:w="6000" w:type="dxa"/>' + margins))
    widths = table_model.table_widths(document, resolved, 9026)
    assert sum(widths) == 6000 and all(abs(a - b) <= 1 for a, b in zip(widths, (1661, 1662, 1662, 1015)))


def test_a_table_with_no_width_keeps_its_columns():
    document, resolved = _resolved(_docx([THREE, THREE + _cell("d")], width=""))
    assert table_model.table_widths(document, resolved, 9026) == (2000, 2000, 2000, 2000)


@pytest.mark.parametrize("rows, grid, row_properties", [
    # one cell more, the layout's first crash (round 53: IndexError in cell_box)
    ([THREE, THREE + _cell("d")], (2000, 2000, 2000), ()),
    # a span far past the grid
    ([THREE, _cell("a", 6000, span=9)], (2000, 2000, 2000), ()),
    # w:gridBefore past the grid
    ([THREE, THREE], (2000, 2000, 2000), ('<w:trPr><w:gridBefore w:val="5"/></w:trPr>',)),
    # no grid at all
    ([THREE, THREE + _cell("d")], (), ()),
    # a tracked merge, Word's own form: the cell merged away kept, with w:cellDel
    ([THREE, _cell("ab", 4000, span=2) + _cell("", extra='<w:cellDel w:id="9" w:author="A" w:date="2026-11-12T09:00:00Z"/>')
      + _cell("c")], (2000, 2000, 2000), ()),
])
@pytest.mark.parametrize("layout", ['<w:tblLayout w:type="fixed"/>', ""])
def test_a_row_off_its_grid_never_stops_the_layout(rows, grid, row_properties, layout):
    options = docx2svg.ConvertOptions()
    svgs = docx2svg.convert_docx_to_svg(_docx(rows, grid=grid, layout=layout, row_properties=row_properties), options)
    assert svgs and "After." in "".join(svgs)
    assert "table-grid-mismatch" in {w.code for w in options.warnings}


def test_a_grid_line_past_the_last_is_the_last():
    assert table_model.line_at([0, 10, 20], 7) == 20
    assert table_model.line_at([0, 10, 20], -1) == 0


def test_a_failure_laying_a_table_out_names_the_table(monkeypatch):
    """What the model cannot have foreseen still says where: docx2svg.LayoutError, its path the
    table's, the failure its cause (docx-agent reports it as ``layout_failed``)."""
    def broken(*args, **kwargs):
        raise IndexError("list index out of range")

    monkeypatch.setattr(table_model, "flow_table", broken)
    with pytest.raises(docx2svg.LayoutError) as raised:
        docx2svg.convert_docx_to_layout(_docx([THREE, THREE]))
    assert raised.value.path == "w:body/w:tbl[1]"
    assert isinstance(raised.value.__cause__, IndexError)
    assert "w:body/w:tbl[1]" in str(raised.value) and "IndexError" in str(raised.value)
    import pickle

    again = pickle.loads(pickle.dumps(raised.value))
    assert again.path == raised.value.path and again.reason == raised.value.reason
