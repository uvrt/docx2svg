"""A table's structure changed with tracking on, drawn as Word's Accept All leaves it (round 54,
R110; ROADMAP.md, "Tables -- measured", stage 7f).

The final view is the document with every revision accepted.  For a table's structure that is
not what Word's *No Markup* screen shows: Word keeps drawing a cell with a pending deletion
(``w:cellDel``) -- a tracked column deletion, a tracked merge's cell merged away -- at its
width until the deletion is accepted, and a ``w:cellMerge`` merges nothing until then.
Measured with Word 16.106 for Mac on a made-up four-column table with its third column
deleted, tracked: No Markup draws the empty column; Accept All leaves three columns, each
keeping its 2,268 twips, the table that much narrower.  Each case here is held to the same
document with its changes accepted, drawn byte for byte the same but for the cells' paths,
which name the document's own rows and cells.  Made-up documents only.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import probe_docx  # noqa: E402

import docx2svg  # noqa: E402
from docx2svg.parse import parse_package  # noqa: E402

WHO = 'w:author="A" w:date="2026-11-12T09:00:00Z"'
DELETED = f'<w:cellDel w:id="9" {WHO}/>'


def _cell(text: str, width: int = 2268, extra: str = "", deleted: bool = False) -> str:
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>' + extra + (DELETED if deleted else "")
    if deleted:
        run = (f'<w:pPr><w:rPr><w:del w:id="7" {WHO}/></w:rPr></w:pPr>'
               f'<w:del w:id="8" {WHO}><w:r><w:delText>{text}</w:delText></w:r></w:del>')
    else:
        run = f"<w:r><w:t>{text}</w:t></w:r>"
    return f"<w:tc><w:tcPr>{props}</w:tcPr><w:p>{run}</w:p></w:tc>"


def _docx(rows: list[str], grid: tuple, width: str = '<w:tblW w:w="0" w:type="auto"/>', layout: str = "") -> bytes:
    columns = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    body = "".join(f"<w:tr>{row}</w:tr>" for row in rows)
    borders = ('<w:tblBorders>' + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                          for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
               + "</w:tblBorders>")
    return probe_docx.package(f"<w:tbl><w:tblPr>{width}{layout}{borders}</w:tblPr><w:tblGrid>{columns}</w:tblGrid>"
                              f"{body}</w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>")


def _drawn(data: bytes, options=None) -> list[str]:
    return [re.sub(r' data-docx-path="[^"]*"', "", svg) for svg in docx2svg.convert_docx_to_svg(data, options)]


REPORT = [("Area", "Lead", "Phase", "Remark"), ("Harbour", "Ines", "draft", "awaits figures"),
          ("Quarry", "Aleid", "open", "two gaps")]


@pytest.mark.parametrize("width, layout", [
    ('<w:tblW w:w="0" w:type="auto"/>', ""),                                 # the measured table
    ('<w:tblW w:w="9072" w:type="dxa"/>', '<w:tblLayout w:type="fixed"/>'),
    ('<w:tblW w:w="5000" w:type="pct"/>', ""),
])
def test_a_tracked_column_deletion_is_drawn_accepted(width, layout):
    tracked = _docx(["".join(_cell(t, deleted=k == 2) for k, t in enumerate(row)) for row in REPORT], (2268,) * 4,
                    width, layout)
    accepted = _docx(["".join(_cell(t) for k, t in enumerate(row) if k != 2) for row in REPORT], (2268,) * 3,
                     width, layout)
    options = docx2svg.ConvertOptions()
    drawn = _drawn(tracked, options)
    assert drawn == _drawn(accepted)
    assert "Phase" not in "".join(drawn)
    assert not options.warnings


def test_a_cell_after_a_deleted_one_keeps_its_path():
    tracked = _docx(["".join(_cell(t, deleted=k == 2) for k, t in enumerate(row)) for row in REPORT], (2268,) * 4)
    document = parse_package(tracked)
    table = next(item for item in document.body if isinstance(item, docx2svg.model.Table))
    assert table.grid == (2268,) * 3
    assert table.cell_paths[0] == ("w:body/w:tbl[1]/w:tr[1]/w:tc[1]", "w:body/w:tbl[1]/w:tr[1]/w:tc[2]",
                                   "w:body/w:tbl[1]/w:tr[1]/w:tc[4]")
    layout = docx2svg.convert_docx_to_layout(tracked)
    assert any('w:tr[1]/w:tc[4]' in svg for svg in docx2svg.convert_docx_to_svg(tracked))
    assert layout.pages


def test_a_deleted_row_keeps_the_paths_of_the_rows_below():
    deleted_row = "<w:trPr><w:del w:id=\"5\" " + WHO + "/></w:trPr>" + "".join(_cell(t) for t in REPORT[1])
    tracked = _docx(["".join(_cell(t) for t in REPORT[0]), deleted_row, "".join(_cell(t) for t in REPORT[2])],
                    (2268,) * 4)
    table = next(item for item in parse_package(tracked).body if isinstance(item, docx2svg.model.Table))
    assert [paths[0] for paths in table.cell_paths] == ["w:body/w:tbl[1]/w:tr[1]/w:tc[1]",
                                                        "w:body/w:tbl[1]/w:tr[3]/w:tc[1]"]


def test_a_row_of_deleted_cells_goes_with_them():
    tracked = _docx(["".join(_cell(t) for t in REPORT[0]), "".join(_cell(t, deleted=True) for t in REPORT[1]),
                     "".join(_cell(t) for t in REPORT[2])], (2268,) * 4)
    accepted = _docx(["".join(_cell(t) for t in REPORT[0]), "".join(_cell(t) for t in REPORT[2])], (2268,) * 4)
    assert _drawn(tracked) == _drawn(accepted)


def test_a_tracked_merge_down_is_drawn_merged():
    """``w:cellMerge`` (docx-agent's tracked merge down, the cell's own ``w:vMerge`` left out
    until accepted): accepting gives the cells the merge it records."""
    top = _cell("Area", extra=f'<w:cellMerge w:id="3" {WHO} w:vMerge="rest"/>') + _cell("Lead")
    below = _cell("", extra=f'<w:cellMerge w:id="4" {WHO} w:vMerge="cont"/>') + _cell("Ines")
    tracked = _docx([top, below], (2268,) * 2)
    accepted = _docx([_cell("Area", extra='<w:vMerge w:val="restart"/>') + _cell("Lead"),
                      _cell("", extra="<w:vMerge/>") + _cell("Ines")], (2268,) * 2)
    assert _drawn(tracked) == _drawn(accepted)


def test_a_deleted_cell_whose_content_is_kept_widens_the_cell_before():
    """Word's Accept All of a deleted cell whose content is not deleted gives its columns
    and width to the cell before it, the content going (measured, round 54: the merge of
    two cells written so, the first cell holding both texts, accepted to one cell over two
    columns, every glyph and rule as Word draws it)."""
    kept = f'<w:tc><w:tcPr><w:tcW w:w="2268" w:type="dxa"/>{DELETED}</w:tcPr><w:p><w:r><w:t>Ines</w:t></w:r></w:p></w:tc>'
    first = ('<w:tc><w:tcPr><w:tcW w:w="2268" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Harbour</w:t></w:r></w:p>'
             '<w:p><w:r><w:t>Ines</w:t></w:r></w:p></w:tc>')
    rows = ["".join(_cell(t) for t in REPORT[0]), first + kept + _cell("draft") + _cell("awaits figures")]
    merged = ('<w:tc><w:tcPr><w:tcW w:w="4536" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>'
              '<w:p><w:r><w:t>Harbour</w:t></w:r></w:p><w:p><w:r><w:t>Ines</w:t></w:r></w:p></w:tc>')
    accepted = ["".join(_cell(t) for t in REPORT[0]), merged + _cell("draft") + _cell("awaits figures")]
    assert _drawn(_docx(rows, (2268,) * 4)) == _drawn(_docx(accepted, (2268,) * 4))
