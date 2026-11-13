"""Deleted paragraph marks at a table's edge, drawn as Word's Accept All leaves them
(round 54, lane J; ROADMAP.md, "Revisions -- measured", R.3 and R.4).

``tests/test_revisions.py`` holds every probe case to Word's own export with every
revision accepted; these hold the rules on small made-up documents -- what the parse makes
of each shape, that the layout made from a previous one is the full layout's after edits
around them, and that no line is drawn lower below mode 15.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_revision_probe as probe  # noqa: E402
import probe_docx  # noqa: E402
import wml  # noqa: E402

from docx2svg import ConvertOptions, lay_out_docx, page_svgs  # noqa: E402
from docx2svg.model import Paragraph, Table  # noqa: E402
from docx2svg.parse.document import parse_package  # noqa: E402

para, run, cell, row, table = probe.para, probe.run, probe.cell, probe.row, probe.table


def _docx(body: str, mode: int | None = 15) -> bytes:
    extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),) if mode else ()
    return probe_docx.package(body, final_section=probe.final_section(), styles=probe.styles(), extra_parts=extra)


def _body(body: str) -> list:
    """The final view's body blocks, the paragraph that ends the body left out."""
    blocks = parse_package(_docx(body + para(run("End.")))).body
    assert blocks[-1].text == "End."
    return blocks[:-1]


def _text(block) -> str:
    if isinstance(block, Paragraph):
        return block.text
    return " | ".join(" / ".join(_text(b) for b in c) for r in block.rows for c in r)


def _cell_edge(at: tuple[int, int], rows: int = 2, columns: int = 2, kept: int = 1) -> str:
    out = ""
    for r in range(rows):
        cells = ""
        for c in range(columns):
            inner = para(run(f"r{r + 1}c{c + 1}"))
            if (r, c) == at:
                inner = "".join(para(run(f"Kept {k + 1}.")) for k in range(kept))
                inner += para(run("Last."), mark_rev="del")
            cells += cell(inner)
        out += row(cells)
    return table(out, columns)


def test_a_deleted_mark_before_a_table_joins_the_first_cells_first_paragraph():
    body = _body(para(run("Before. "), props=probe.PA, mark_rev="del")
                 + table(row(cell(para(run("A1"), props={"jc": "right"})) + cell(para(run("B1"))))))
    assert len(body) == 1 and isinstance(body[0], Table)
    first = body[0].rows[0][0][0]
    # Its text first, under the cell paragraph's properties; the cell's paragraph's path.
    assert first.text == "Before. A1"
    assert first.properties.alignment == "right"
    assert first.path == "w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]"
    assert first.joins == (1,)
    assert body[0].joined


def test_a_chain_before_a_table_joins_in_order_and_a_nested_table_is_entered():
    nested = table(row(cell(para(run("Inner"))) + cell(para(run("Side")))))
    body = _body(para(run("One. "), mark_rev="del") + para(run("Two. "), mark_rev="del")
                 + table(row(cell(nested + para(run("Under"))) + cell(para(run("B1"))))))
    assert len(body) == 1
    inner = body[0].rows[0][0][0]
    assert isinstance(inner, Table)
    assert inner.rows[0][0][0].text == "One. Two. Inner"


def test_a_deleted_mark_ending_a_section_before_a_table_stays_where_it_stands():
    # Not measured: kept as the No Markup view draws it, with the next paragraph's properties.
    sect = probe.section()
    body = _body(para(run("Ends a section."), mark_rev="del", sect=sect) + table(row(cell(para(run("A1"))))))
    assert [type(b) for b in body] == [Paragraph, Table]
    assert body[0].text == "Ends a section."


def test_a_cells_last_mark_takes_the_next_cell_in_and_its_other_paragraphs_leave_the_table():
    body = _body(_cell_edge((0, 0)))
    assert [type(b) for b in body] == [Paragraph, Table]
    assert body[0].text == "Kept 1." and body[0].table_style_id is None and body[0].table_conditions == ()
    # The row: one cell, its last paragraph joined to the next cell's; the grid kept (row 2
    # covers the second column), the row ending short.
    assert _text(body[1]) == "Last.r1c2 | r2c1 | r2c2"
    assert [len(r) for r in body[1].rows] == [1, 2]
    assert body[1].grid == (3000, 3000)
    assert body[1].rows[0][0][0].path == "w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[2]"


def test_in_a_later_row_the_table_is_split_where_the_paragraphs_leave_it():
    body = _body(_cell_edge((1, 0)))
    assert [type(b) for b in body] == [Table, Paragraph, Table]
    assert _text(body[0]) == "r1c1 | r1c2" and body[0].grid == (3000, 3000)
    assert body[1].text == "Kept 1."
    # The second part: one row, one cell; the column only the cell taken in covered goes.
    assert _text(body[2]) == "Last.r2c2" and body[2].grid == (3000,)
    assert body[0].path == body[2].path == "w:body/w:tbl[1]"


def test_a_later_cells_paragraphs_stay_in_it():
    body = _body(_cell_edge((0, 1), rows=1, columns=3))
    assert [type(b) for b in body] == [Table]
    assert _text(body[0]) == "r1c1 | Kept 1. / Last.r1c3"
    assert body[0].grid == (3000, 3000)


@pytest.mark.parametrize("at", [(0, 1), (1, 1)], ids=["a row's end", "the table's last cell"])
def test_a_cell_with_no_next_cell_in_its_row_goes_with_its_content(at):
    body = _body(_cell_edge(at))
    assert [type(b) for b in body] == [Table]
    assert "Last." not in _text(body[0]) and "Kept" not in _text(body[0])
    assert sorted(len(r) for r in body[0].rows) == [1, 2]
    assert body[0].grid == (3000, 3000)


def test_no_line_is_drawn_lower_below_mode_15():
    # R.4 was Word's No Markup screen; Accept All draws the joined paragraph as one.
    joined = (para(run("Short alpha. ") + "<w:r><w:br/></w:r>", props=probe.PA, mark_rev="del")
              + para(run(probe.TEXT_B), props={"spacing": {"before": 480, "after": 0}}))
    tops = {}
    for mode in (None, 15):
        layout, _, _ = lay_out_docx(_docx(joined + para(run("After.")), mode))
        tops[mode] = [line.baseline for line in layout.pages[0].lines]
    assert tops[None] == tops[15]


def _layouts(body: str, previous=None):
    data = _docx(body)
    options = ConvertOptions()
    layout, data, fonts = lay_out_docx(data, options, previous=previous)
    full_options = ConvertOptions()
    full, full_data, full_fonts = lay_out_docx(data, full_options, keep=False)
    return layout, page_svgs(layout, data, fonts=fonts), page_svgs(full, full_data, fonts=full_fonts)


def _document(before: str, cell_text: str, kept: str) -> str:
    filler = "".join(para(run(f"Paragraph {k} of the made-up report, with words enough to wrap once on the "
                              "page so that every edit moves the lines after it.")) for k in range(12))
    first = para(run(before), props=probe.PA, mark_rev="del")
    edge = table(row(cell(para(run(kept)) + para(run(cell_text), mark_rev="del")) + cell(para(run("Next cell."))))
                 + row(cell(para(run("Below."))) + cell(para(run("Below right.")))))
    return filler + first + table(row(cell(para(run("First cell."))) + cell(para(run("Side."))))) + filler + edge


def test_edits_around_the_joins_lay_out_as_a_full_layout():
    versions = [("Before the table. ", "Cell last.", "Kept."),
                ("Before the table, now with more words than fit on one line of the cell. ", "Cell last.", "Kept."),
                ("Before the table, now with more words than fit on one line of the cell. ",
                 "Cell last, longer now.", "Kept."),
                ("Short. ", "Cell last, longer now.", "Kept, and grown to wrap within its line of the page "
                                                       "once or twice as it is drawn out of the table."),
                ("Short. ", "Cell last.", "Kept.")]
    previous = None
    for version in versions:
        layout, svgs, full = _layouts(_document(*version), previous)
        assert svgs == full, version
        previous = layout
