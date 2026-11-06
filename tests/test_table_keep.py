"""Table rows under ``w:keepNext`` (round 24): a row whose first cell's first paragraph keeps
with the next goes to the next page with the row after it, a chain of them together, the
last row with the paragraph after the table, and a chain from the first row with the
table -- and with a ``keepNext`` paragraph before it.  Round 36: such a row is not split
at the page's foot, and in mode 15 a chain taller than a page breaks before its table.  ``tools/make_table_keep_probe.py``,
recorded by ``tools/read_table_keep_probe.py`` in
``tests/fixtures/table-keep-observations.json``."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_table_keep_probe as reader  # noqa: E402

from docx2svg import table  # noqa: E402

#: The cases the model does not draw as Word does, none of them a keep: a row holding a
#: nested table across the page's foot, which Word splits between the nested table's rows
#: or after it and the model moves whole (it does not split a nested table).
NOT_MODELLED = {"in a cell plain", "plain nested row"}
#: With no ``settings.xml`` a ``w:cantSplit`` row taller than a page runs off the page's
#: foot whole, in Word and the model alike; Word's export leaves out its lines below the
#: page, which the model's pages list.
UNDRAWN = {("none", "cantSplit tall row")}


def _disagree(compared: dict) -> set:
    return {(setting, name) for setting, per in compared.items() for name, (word, model) in per.items()
            if word != model}


def _known(compared: dict) -> set:
    return {(setting, name) for setting in compared for name in NOT_MODELLED} | UNDRAWN


def test_every_case_puts_each_row_on_the_page_word_does():
    compared = reader.offline()
    assert _disagree(compared) == _known(compared)
    assert reader.tally(compared) == (123, 128)


def _variant(monkeypatch, rule):
    monkeypatch.setattr(table, "keeps_with_next", rule)


def _first_cell_any(row) -> bool:
    return any(p.resolved.get("keepNext") for p in row.cells[0].paragraphs if p.table is None)


def _any_cell_first(row) -> bool:
    return any(cell.paragraphs and cell.paragraphs[0].resolved.get("keepNext") for cell in row.cells)


@pytest.mark.long
@pytest.mark.parametrize("what,rule,expected", [
    ("rows never keep", lambda row: False, {
        "all", "but last", "row 5", "rows 4-5", "row 5 cell a", "row 5 first", "foot last row", "foot all",
        "long all", "long but last", "long header all", "before all", "cantSplit row 5", "header all",
        "nested inner", "row 5 a first only", "row 5 style", "row 5 wrapped, cantSplit", "all, page break after",
        "kept row 6 of 4", "kept row 6 of 8", "kept row 6 of 4, all", "kept last row", "kept rows of 4",
        "kept tall row", "kept tall row at the top", "kept nested row", "header, kept row 6", "long kept rows of 4",
        "long kept rows of 4, header", "chain head par table", "chain head par table at the top",
        "chain head table", "chain table par", "chain two tables"}),
    ("any paragraph of the first cell", _first_cell_any, {"row 5 last", "nested inner", "nested outer",
                                                          "kept nested row"}),
    ("the first paragraph of any cell", _any_cell_first, {"row 5 cell b", "row 5 b first only", "nested inner",
                                                          "nested outer", "kept nested row"}),
])
def test_refuted_rules_disagree_where_word_shows_them_wrong(monkeypatch, what, rule, expected):
    _variant(monkeypatch, rule)
    compared = reader.offline()
    found = {name for _, name in _disagree(compared) - _known(compared)}
    assert found == expected, what


@pytest.mark.long
def test_a_kept_row_is_not_split_at_the_foot(monkeypatch):
    """Round 36: split as any row, a kept row disagrees with Word wherever one meets the
    page's foot -- in both settings, but in mode 15 also in a chain from the page's
    start (``long kept rows of 4``), where below mode 15 Word splits it."""
    place = table._place_table
    monkeypatch.setattr(table, "_place_table", lambda *a, **k: place(*a, **{**k, "split_kept": True}))
    compared = reader.offline()
    every = {"kept row 6 of 4", "kept row 6 of 8", "kept row 6 of 4, all", "kept last row", "kept tall row",
             "kept tall row at the top", "kept nested row", "header, kept row 6", "long kept rows of 4, header"}
    assert _disagree(compared) - _known(compared) == (
        {(setting, name) for setting in compared for name in every} | {("15", "long kept rows of 4")})


@pytest.mark.long
def test_below_mode_15_a_kept_row_in_a_chain_from_the_pages_start_splits(monkeypatch):
    from docx2svg import paginate

    monkeypatch.setattr(paginate, "_kept_row_breaks", lambda *args: False)
    compared = reader.offline()
    assert _disagree(compared) - _known(compared) == {("none", "long kept rows of 4"),
                                                      ("none", "long kept rows of 4, header")}


def test_the_documents_rule_is_the_layouts(monkeypatch):
    """:func:`docx2svg.table.lead_paragraph` (a document's rows, before any layout: what
    docx-agent reads) finds, for every row of every case, the paragraph the layout's
    :func:`~docx2svg.table.keeps_with_next` decides by."""
    real = table.keeps_with_next
    seen = []

    def both(row):
        flow = table._lead(row, table._flow_first, table._flow_nested)
        model = table.lead_paragraph(tuple(cell.blocks for cell in row.row.cells))
        seen.append(model is (flow.paragraph if flow is not None else None))
        return real(row)

    monkeypatch.setattr(table, "keeps_with_next", both)
    reader.offline()
    assert len(seen) > 100 and all(seen)


def _rows_of(body: str):
    import make_table_keep_probe as probe
    import probe_docx

    from docx2svg import parse_package

    document = parse_package(probe_docx.package(body, styles=probe.styles()))
    (found,) = [block for block in document.body if hasattr(block, "rows")]
    return document, found.rows


@pytest.mark.parametrize("name,keep,expected", [
    ("first cell only", (True, False), True),
    ("every cell", (True, True), True),
    ("second cell only", (False, True), False),
    ("first cell, its second paragraph only", ((False, True), (False, False)), False),
    ("none", (False, False), False),
])
def test_a_documents_row_keeps_by_its_first_cell(name, keep, expected):
    import make_table_keep_probe as probe

    lines = 2 if isinstance(keep[0], tuple) else 1
    document, rows = _rows_of(probe._two(probe._row("R0", keep, lines=lines) + probe._row("R1")))
    assert table.row_keeps_with_next(document, rows[0]) is expected, name
    assert table.row_keeps_with_next(document, rows[1]) is False


def test_a_documents_row_keeps_by_its_style_and_by_a_nested_tables_first_paragraph():
    import make_table_keep_probe as probe

    styled = probe._row("R0").replace("<w:pPr>", '<w:pPr><w:pStyle w:val="Kept"/>', 1)
    inner = f"<w:tr>{probe._cell(probe._nested('N1', True) + probe._p('R1a'))}{probe._cell(probe._p('R1b'))}</w:tr>"
    outer = (f"<w:tr>{probe._cell(probe._nested('N2', False) + probe._p('R2a', True))}"
             f"{probe._cell(probe._p('R2b', True))}</w:tr>")
    document, rows = _rows_of(probe._two(styled + inner + outer + probe._row("R3")))
    assert [table.row_keeps_with_next(document, row) for row in rows] == [True, True, False, False]
    assert table.lead_paragraph(rows[1]).path.endswith("w:tc[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]")
    assert table.lead_paragraph(()) is None and table.lead_paragraph(((),)) is None
