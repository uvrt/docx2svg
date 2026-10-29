"""Table rows under ``w:keepNext`` (round 24): a row whose first cell's first paragraph keeps
with the next goes to the next page with the row after it, a chain of them together, the
last row with the paragraph after the table, and a chain from the first row with the
table -- and with a ``keepNext`` paragraph before it.  ``tools/make_table_keep_probe.py``,
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

#: The one case the model does not draw as Word does: a row holding a nested table across
#: the page's foot, which Word splits between the nested table's rows and the model moves
#: whole (it does not split a nested table) -- not a keep.
NOT_MODELLED = {"in a cell plain"}


def _disagree(compared: dict) -> set:
    return {(setting, name) for setting, per in compared.items() for name, (word, model) in per.items()
            if word != model}


def test_every_case_puts_each_row_on_the_page_word_does():
    compared = reader.offline()
    assert _disagree(compared) == {(setting, name) for setting in compared for name in NOT_MODELLED}
    assert reader.tally(compared) == (66, 68)


def _variant(monkeypatch, rule):
    monkeypatch.setattr(table, "keeps_with_next", rule)


def _first_cell_any(row) -> bool:
    return any(p.resolved.get("keepNext") for p in row.cells[0].paragraphs if p.table is None)


def _any_cell_first(row) -> bool:
    return any(cell.paragraphs and cell.paragraphs[0].resolved.get("keepNext") for cell in row.cells)


@pytest.mark.parametrize("what,rule,expected", [
    ("rows never keep", lambda row: False, {
        "all", "but last", "row 5", "rows 4-5", "row 5 cell a", "row 5 first", "foot last row", "foot all",
        "long all", "long but last", "long header all", "before all", "cantSplit row 5", "header all",
        "nested inner", "row 5 a first only", "row 5 style", "row 5 wrapped, cantSplit"}),
    ("any paragraph of the first cell", _first_cell_any, {"row 5 last", "nested inner", "nested outer"}),
    ("the first paragraph of any cell", _any_cell_first, {"row 5 cell b", "row 5 b first only", "nested inner",
                                                          "nested outer"}),
])
def test_refuted_rules_disagree_where_word_shows_them_wrong(monkeypatch, what, rule, expected):
    _variant(monkeypatch, rule)
    found = {name for _, name in _disagree(reader.offline())} - NOT_MODELLED
    assert found == expected, what
