"""An empty table row against what Word drew, offline (ROADMAP.md, "Tables -- measured",
stage 3a: ``w:hideMark``).

``tests/fixtures/hide-mark-observations.json`` holds, for ``tools/make_hide_mark_probe.py``,
Word's text objects, each case's row distances as Word drew them, and every face number
the renderer asked for (``tools/read_hide_mark_probe.py``).  Each document is laid out
here from those numbers alone: every case, every line and every glyph as Word has them.
The rules the probe refuted are held to their scores too.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import make_hide_mark_probe as probe  # noqa: E402
import read_hide_mark_probe as reader  # noqa: E402
import render_record  # noqa: E402

from docx2svg import table as table_model  # noqa: E402
from docx2svg.paginate import is_empty  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
FAMILIES = {"control": [4, 4], "height": [7, 7], "content": [18, 18], "cells": [15, 15], "body": [4, 4]}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_empty_rows_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["word"] == recorded["word"]
    assert result["disagree"] == recorded["disagree"] == []
    assert result["families"] == recorded["families"] == FAMILIES
    assert result["lines"] == recorded["lines"] == [190, 190]
    assert result["glyphs"] == recorded["glyphs"]
    assert result["warnings"] == recorded["warnings"] == []
    assert result["pages"] == recorded["pages"] == len(probe.CASES)


def test_the_four_settings_agree():
    words = [DATA["documents"][name]["word"] for name, _ in reader.documents()]
    assert all(word == words[0] for word in words)


def _every_empty_paragraph(cell, paragraph):
    return bool(cell.properties.get("hideMark")) and is_empty(paragraph)


@pytest.mark.parametrize("rule,families", [
    # w:hideMark ignored (the model before): every empty row a line tall.
    (lambda cell, paragraph: False,
     {"control": [4, 4], "height": [4, 7], "content": [5, 18], "cells": [5, 15], "body": [4, 4]}),
])
def test_refuted_rules_keep_their_scores(rule, families, fonts):
    name, data = reader.documents()[-1]
    with mock.patch.object(table_model, "hides_mark", rule):
        result = reader.score(name, data, fonts, DATA["documents"][name])
    assert result["families"] == families


def test_every_empty_paragraph_of_the_cell_hidden_is_refuted(fonts):
    """Hiding each empty paragraph of a ``w:hideMark`` cell, not just its last: two empty
    paragraphs are one line in Word, not none."""
    name, data = reader.documents()[-1]
    original = table_model.flow_table

    def flow_table(*args, **kwargs):
        flow = original(*args, **kwargs)
        for row in flow.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.hidden = paragraph.hidden or _every_empty_paragraph(cell.cell, paragraph.paragraph)
                cell.stack, cell.after = table_model.cell_stack(cell.paragraphs, flow.previous_style,
                                                                flow.default_style)
                cell.content = table_model.content_height(cell.paragraphs, flow.previous_style, flow.default_style)
            row.height = table_model.row_height(row)
        return flow

    with mock.patch.object(table_model, "flow_table", flow_table):
        result = reader.score(name, data, fonts, DATA["documents"][name])
    assert sorted(key for key in result["disagree"]) == ["content empty2", "content empty2 no height"]


def test_what_empty_means():
    from docx2svg.model import Paragraph, Run

    cell = table_model.Cell(0, 0, 0, 1, (), {"hideMark": True})
    assert table_model.hides_mark(cell, Paragraph())
    assert table_model.hides_mark(cell, Paragraph(runs=(Run(""),)))
    assert not table_model.hides_mark(cell, Paragraph(runs=(Run(" "),)))
    assert not table_model.hides_mark(cell, Paragraph(runs=(Run("", breaks=((0, "tab"),)),)))
    assert not table_model.hides_mark(table_model.Cell(0, 0, 0, 1, (), {}), Paragraph())
