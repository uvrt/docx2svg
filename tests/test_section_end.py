"""An empty paragraph that holds a section break stays on its page and takes no room
(ROADMAP.md, "Phase 4 -- measured", 4.11): ``tools/make_section_end_probe.py``, recorded by
``tools/read_section_end_probe.py`` in ``tests/fixtures/section-end-observations.json``.
The model reads every one of the 192 cases as Word laid them out in every setting (91 and
92 before), and makes every page Word made (464 against Word's 409 and 410 before).  Seven
cases per setting were recorded as off until the reader stopped folding the line after an
empty 48 pt paragraph at a page's top into the paragraph's mark (Word draws the mark as a
space): read apart, Word stands that paragraph its full 244 px line, as the model does."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_section_end_probe as probe  # noqa: E402
import read_section_end_probe as reader  # noqa: E402
import render_record  # noqa: E402

from docx2svg import paginate  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
FONTS = render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_a_section_break_paragraph_against_word(setting):
    recorded = DATA["documents"][setting]
    lines, pages = reader.model_lines(probe.build(setting), FONTS)
    model = reader.outcomes(lines)
    disagree = sorted(probe.CASES[n].key for n, outcome in model.items()
                      if outcome != recorded["word"][probe.CASES[n].key])
    assert disagree == recorded["disagree"] == []
    assert pages == recorded["pages"]


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_the_document_s_last_paragraph_still_goes_to_the_next_page(setting):
    for case in probe.END_CASES:
        _, pages = reader.model_lines(probe.build_end(case, setting), FONTS)
        assert pages == DATA["documents"][setting]["ends"][case.key], case.key


def test_only_an_empty_paragraph_holding_a_section_break_is_a_section_mark():
    from docx2svg.model import Paragraph, Run

    assert paginate.is_empty(Paragraph())
    assert not paginate.is_empty(Paragraph(runs=(Run("x"),)))
    assert not paginate.is_empty(Paragraph(runs=(Run("", breaks=((0, "tab"),)),)))
    # In mode 15 straight after a table the mark takes its room; otherwise none.
    table = paginate.TableItem(0, 0, None)
    assert paginate.section_mark_room(table, 15)
    assert not paginate.section_mark_room(table, 14)
    assert not paginate.section_mark_room(None, 15)
