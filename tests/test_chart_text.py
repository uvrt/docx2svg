"""Chart titles, axis titles and short charts against what Word drew, offline (ROADMAP.md,
F.23): ``tools/make_chart_text_probe.py``'s 80 charts, recorded by
``tools/read_chart_text_probe.py`` in ``tests/fixtures/chart-text-observations.json``, and
drawn here by ``ooxml-common``'s chart layout under its ``WORD`` rules.

What is left is pinned by case: Word's own "Chart Title", which it draws in its interface's
language (three texts: two series, and ``c:autoTitleDeleted`` ``1`` over a ``c:title`` with
no text); a value axis' title stated unturned (``rot="0"``), which is not laid out (twelve
texts, eighteen fills and 38 strokes -- the whole plot moves); a chart style other than
Word 365's (``c14:style`` 101, twenty fills: its series' colours); and a plot of no height,
where Word draws its top gridline over its axis (one stroke in each of three charts).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_chart_text_probe as probe  # noqa: E402
import read_chart_text_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
#: What the comparison is known to miss, per case (module docstring).
KNOWN = {21: (1, 0, 0), 22: (1, 0, 0), 23: (1, 0, 0), 37: (12, 18, 38), 49: (0, 0, 1), 61: (0, 0, 1),
         71: (0, 20, 0), 76: (0, 0, 1)}


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_chart_titles_and_short_charts_against_word(name, data):
    recorded = DATA["documents"][name]
    fonts = render_record.RecordedFonts(DATA["faces"])
    result = reader.score(data, fonts, recorded["word"])
    assert result["scores"] == recorded["scores"]
    for index, found in enumerate(result["cases"]):
        missed = tuple(compared - agree for agree, compared in found.values())
        assert missed == KNOWN.get(index, (0, 0, 0)), (index, probe.CASES[index].note,
                                                         result["problems"].get(index))


def _one(case: probe.Case) -> bytes:
    probe.CASES, saved = (case,), probe.CASES
    try:
        return probe.build("15")
    finally:
        probe.CASES = saved


def _render(data: bytes):
    from docx2svg import ConvertOptions, convert_docx_to_layout

    options = ConvertOptions()
    layout = convert_docx_to_layout(data, options)
    texts = [p.markup for placed in layout.pages[0].floats for p in placed.primitives
             if p.kind == "markup" and p.what == "chart"]
    return texts, [str(warning) for warning in options.warnings]


@pytest.mark.faces
def test_word_s_own_titles_are_said_and_take_their_band():
    texts, warnings = _render(_one(probe.CASES[21]))
    assert any("chart-title-not-drawn" in w for w in warnings), warnings
    texts, warnings = _render(_one(probe.CASES[17]))
    assert ">Plan<" in texts[0] and not [w for w in warnings if "chart" in w], warnings


@pytest.mark.faces
def test_an_axis_title_drawn_turned_and_one_not_laid_out_is_said():
    texts, warnings = _render(_one(probe.CASES[28]))
    assert "Revenue" in texts[0] and "rotate(-90" in texts[0] and not [w for w in warnings if "chart" in w]
    texts, warnings = _render(_one(probe.CASES[37]))
    assert "Revenue" not in texts[0]
    assert any("chart-axis-title-not-drawn" in w and "turned or placed otherwise" in w for w in warnings), warnings
    default = probe.chart(val_title=probe.titled(None))
    texts, warnings = _render(_one(probe.Case("axis", "no text", default)))
    assert any("chart-axis-title-not-drawn" in w and "Word's own words" in w for w in warnings), warnings
