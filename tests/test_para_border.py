"""Paragraph bottom borders against Word: ``tools/make_para_border_probe.py``, recorded in
``tests/fixtures/para-border-observations.json``.  Not settled (ROADMAP.md, Phase 5.15):
this holds the model's count, so that a rule found later is seen to raise it."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_para_border_probe as probe  # noqa: E402
import read_para_border_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
#: Bottom borders on Word's row, of 385 a document.
AGREE = {"none": 234, "15": 234}


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_bottom_borders_against_word(setting):
    from docx2svg import _render

    fonts = render_record.RecordedFonts(DATA["faces"])
    _, layout, _, _ = _render(probe.build(setting), render_record.options(fonts))
    ours = reader.model_rows(layout)
    word = DATA["documents"][f"para-border-{setting}"]
    assert len(ours) == len(word) == len(probe.CASES)
    assert sum(w is not None and w[0] == o[1] for o, w in zip(ours, word)) == AGREE[setting]
