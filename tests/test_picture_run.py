"""A picture's line in mode 15 is at least its run's text height, against Word
(``tools/make_picture_run_probe.py``, recorded in
``tests/fixtures/picture-run-observations.json``; ROADMAP.md, Phase 5.16)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_picture_run_probe as probe  # noqa: E402
import read_picture_run_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_the_line_after_each_picture_on_words_baseline(setting):
    from docx2svg import _render

    fonts = render_record.RecordedFonts(DATA["faces"])
    _, layout, _, _ = _render(probe.build(setting), render_record.options(fonts))
    assert reader.model_after(layout) == DATA["documents"][f"picture-run-{setting}"]
