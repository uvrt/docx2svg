"""Justified lines in compatibility mode 15 run past the edge by a quarter of each space,
rounded: ``tools/make_justify_probe.py`` against Word, recorded in
``tests/fixtures/justify-observations.json`` (ROADMAP.md, Phase 3.8).

The controls -- mode 15 left-aligned, mode 14 and no ``settings.xml`` justified -- hold no
line past the edge, and neither does the model there.  Every line of the three rounds
breaks where Word breaks it."""

from __future__ import annotations

import functools
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_justify_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
KNOWN: dict[str, int] = {}
FACES = json.loads((REPO / "tests" / "fixtures" / "justify-faces.json").read_text(encoding="utf-8"))


#: Every document of the three rounds, by name (``reader.documents``, which builds them).
NAMES = [f"justify-{name}" for name in (*reader.probe.SETTINGS, *reader.probe.FINE_SETTINGS,
                                        *reader.probe.THIRD_SETTINGS)]


#: Computed on first use, not when the module is collected: every pytest-xdist worker
#: collects every module, and only the one that runs these tests needs them (the module's
#: tests share one worker: tests/conftest.py, ``LAZY_MODULES``).
@functools.cache
def _documents() -> dict[str, tuple]:
    return {name: (data, cases) for name, data, cases in reader.documents()}


@pytest.mark.parametrize("name", NAMES)
def test_justified_lines_break_where_word_breaks_them(name):
    from docx2svg import _render

    data, cases = _documents()[name]

    _, layout, _, _ = _render(data, render_record.options(render_record.RecordedFonts(FACES)))
    ours = reader.model_held(layout, cases)
    word = DATA["documents"][name]
    assert len(ours) == len(word) == len(cases)
    assert sum(a != b for a, b in zip(ours, word)) == KNOWN.get(name, 0)
