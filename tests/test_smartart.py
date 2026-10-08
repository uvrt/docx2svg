"""SmartArt against what Word drew, offline (ROADMAP.md, "Floating drawings -- measured",
F.21): ``tools/make_smartart_probe.py``'s diagrams, each carrying the drawing Word caches
for its data, recorded by ``tools/read_smartart_probe.py`` in
``tests/fixtures/smartart-observations.json`` with that cache -- and drawn here from the
cache by ``docx2svg.diagram``: every shape, outline and line of text where Word drew it.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_smartart_probe  # noqa: E402
import read_smartart_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_smartart_against_word(name, data):
    recorded = DATA["documents"][name]
    fonts = render_record.RecordedFonts(DATA["faces"])
    result = reader.score(data, fonts, recorded["word"])
    assert result["scores"] == recorded["scores"]
    for family, (agree, compared) in result["scores"].items():
        assert agree == compared, (family, result["problems"])


def test_the_probe_carries_the_cache_it_was_recorded_with():
    assert make_smartart_probe.CACHES and len(make_smartart_probe.CACHES) == len(make_smartart_probe.CASES)
    assert {int(k): v for k, v in DATA["caches"].items()} == make_smartart_probe.CACHES


def test_a_diagram_with_no_cached_drawing_is_a_placeholder_and_a_warning():
    from docx2svg import ConvertOptions, convert_docx_to_layout

    saved = make_smartart_probe.CACHES
    make_smartart_probe.CACHES = {}
    try:
        data = make_smartart_probe.build("15")
    finally:
        make_smartart_probe.CACHES = saved
    options = ConvertOptions(pages=[1])
    layout = convert_docx_to_layout(data, options)
    kinds = [p.kind for placed in layout.pages[0].floats for p in placed.primitives]
    assert kinds == ["placeholder"]
    assert any("diagram-no-cached-drawing" in str(w) for w in options.warnings)
