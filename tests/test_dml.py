"""DrawingML's paint against what Word drew, offline (ROADMAP.md, "DrawingML drawn by the
shared renderers").

``tests/fixtures/dml-observations.json`` holds, for ``tools/make_dml_probe.py``, Word's
fill colours, gradient shadings, strokes and pattern tiles, and every face number the
renderer asked for (``tools/read_dml_probe.py``).  The document is laid out here from
those numbers alone and its paint compared with Word's.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_dml_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))

#: What the comparison is known to miss, and why: three dashed lines whose dashes the
#: reader cannot take apart in Word's PDF -- a square-capped long dash-dot whose filled
#: dashes the extraction merges, and two round-dotted ellipses (one floating, one
#: inline), which are not straight lines.
KNOWN = {"dml-15": {"dash": 3}}


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_drawingml_paint_against_word(name, data):
    recorded = DATA["documents"][name]
    fonts = render_record.RecordedFonts(DATA["faces"])
    result = reader.score(data, fonts, recorded["word"])
    assert result["scores"] == recorded["scores"]
    for family, (agree, compared) in result["scores"].items():
        assert compared - agree == KNOWN.get(name, {}).get(family, 0), (family, result["problems"])
