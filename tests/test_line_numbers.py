"""Line numbers (``w:lnNumType``) drawn as Word draws them (ROADMAP.md, "Phase 4 --
measured", 4.18).  ``tools/make_line_number_probe.py``, recorded by
``tools/read_line_number_probe.py`` in ``tests/fixtures/line-number-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_line_number_probe as reader  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()


def _score(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    return read_render.score(name, data, fonts, DATA["documents"][name]["objects"], [])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_line_number_where_word_draws_it(name, data):
    result, _rects, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == []
    assert result.matched == result.word == result.model == 6255
    assert result.x_agree == result.y_agree == result.face_agree == result.size_agree == result.matched


@pytest.mark.parametrize("index,expected", [
    (0, (5885, 5580, 143, 370)),  # before: no line numbers drawn
    (1, (5801, 5337, 450, 454)),  # the count starts at w:start, not one past it
    (2, (6251, 6251, 3, 4)),  # w:restart continuous restarts with the section
])
def test_refuted_rules_keep_their_scores(index, expected):
    name, data = DOCUMENTS[1]
    _what, variant = reader.refuted()[index]
    with variant():
        result, _rects, _warnings = _score(name, data)
    assert (result.matched, result.y_agree, result.extra, result.not_drawn) == expected


def test_line_numbers_are_spans_of_their_own_with_the_section_s_path():
    from docx2svg import convert_docx_to_layout

    from docx2svg.paths import resolve_path
    from docx2svg.xmlutil import parse_xml
    import zipfile, io

    data = DOCUMENTS[0][1]
    layout = convert_docx_to_layout(data)
    spans = [span for page in layout.pages for line in page.lines for span in line.spans if span.kind == "line-number"]
    assert spans and all(span.path.endswith("/w:lnNumType") for span in spans)
    root = parse_xml(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))
    assert {resolve_path(root, span.path).tag.rsplit("}", 1)[1] for span in spans} == {"lnNumType"}
