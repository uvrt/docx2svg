"""Every baseline of whole documents, predicted from *resolved* properties, without Word.

``tests/fixtures/baseline-observations.json`` holds the lines Word drew (page, baseline,
text) and the four metric integers of every face involved, recorded by
``tools/read_baselines.py --record``.  The documents are the committed fixtures:
``style-document.docx`` (generated: nothing is stated directly, everything inherits),
``layout-sweep.docx`` (Phase 0), the four third-party ``samplelib`` documents and the
five ``wordto`` ones.  The four ``filesamples`` documents may not be committed; they are
in ``tests/test_filesamples.py``, which skips where ``scratch/`` does not hold them.

The scores are pinned, so a change that moves one is seen.  Every in-scope baseline of
every document is exact: since the line box (ROADMAP.md, "The line box") for the earlier
documents, and since the page-top rule (ROADMAP.md, "The page top") for ``wordto``, whose
first pages were all 100 px (24 pt) low before it.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(FIXTURES.parent.parent / "tools"))

import baselines  # noqa: E402

from docx2svg import parse_package  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

DATA = json.loads((FIXTURES / "baseline-observations.json").read_text())


def metrics(face, bold=False, italic=False):
    found = DATA["faces"].get(f"{face}|{int(bold)}|{int(italic)}")
    return None if found is None else FaceMetrics(*found)


def _path(name: str) -> Path:
    for directory in (FIXTURES, FIXTURES / "samplelib", FIXTURES / "wordto"):
        if (directory / name).exists():
            return directory / name
    raise FileNotFoundError(name)


def _score(name: str, **options):
    data = _path(name).read_bytes()
    document = parse_package(data)
    drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
    return baselines.predict(document, drawn, metrics, package=data, **options)


#: (exact, scored, out of scope).  Out of scope: tables, drawing paragraphs, and the
#: lines below them on the same page.
EXPECTED = {
    "layout-sweep.docx": (91, 91, 0),
    "style-document.docx": (34, 34, 0),
    "sample-resume.docx": (20, 20, 0),
    "sample-simple.docx": (22, 22, 7),
    "sample-long.docx": (453, 453, 35),
    "sample-blank.docx": (0, 0, 0),
    # wordto (ROADMAP.md, "Nine more real documents").  Each opens with a Heading 1
    # whose 24 pt space before Word keeps at the top of the first page: the first
    # paragraph of a section keeps it (ROADMAP.md, "The page top").  Pages 2 and 3 of
    # sample-10pages start after natural breaks, which drop it.
    "sample-1page.docx": (7, 7, 0),
    "sample-5pages.docx": (26, 26, 30),
    "sample-10pages.docx": (70, 70, 19),
    "sample-with-images.docx": (8, 8, 9),
    "sample-with-table.docx": (5, 5, 39),
}

#: Glyphs the checker loses, all in table cells (out of scope): a row whose first-column
#: cell is Calibri Bold is drawn with that cell's baseline 1 px below the Cambria cells,
#: so the row arrives as two lines in the wrong order and the walk loses its place.
#: Not cascade errors -- the same cells are drawn in the resolved face where the row
#: shares one baseline.
GLYPHS_LOST_IN_TABLE_ROWS = {
    "sample-5pages.docx": 8,
    "sample-10pages.docx": 2,
    "sample-with-table.docx": 2,
}


@pytest.mark.parametrize("name", sorted(EXPECTED))
def test_baselines_from_the_cascade(name):
    results = _score(name)
    assert baselines.summary(results) == EXPECTED[name]
    assert not [r for r in results if r.status == "miss"]


@pytest.mark.parametrize("name", sorted(EXPECTED))
def test_every_glyph_is_drawn_in_the_face_and_size_the_cascade_resolves(name):
    """31,057 glyphs across the six documents, compared one by one: family, weight,
    slant and drawn size.  The third-party documents' table header cells are bold only
    through the table style's conditional formats (``w:tblStylePr``)."""
    document = parse_package(_path(name).read_bytes())
    drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
    compared, problems = baselines.check_glyphs(document, drawn, metrics)
    assert len(problems) == GLYPHS_LOST_IN_TABLE_ROWS.get(name, 0), problems
    assert [compared, len(problems)] == DATA["glyphs"][name]


def test_the_generated_document_states_nothing_directly():
    document = parse_package((FIXTURES / "style-document.docx").read_bytes())
    for paragraph in document.paragraphs:
        declared = paragraph.properties.declared if paragraph.properties else {}
        assert not {k for k in declared if not k.startswith("numPr")}, paragraph.text
        for run in paragraph.runs:
            assert not (run.properties.declared if run.properties else {}), run.text


def test_the_committed_document_is_what_the_generator_writes():
    import io
    import zipfile

    import make_style_document

    committed = zipfile.ZipFile(FIXTURES / "style-document.docx")
    regenerated = zipfile.ZipFile(io.BytesIO(make_style_document.build()))
    assert committed.namelist() == regenerated.namelist()
    for name in committed.namelist():
        assert committed.read(name) == regenerated.read(name), name


def test_a_raised_run_is_folded_into_its_line_and_a_table_row_is_not():
    """``merge_raised``: a superscript drawn smaller on its own baseline rejoins its line;
    two same-size groups 1 px apart (a table row whose cells round differently) do not."""
    from quartz_pdf import Line, TextRun

    def run(x, y, size, text):
        return TextRun(0, "F", size, x, y, text)

    lines = [
        Line(0, 819.0, [run(300, 819.0, 33.0, "script")]),
        Line(0, 836.0, [run(0, 836.0, 50.0, "have a super"), run(400, 836.0, 50.0, " and")]),
        Line(0, 2700.0, [run(500, 2700.0, 46.0, ".docx")]),
        Line(0, 2701.0, [run(0, 2701.0, 46.0, "Word Document")]),
    ]
    merged = baselines.merge_raised(lines)
    assert [(line.y, "".join(r.text for r in line.runs)) for line in merged] == [
        (836.0, "have a superscript and"), (2700.0, ".docx"), (2701.0, "Word Document"),
    ]


def test_the_refuted_spacing_rules_really_are_refuted():
    for name in ("sample-long.docx", "style-document.docx"):
        document = parse_package(_path(name).read_bytes())
        drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
        best = baselines.summary(baselines.predict(document, drawn, metrics))[0]
        added = baselines.summary(baselines.predict(document, drawn, metrics, collapse="sum"))[0]
        kept = baselines.summary(baselines.predict(document, drawn, metrics, before_at_page_top=True))[0]
        outside = baselines.summary(baselines.predict(document, drawn, metrics, line_box=False))[0]
        assert added < best and kept < best and outside < best


def test_each_line_holds_the_characters_word_put_on_it():
    """``line_shares``: every multi-line paragraph of every document, split by the drawn
    lines' text, gives each line exactly what it drew (whitespace and case aside), and
    every character to some line."""
    checked = 0
    for name in DATA["documents"]:
        data = _path(name).read_bytes()
        document = parse_package(data)
        drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
        lines_of: dict = {}
        for block, line in baselines.match_lines(document, baselines.blocks(document, data), drawn):
            lines_of.setdefault(block, []).append(line)
        all_blocks = baselines.blocks(document, data)
        for block, lines in lines_of.items():
            if block < 0 or len(lines) < 2 or all_blocks[block].excluded:
                continue
            paragraph = all_blocks[block].paragraph
            shares = baselines.line_shares(document, paragraph, [line.text for line in lines])
            texts = [baselines._norm("".join(char for _, char in share)) for share in shares]
            assert "".join(texts) == baselines._norm(baselines._visible_text(document, paragraph))
            for text, line in zip(texts, lines):
                assert baselines._norm(line.text).endswith(text), (name, line.text, text)
            checked += 1
    assert checked > 100
