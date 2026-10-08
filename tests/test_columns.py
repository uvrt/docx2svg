"""Text columns against what Word drew, offline (ROADMAP.md, "Phase 4 -- measured", 4.14).

``tests/fixtures/columns-observations.json`` holds, for ``tools/make_columns_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_columns_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_columns_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]
#: A footnote in a section a continuous break joins to another, an endnote in columns in
#: mode 15: measured, not laid out -- the layout stops at the first.
STOPPED = {name for name, _ in DOCUMENTS if name.startswith("cols-notesregion-")} | {"cols-endnotes-15"}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


def test_every_document_is_recorded():
    assert sorted(DATA["documents"]) == sorted(name for name, _ in reader.documents())


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_columns_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["rules"] == recorded["rules"]
    assert result["pages"] == recorded["pages"]
    assert result["warnings"] == recorded["warnings"]
    assert result.get("pictures") == recorded.get("pictures")
    if name in STOPPED:
        assert result["warnings"] == ["layout-stopped:columns"]
        return
    if result["pictures"] is not None:
        # Every floating picture where Word draws it.
        assert result["pictures"][0] == result["pictures"][1]
    # Every line, every separator and table border where Word drew it, on Word's pages.
    assert result["lines"][0] == result["lines"][1]
    assert result["rules"] == [0, 0]
    assert result["pages"][0] == result["pages"][1]
    assert result["result"].extra == 0
    assert result["result"].not_drawn == 0
    assert result["warnings"] == []


def _document(name: str):
    import make_columns_probe as probe

    from docx2svg import parse_package

    return parse_package(probe.build(name))


def test_equal_columns_share_the_width_in_whole_twips():
    from docx2svg.linebreak import column_boxes

    sections = _document("cols-geometry-15").sections
    # 9,164 twips of text: two columns 355 apart are 4,404 wide (4,404.5), five 250
    # apart 1,632 (1,632.8); each starts where the one before ends plus the space.
    assert column_boxes(sections[1]) == ((0, 4404, 355), (4759, 4404, 355))
    assert [box[0] for box in column_boxes(sections[5])] == [0, 1882, 3764, 5646, 7528]
    # Unequal columns as stated, whatever their sum.
    assert column_boxes(sections[7]) == ((0, 1501, 1033), (2534, 2999, 0))
    assert sections[0].column_separator is False
    assert _document("cols-sep-15").sections[0].column_separator is True


def test_a_column_edge_on_a_half_pixel_rounds_by_the_mode():
    """3,324 twips from the page's edge is 692.5 px: 693 below mode 15, 692 in it (the
    length held in whole layout units first)."""
    from docx2svg import _render

    for name, expected in (("cols-geometry-14", 693), ("cols-geometry-none", 693), ("cols-geometry-15", 692)):
        _documents, layout, _data, _fonts = _render(dict(DOCUMENTS)[name],
                                                    render_record.options(render_record.RecordedFonts(DATA["faces"])))
        starts = sorted({round(float(line.spans[0].xs[0])) for line in layout.pages[5].lines if line.spans})
        assert expected in starts, (name, starts)


def _rules(name: str, kind: str) -> list[tuple[int, int, int]]:
    from docx2svg import _render

    _documents, layout, _data, _fonts = _render(dict(DOCUMENTS)[name],
                                                render_record.options(render_record.RecordedFonts(DATA["faces"])))
    return [(round(float(rule.x)), round(float(rule.x + rule.width)), round(float(rule.y)))
            for page in layout.pages for rule in page.rules if rule.kind == kind]


def test_a_column_s_footnote_separator_is_cut_at_its_edge_below_mode_15():
    """Below mode 15 each column's notes stand under their own separator, 2,880 twips
    long or cut at the column's right edge (2,787 twips: 581 px); in mode 15 one
    separator over the first column, 2,880 twips (600 px) whatever its width."""
    for name in ("cols-notes3-14", "cols-notes3-none"):
        assert _rules(name, "separator") == [(300, 881, 2947), (965, 1546, 2947), (1629, 2210, 2998)]
    assert _rules("cols-notes3-15", "separator") == [(300, 900, 2947)]
    assert _rules("cols-notesunequal-15", "separator") == [(300, 900, 2896)]


def test_notes_a_column_cannot_hold_go_on_under_the_next_column_s_continuation_separator():
    """``notesover`` below mode 15: a note of 21 lines, 14 at the first column's foot, the
    rest at the second's under the continuation separator, to that column's right edge."""
    assert _rules("cols-notesover-14", "separator") == [(300, 900, 2439), (1330, 2210, 2744)]


def test_a_mode_15_page_keeps_its_note_area_free_in_every_column():
    """``notesright`` in mode 15: the second column's notes, balanced over both columns,
    end the first column six lines short of its 51 (below mode 15 it holds them all)."""
    def last_left(name):
        objects = DATA["documents"][name]["objects"]
        return max(o[1] for o in objects if o[5] == 3000000 and o[3] == 46.0)

    assert last_left("cols-notesright-15") == 2818.0
    assert last_left("cols-notesright-14") == 3154.0
    for name in ("cols-notesright-15", "cols-notesright-14"):
        assert DATA["documents"][name]["lines"][0] == DATA["documents"][name]["lines"][1]


def test_every_line_says_which_text_column_it_is_in():
    """:attr:`Line.column`: the 0-based column of a section of several, a note's line the
    column it stands under; ``None`` in a section of one column.  On every page, each
    column's lines start right of the column before's, and every page of several columns
    has lines in its first."""
    from docx2svg import convert_docx_to_layout

    fonts = render_record.RecordedFonts(DATA["faces"])
    seen = set()
    for name, data in DOCUMENTS:
        layout = convert_docx_to_layout(data, render_record.options(fonts))
        for page in layout.pages:
            lefts: dict = {}
            for line in page.lines:
                seen.add(line.column)
                if line.story is not None:
                    assert line.column is None
                if line.spans and line.spans[0].xs:
                    lefts.setdefault(line.column, []).append(line.spans[0].xs[0])
            columns = sorted(c for c in lefts if c is not None)
            for before, after in zip(columns, columns[1:]):
                assert min(lefts[after]) > min(lefts[before]), (name, page.number, after)
    assert seen >= {None, 0, 1, 2, 3, 4}
    regions = convert_docx_to_layout(dict(DOCUMENTS)["cols-regions-15"], render_record.options(fonts))
    assert {line.column for line in regions.pages[0].lines} == {None, 0, 1}

