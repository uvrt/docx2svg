"""List numbering against Word, offline (ROADMAP.md, "Numbering and lists -- measured").

``tests/fixtures/numbering-observations.json`` holds, for ``tools/make_numbering_probe.py``,
Word's text objects, each case's lines as Word drew them (the list labels among them) and
every face number the renderer asked for (``tools/read_numbering_probe.py``).  Each
document is laid out here from those numbers alone: every case, every line and every
glyph as Word has them.  The counting rules the probe refuted are held to their scores.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import make_numbering_probe as probe  # noqa: E402
import read_numbering_probe as reader  # noqa: E402
import render_record  # noqa: E402

from docx2svg import linebreak  # noqa: E402
from docx2svg.parse import document as parse_document  # noqa: E402
from docx2svg.parse import styles as parse_styles  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
FAMILIES = {"override": [15, 15], "instance": [8, 8], "restart": [8, 8], "legal": [5, 5], "story": [7, 7]}
#: What Word drew, case by case: the list lines, ``label+text`` (``7.B0``: instance B, level 0).
LABELS = {
    "override start 7": "7.B0 8.B0 9.B0",
    "override after another instance": "1.A0 2.A0 5.B0 6.B0 7.A0 8.A0",
    "override restart at 1 after another": "1.A0 2.A0 3.A0 1.B0 2.B0 3.A0",
    "override used again after another": "5.B0 6.A0 7.B0 8.A0 9.B0",
    "override two overridden, interleaved": "5.B0 10.C0 11.B0 12.C0",
    "override at level 1": "1.B0 c.B1 d.B1 2.B0 c.B1 d.B1",
    "override at level 1, then an instance without": "1.B0 c.B1 d.B1 2.A0 a.A1 b.B1",
    "override first item deeper than the override": "a.B1 5.B0 a.B1 6.B0",
    "override a whole level": "1.A0 2.A0 III.B0 IV.B0 5.A0",
    "override a whole level and a start": "1.A0 2.A0 IV.B0 V.B0 6.A0",
    "instance two, no override": "1.A0 2.A0 3.B0 4.B0 5.A0",
    "instance two definitions alike": "1.A0 2.A0 1.B0 2.B0 3.A0",
    "instance one nsid, other levels": "1.A0 2.A0 3.B0 4.B0 5.A0",
    "instance numStyleLink": "1.A0 2.A0 1.B0 2.B0 3.A0",
    "restart lvlRestart 0": "1.A0 a.A1 b.A1 2.A0 c.A1",
    "restart lvlRestart 1 at level 2": "1.A0 a.A1 i.A2 ii.A2 b.A1 iii.A2 2.A0 a.A1 i.A2",
    "restart a level skipped": "1.A0 2.A0 2.1.1.A2 2.1.2.A2 2.2.A1 2.2.1.A2",
    "restart first item at level 1, no %1": "a.A1 b.A1 2.A0",
    "legal roman parent, isLgl": "I.A0 1.1.A1 II.A0 2.1.A1",
    "legal letters, isLgl on a letter level": "a.A0 1.1.A1 1.2.A1",
    "legal isLgl on level 0": "1.A0 2.A0 II.a.A1",
    "story two footnotes": "1.A0 2 Noted 2.A0 3 Noted 3.A0 2 Note 1.A0 2.A0 3 Note 3.A0",
    "story two text boxes": "1.A0 Anchor 2.A0 Anchor 3.A0 1.A0 2.A0 3.A0",
    "story header": "1.A0 2.A0 1.A0 2.A0",
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_the_labels_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["word"] == recorded["word"]
    assert result["disagree"] == recorded["disagree"] == []
    assert result["families"] == recorded["families"] == FAMILIES
    assert result["lines"] == recorded["lines"] == [266, 266]
    assert result["glyphs"] == recorded["glyphs"]
    assert result["warnings"] == recorded["warnings"] == []
    assert result["pages"] == recorded["pages"] == len(probe.CASES)


def test_word_counts_alike_in_every_setting():
    word = DATA["documents"]["numbering-none"]["word"]
    assert DATA["documents"]["numbering-14"]["word"] == word
    assert DATA["documents"]["numbering-15"]["word"] == word


@pytest.mark.parametrize("key,expected", sorted(LABELS.items()))
def test_what_word_drew(key, expected):
    assert " ".join(reader.labels(DATA["documents"]["numbering-15"]["word"][key])) == expected


_PARSE = parse_styles.parse_numbering


def _per_instance(xml, styles=None):
    found = _PARSE(xml, styles)
    found.lists = {num: ("num", num) for num in found.lists}
    return found


def _no_restart(xml, styles=None):
    found = _PARSE(xml, styles)
    found.restarts = frozenset()
    return found


@pytest.mark.parametrize("patch,families,lines", [
    # Counted per w:num, as docx2svg did: overrides and shared lists both fail.
    ((parse_document, "parse_numbering", _per_instance),
     {"override": [5, 15], "instance": [3, 8], "restart": [8, 8], "legal": [5, 5], "story": [7, 7]}, [229, 266]),
    # A startOverride as the level's start only, never restarting the shared count.
    ((parse_document, "parse_numbering", _no_restart),
     {"override": [9, 15], "instance": [8, 8], "restart": [8, 8], "legal": [5, 5], "story": [7, 7]}, [249, 266]),
    # Each text box counted on its own.
    ((linebreak, "text_box_labels", lambda document: {}),
     {"override": [15, 15], "instance": [8, 8], "restart": [8, 8], "legal": [5, 5], "story": [6, 7]}, [265, 266]),
])
def test_refuted_rules_keep_their_scores(patch, families, lines, fonts):
    name = "numbering-15"
    data = dict(reader.documents())[name]
    with mock.patch.object(*patch):
        result = reader.score(name, data, fonts, DATA["documents"][name])
    assert result["families"] == families
    assert result["lines"] == lines


def test_the_public_entry_point_counts_as_the_layout_does():
    """``ListCounters.item(numbering, numId, ilvl)`` -- the numbering part alone, no document
    or paragraph of docx2svg's -- gives every body paragraph of every case the label the
    layout's ``label`` gives it, item after item."""
    from docx2svg import parse_package
    from docx2svg.resolve import resolve_paragraph
    from docx2svg.model import Paragraph

    name = "numbering-15"
    data = dict(reader.documents())[name]
    document = parse_package(data)
    numbering = parse_styles.Numbering(document.numbering, document.numbering_lists,
                                       document.numbering_restarts)
    by_layout, by_entry = linebreak.ListCounters(), linebreak.ListCounters()
    compared = 0
    for paragraph in document.body:
        if not isinstance(paragraph, Paragraph) or paragraph.path in document.list_labels:
            continue
        pp = resolve_paragraph(document, paragraph)
        num_id, ilvl = pp.get("numPr.numId"), pp.get("numPr.ilvl", 0) or 0
        expected = by_layout.label(document, paragraph, pp)
        item = by_entry.item(numbering, num_id, ilvl)
        assert item.label == expected, paragraph.path
        if expected is not None:
            key = document.numbering_lists.get(num_id, ("num", num_id))
            assert item.number == by_layout.counts[key][ilvl]
            compared += 1
    assert compared > 60


def test_the_public_entry_point_on_a_numbering_part_alone():
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    xml = (f'<w:numbering {w}><w:abstractNum w:abstractNumId="0">'
           '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
           '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/>'
           '</w:lvl></w:abstractNum>'
           '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>'
           '<w:num w:numId="4"><w:abstractNumId w:val="0"/>'
           '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num>'
           '</w:numbering>').encode()
    numbering = _PARSE(xml)
    counters = linebreak.ListCounters()
    assert counters.item(numbering, 3, 0) == ("1.", 1)
    assert counters.item(numbering, 3, 1) == ("1.a)", 1)
    assert counters.item(numbering, 3, 1) == ("1.b)", 2)
    assert counters.item(numbering, 3, 0) == ("2.", 2)
    assert counters.item(numbering, 4, 0) == ("7.", 7)  # a startOverride restarts the shared list
    assert counters.item(numbering, 3, 0) == ("8.", 8)
    assert counters.item(numbering, 3, 5) == (None, None)  # no such level
    assert counters.item(numbering, 99, 0) == (None, None)  # no such instance
