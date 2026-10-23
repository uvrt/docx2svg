"""``PAGEREF``, ``REF`` and ``SEQ`` computed as Word computes them on export (ROADMAP.md,
"Headers, footers and fields -- measured", H.8).  ``tools/make_xref_field_probe.py``,
recorded by ``tools/read_xref_field_probe.py`` in
``tests/fixtures/xref-field-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import read_xref_field_probe  # noqa: E402
import render_record  # noqa: E402

from docx2svg import fields  # noqa: E402
from docx2svg.parse import document as parse_document  # noqa: E402

DATA = json.loads(read_xref_field_probe.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = read_xref_field_probe.documents()
#: What is not computed, and why Word's drawing differs: a missing bookmark's ``REF`` and
#: ``PAGEREF`` (Word's error text, in the application's language), ``PAGEREF \\p``
#: ("above"), ``REF`` of a bookmark over two paragraphs (Word draws both), ``SEQ \\s``.
WARNINGS = ["field-not-computed:PAGEREF", "field-not-computed:REF", "field-not-computed:SEQ"]


def _score(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    result, _rects, warnings = read_render.score(name, data, fonts, DATA["documents"][name]["objects"], [])
    return result, warnings


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_computable_cross_reference_is_drawn_as_word_draws_it(name, data):
    result, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == DATA["documents"][name]["warnings"] == WARNINGS
    # Every glyph off is one of the five fields not computed: a closing bracket after a
    # cached result of another width, and one line moved by the paragraph Word adds.
    assert result.matched == 2033 and result.word == 2147
    assert result.y_agree == 2032 and result.face_agree == result.size_agree == result.matched
    assert sorted({problem.split("'")[1] for problem in result.problems}) == ["]"]


@pytest.mark.parametrize("index,expected", [
    (0, (1853, 1809, 1656, 1852)),  # before: every PAGEREF, REF and SEQ drawn as cached
    (1, (2033, 1901, 1825, 1907)),  # REF in its instruction's format: Word draws the bookmark's own
    (2, (2029, 2022, 2028, 2029)),  # PAGEREF in its own section's format: Word uses the bookmark's page's
])
def test_refuted_rules_keep_their_scores(index, expected):
    name, data = DOCUMENTS[1]
    _what, patch = read_xref_field_probe.refuted()[index]
    with mock.patch.object(*patch):
        result, _warnings = _score(name, data)
    assert (result.matched, result.x_agree, result.y_agree, result.face_agree) == expected


def test_sequences_count_as_word_counts():
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

    def seq(instruction):
        return (f'<w:p><w:fldSimple w:instr=" {instruction} "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>')

    body = "".join(seq(i) for i in ("SEQ Figure", "SEQ figure \\* roman", "SEQ Table", "SEQ Figure \\c",
                                    "SEQ Figure \\h", "SEQ Figure \\r 7", "SEQ Figure \\* alphabetic",
                                    "SEQ Eq \\s 1", "SEQ Eq"))
    root = parse_document.parse_xml(f'<w:document {w}><w:body>{body}</w:body></w:document>'.encode())
    blocks = parse_document._story(parse_document.child(root, "body"), path="w:body",
                                   fields=parse_document._Fields(cross_references=True))
    values = fields.sequence_values(fields.body_paragraphs(blocks))
    texts = [values[(k, 0)] for k in range(9)]
    assert texts[:7] == ["1", "ii", "1", "2", "", "7", "h"]
    assert all(isinstance(value, fields.Uncomputable) for value in texts[7:])


#: Word 16.106 (macOS), round 16's ``f4_probe.py``: SEQ fields in document order, each
#: updated by Word, and what it drew -- ``None`` where it drew an error in its language.
SEQ_ORDERS = [
    ('SEQ Aa \\# "00"', "01"), ('SEQ Aa \\# "00"', "02"),
    ('SEQ \\# "00" Bb', "01"), ('SEQ \\# "00" Bb', "02"),
    ('SEQ \\#"00" Cc', "01"), ('SEQ \\#"00" Cc', "02"),
    ("SEQ Dd", "1"), ('SEQ \\# "00" Dd', "02"), ('SEQ Dd \\# "000"', "003"),
    ("SEQ \\* ARABIC Ee", "1"), ("SEQ \\* roman Ee", "ii"), ("SEQ Ee \\* ROMAN", "III"),
    ("SEQ Ff \\r 5", "5"), ("SEQ \\r 7 Ff", "7"), ("SEQ Ff", "8"),
    ("SEQ Gg", "1"), ("SEQ \\c Gg", "1"), ("SEQ Gg \\c", "1"),
    ("SEQ \\h Hh", ""), ("SEQ Hh", "2"), ("SEQ Hh \\h", ""), ("SEQ Hh", "4"),
    ("SEQ \\n Ii", "1"), ("SEQ \\n Ii \\* MERGEFORMAT", "2"),
    ('SEQ \\r 3 \\# "00" Jj', "03"), ('SEQ \\* MERGEFORMAT \\# "00" Jj', None),
    ("SEQ \\* alphabetic Kk", "a"), ('SEQ \\# "#0" Kk', " 2"),
    ("SEQ \\r Mm", None), ("SEQ Mm", "1"),
    ('SEQ Nn \\* MERGEFORMAT \\# "00"', None), ('SEQ Nn \\# "00" \\* MERGEFORMAT', "02"),
    ('SEQ Oo \\* ARABIC \\# "00"', None), ('SEQ \\# "00" \\* ARABIC Oo', "2"),
    ('SEQ \\# "##0" Qq', "  1"), ('SEQ \\# "#" Qq', "2"), ('SEQ \\# "0#" Qq', "03"),
    ('SEQ \\r 120 \\# "#0" Rr', "120"), ("SEQ \\# 00 Rr", "121"),
    ('SEQ "Tt"', "1"), ('SEQ \\# "00" "Tt"', "02"),
    ("SEQ \\h \\r 4 Uu", ""), ("SEQ Uu", "5"),
]


def test_sequence_switches_stand_before_or_after_the_identifier():
    """Word reads a SEQ's identifier wherever its switches stand (round 16, F4)."""
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    body = "".join(f'<w:p><w:fldSimple w:instr=" {instruction.replace(chr(34), "&quot;")} ">'
                   '<w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>' for instruction, _ in SEQ_ORDERS)
    root = parse_document.parse_xml(f'<w:document {w}><w:body>{body}</w:body></w:document>'.encode())
    blocks = parse_document._story(parse_document.child(root, "body"), path="w:body",
                                   fields=parse_document._Fields(cross_references=True))
    values = fields.sequence_values(fields.body_paragraphs(blocks))
    drawn = [values[(k, 0)] for k in range(len(SEQ_ORDERS))]
    assert [None if isinstance(v, fields.Uncomputable) else v for v in drawn] == [w for _, w in SEQ_ORDERS]
    assert fields.sequence_switches(' SEQ \\# "00" Item ') == ("Item", [("#", "00")])
    assert fields.sequence_switches(" SEQ \\r Item ") == (None, [("r", "Item")])


def test_field_text_takes_pageref():
    assert fields.field_text(" PAGEREF _Toc1 \\h ", "PAGEREF", page=11, pages=None, section_pages=None,
                             section_format="lowerRoman") == "xi"
    assert fields.field_text(' PAGEREF x \\# "00" ', "PAGEREF", page=3, pages=None, section_pages=None,
                             section_format=None) == "03"
    with pytest.raises(fields.Uncomputable):
        fields.field_text(" PAGEREF x \\p ", "PAGEREF", page=3, pages=None, section_pages=None, section_format=None)
