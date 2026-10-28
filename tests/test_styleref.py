"""``STYLEREF`` computed per page as Word computes it (round 21): in a header or footer, the
first text of its style on the page (``\\l``: the last), else the last before the page, else
the first after it; in the body, the last before the field, else the first after it.

``tools/make_styleref_probe.py`` makes the documents; Word's values, page by page, are
``tests/fixtures/styleref-observations.json`` (Word for Mac 16, English interface, every
field updated, read from its PDF).  Every probe field is labelled ``X=[...]`` and cached as
``?``: a field docx2svg does not compute is drawn as ``?``, with a warning.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_styleref_probe  # noqa: E402

from docx2svg import ConvertOptions, fields, lay_out_docx  # noqa: E402
from docx2svg.model import Style, StyleSheet  # noqa: E402

DATA = json.loads((REPO / "tests" / "fixtures" / "styleref-observations.json").read_text(encoding="utf-8"))
LABEL = re.compile(r"(\w+)=\[([^\]]*)\]")

#: The fields drawn as cached, per document, and why: the paragraph number switches
#: (``\\n``, ``\\r``, ``\\w``, ``\\s``), ``STYLEREF`` to the default character style, a switch
#: Word does not know, and a heading with a tab in it.
CACHED = {
    "switches": {"s1", "s2", "s3", "s5", "s7", "s8", "s9", "sa", "sb", "se"},
    "latent": {"l8", "lc"},
    "extra": {"S", "SL", "FT", "FTL"},
}


def _values(data: bytes) -> tuple[list[dict], list]:
    options = ConvertOptions()
    layout, _, _ = lay_out_docx(data, options)
    pages = []
    for page in layout.pages:
        text = " ".join("".join("".join(span.chars) for span in line.spans) for line in page.lines)
        pages.append({label: " ".join(value.split()) for label, value in LABEL.findall(text)})
    return pages, options.warnings


@pytest.mark.parametrize("name", sorted(DATA["documents"]))
def test_every_computed_styleref_is_the_text_word_shows(name):
    word = DATA["documents"][name]
    pages, warnings = _values(make_styleref_probe.case(name))
    assert len(pages) == len(word)
    cached = CACHED.get(name, set())
    compared = 0
    for number, (ours, theirs) in enumerate(zip(pages, word), start=1):
        assert set(ours) == set(theirs), number
        for label, text in theirs.items():
            if ours[label] == "?":
                assert label in cached, (number, label, text)
                continue
            assert ours[label] == text, (number, label)
            compared += 1
    assert compared
    codes = {warning.code for warning in warnings if warning.code.startswith("field")}
    assert codes == ({"field-not-computed:STYLEREF"} if cached else set()), codes


def test_a_header_that_wraps_on_one_page_moves_that_pages_body_only():
    # Word (``height``): the body's first glyph 5.2 pt lower on the page whose header's
    # STYLEREF wraps to two lines; the cached result changes nothing (``heightlong``).
    for name in ("height", "heightlong"):
        layout, _, _ = lay_out_docx(make_styleref_probe.case(name))
        tops = [page.column[2] for page in layout.pages]
        assert tops[1] > tops[0] == tops[2] == tops[3], name


def test_style_reference_switches():
    assert fields.style_reference(' STYLEREF "Chapter title" \\l ') == fields.StyleReference("Chapter title", True)
    assert fields.style_reference("STYLEREF 1 \\p").position
    assert fields.style_reference('STYLEREF " Chapter title "').name == " Chapter title "
    for instruction in ('STYLEREF "Heading 1" \\n', 'STYLEREF "Heading 1" \\w', 'STYLEREF "Heading 1" \\x',
                        'STYLEREF "Heading 1" \\* roman', 'STYLEREF \\l "Heading 1"', "STYLEREF Chapter title"):
        with pytest.raises(fields.Uncomputable):
            fields.style_reference(instruction)


def test_find_style_by_name_case_ignored_or_by_heading_level():
    sheet = StyleSheet(styles={"Heading1": Style("Heading1", "paragraph", "heading 1"),
                               "Chaptertitle": Style("Chaptertitle", "paragraph", "Chapter title"),
                               "Keyterm": Style("Keyterm", "character", "Key term")})
    assert fields.find_style("HEADING 1", sheet) == ("paragraph", frozenset({"Heading1"}))
    assert fields.find_style("1", sheet) == ("paragraph", frozenset({"Heading1"}))
    assert fields.find_style("3", sheet) == ("paragraph", frozenset())
    assert fields.find_style("key term", sheet) == ("character", frozenset({"Keyterm"}))
    assert fields.find_style("Chaptertitle", sheet) is None  # a style id is not a name
    assert fields.find_style(" Chapter title ", sheet) is None
    # In a language whose error texts are not measured, an unknown name is not computed.
    reference = fields.style_reference('STYLEREF "Kop 1"')
    assert fields.style_reference_target(reference, sheet, "en-US").startswith("Error! Use the Home tab")
    with pytest.raises(fields.Uncomputable):
        fields.style_reference_target(reference, sheet, "nl-NL")
