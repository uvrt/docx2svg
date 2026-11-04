"""A table of contents' entries: the ``Hyperlink`` style in a ``TOC`` field's links.

Word (16, Mac; measured with made-up documents, ``toc_link_layer``'s docstring) draws the
text of a ``w:hyperlink`` inside a ``TOC`` field's result in the entry's paragraph style:
of the ``Hyperlink`` (or ``FollowedHyperlink``) character style only the fonts and the
capitals are kept, and direct formatting still applies.  The documents here are made up.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

import probe_docx  # noqa: E402
import wml  # noqa: E402

from docx2svg import convert_docx_to_svg, parse_package  # noqa: E402
from docx2svg.resolve import resolve_run  # noqa: E402

LOUD = {"rFonts": {"ascii": "Courier New", "hAnsi": "Courier New"}, "b": True, "i": True, "caps": True,
        "strike": True, "color": "FF0000", "spacing": 40, "w": 150, "position": 6, "sz": 28, "u": "double",
        "shd": {"val": "clear", "color": "auto", "fill": "00FFFF"}, "vertAlign": "superscript"}
MARK = '<w:rStyle w:val="{}"/><w:noProof/>'


def _fld(kind: str) -> str:
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def _instr(text: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'


def _text(text: str, rpr: str) -> str:
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def _entry(text: str, rpr: str, *, link: bool = True, head: str = "", level: int = 1) -> str:
    inner = _text(text, rpr) + '<w:r><w:tab/></w:r>' + _text("1", "<w:noProof/>")
    body = f'<w:hyperlink w:anchor="_Toc1" w:history="1">{inner}</w:hyperlink>' if link else inner
    return f'<w:p><w:pPr><w:pStyle w:val="Contents{level}"/></w:pPr>{head}{body}</w:p>'


def _document(entries: list[tuple[str, str, bool]], *, styles: list[str] = (), link_style: dict | None = None,
              link_id: str = "Hyperlink", link_name: str = "Hyperlink", link_based: str | None = None,
              contents_rpr: dict | None = None, instruction: str = ' TOC \\o "1-3" \\h \\z \\u ',
              after: str = "") -> bytes:
    """A TOC field holding ``entries`` (text, run properties, in a ``w:hyperlink``), then ``after``."""
    head = _fld("begin") + _instr(instruction) + _fld("separate")
    body = "".join(_entry(text, rpr, link=link, head=head if k == 0 else "")
                   for k, (text, rpr, link) in enumerate(entries))
    body += "<w:p>" + _fld("end") + "</w:p>" + after
    body += '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc1"/>' \
            + wml.run("A heading") + '<w:bookmarkEnd w:id="1"/></w:p>'
    sheet = [wml.style("paragraph", "Normal", default=True),
             wml.style("character", "DefaultParagraphFont", name="Default Paragraph Font", default=True),
             wml.style("paragraph", "Heading1", name="heading 1", based_on="Normal", ppr_={"outlineLvl": 0}),
             wml.style("paragraph", "Contents1", name="toc 1", based_on="Normal", rpr_=contents_rpr),
             wml.style("character", link_id, name=link_name, based_on=link_based or "DefaultParagraphFont",
                       rpr_=LOUD if link_style is None else link_style),
             *styles]
    return probe_docx.package(body, styles=wml.styles_part(
        sheet, run_defaults={"rFonts": {"ascii": "Arial", "hAnsi": "Arial"}, "sz": 20}))


def _resolved(data: bytes, text: str):
    document = parse_package(data)
    for paragraph in document.body:
        for run in getattr(paragraph, "runs", ()):
            if run.text == text:
                return run, resolve_run(document, paragraph, run)
    raise AssertionError(f"no run {text!r}")


def _face(resolved) -> str | None:
    face = resolved.get("rFonts.ascii")
    return getattr(face, "face", face)


def test_an_entrys_link_takes_only_the_hyperlink_styles_fonts_and_capitals():
    run, resolved = _resolved(_document([("Entry one", MARK.format("Hyperlink"), True)]), "Entry one")
    assert run.toc_link
    assert _face(resolved) == "Courier New"
    assert resolved.get("caps") is True
    for key in ("b", "i", "strike"):
        assert not resolved.get(key), key
    for key in ("color", "spacing", "w", "position", "u", "shd", "vertAlign"):
        assert resolved.get(key) is None, key
    assert resolved.get("sz") == 20


def test_small_capitals_are_drawn_as_capitals():
    data = _document([("Entry one", MARK.format("Hyperlink"), True)], link_style={"smallCaps": True, "vanish": True})
    _, resolved = _resolved(data, "Entry one")
    assert resolved.get("caps") is True and not resolved.get("smallCaps") and not resolved.get("vanish")


def test_direct_formatting_on_the_entry_still_applies():
    rpr = MARK.format("Hyperlink") + '<w:b/><w:color w:val="00B050"/><w:u w:val="dotted"/>'
    _, resolved = _resolved(_document([("Entry one", rpr, True)]), "Entry one")
    assert resolved.get("b") is True and resolved.get("color") == "00B050" and resolved.get("u") == "dotted"


def test_the_entrys_paragraph_style_draws_it():
    data = _document([("Entry one", MARK.format("Hyperlink"), True)], contents_rpr={"color": "7030A0", "b": True})
    _, resolved = _resolved(data, "Entry one")
    assert resolved.get("color") == "7030A0" and resolved.get("b") is True   # the style's bold toggles nothing


@pytest.mark.parametrize("instruction", [' TOC \\o "1-3" \\h \\z \\u ', ' TOC \\o "1-3" \\z \\u '])
def test_without_a_link_the_style_is_drawn_whole(instruction):
    data = _document([("Linked", MARK.format("Hyperlink"), True), ("Unlinked", MARK.format("Hyperlink"), False)],
                     instruction=instruction)
    assert _resolved(data, "Linked")[1].get("color") is None
    run, resolved = _resolved(data, "Unlinked")
    assert not run.toc_link and resolved.get("color") == "FF0000" and resolved.get("u") == "double"


def test_a_link_outside_a_table_of_contents_is_drawn_whole():
    after = ('<w:p><w:hyperlink w:anchor="_Toc1">' + _text("Plain link", MARK.format("Hyperlink")) + '</w:hyperlink></w:p>'
             + '<w:p>' + _fld("begin") + _instr(' HYPERLINK \\l "_Toc1" ') + _fld("separate")
             + _text("Field link", MARK.format("Hyperlink")) + _fld("end") + '</w:p>')
    data = _document([("Entry one", MARK.format("Hyperlink"), True)], after=after)
    for text in ("Plain link", "Field link"):
        run, resolved = _resolved(data, text)
        assert not run.toc_link and resolved.get("color") == "FF0000" and resolved.get("b") is True, text


def test_the_style_is_known_by_its_name_not_its_id():
    by_name = _document([("Entry one", MARK.format("LinkA"), True)], link_id="LinkA")
    assert _resolved(by_name, "Entry one")[1].get("color") is None
    by_id = _document([("Entry one", MARK.format("Hyperlink"), True)], link_name="Link B")
    assert _resolved(by_id, "Entry one")[1].get("color") == "FF0000"


def test_followed_hyperlink_too_but_not_a_style_based_on_hyperlink():
    followed = wml.style("character", "FollowedHyperlink", name="FollowedHyperlink",
                         based_on="DefaultParagraphFont", rpr_={"color": "954F72", "u": "single"})
    derived = wml.style("character", "LinkC", name="Link C", based_on="Hyperlink", rpr_={"b": True})
    data = _document([("Followed", MARK.format("FollowedHyperlink"), True), ("Derived", MARK.format("LinkC"), True)],
                     styles=[followed, derived])
    assert _resolved(data, "Followed")[1].get("color") is None
    assert _resolved(data, "Derived")[1].get("color") == "FF0000"


def test_what_the_hyperlink_style_inherits_is_left_out_with_it():
    base = wml.style("character", "AccentD", name="Accent D", based_on="DefaultParagraphFont",
                     rpr_={"color": "00B050", "u": "wave"})
    data = _document([("Entry one", MARK.format("Hyperlink"), True)], styles=[base], link_style={},
                     link_based="AccentD")
    _, resolved = _resolved(data, "Entry one")
    assert resolved.get("color") is None and resolved.get("u") is None


def test_the_entry_is_drawn_black_and_not_underlined():
    link = {"color": "0563C1", "u": "single"}
    after = '<w:p><w:hyperlink w:anchor="_Toc1">' + _text("Plain link", MARK.format("Hyperlink")) + '</w:hyperlink></w:p>'
    svg = convert_docx_to_svg(_document([("Entry one", MARK.format("Hyperlink"), True)], link_style=link,
                                        after=after))[0]
    entry = re.search(r'<text data-docx-path="w:body/w:p\[1\]/w:hyperlink\[1\]/w:r\[1\]"[^>]*fill="(#\w+)"', svg)
    plain = re.search(r'<text data-docx-path="w:body/w:p\[3\]/w:hyperlink\[1\]/w:r\[1\]"[^>]*fill="(#\w+)"', svg)
    assert entry and entry.group(1) == "#000000"
    assert plain and plain.group(1) == "#0563C1"
    unders = re.findall(r'data-docx-path="w:body/w:p\[(\d)\][^"]*" data-docx-kind="underline"', svg)
    assert unders and set(unders) == {"3"}
