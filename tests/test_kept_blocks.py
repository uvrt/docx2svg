"""The body's blocks kept from one version's parse for the next's (round 41,
:class:`docx2svg.parse.document.KeptBlocks`), and the parse a layout hands on
(:func:`docx2svg.parsed_document`).

**The property**: a document parsed with the blocks kept from earlier versions is the
document parsed from its bytes alone -- every field of every model object, paths and ids
included (compared as pickles) -- after any edit of a made-up report (:func:`report_docs.edit`),
each version parsed with what the one before kept; and nothing a layout or a drawing of its
pages does changes a model object (a layout's parse, laid out, drawn and laid out again
from, is still the parse of its bytes).  Laying out from a previous layout takes its kept
blocks (``tests/test_relayout.py`` holds those layouts to a full layout's, page for page).

Every document is made up (``tools/report_docs.py`` and the small packages below).
"""

from __future__ import annotations

import io
import pickle
import random
import zipfile

import pytest

import report_docs
from docx2svg import lay_out_docx, page_svgs, parsed_document
from docx2svg.parse.document import KeptBlocks, parse_package
from docx2svg.relayout import segments

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _same(one, other) -> bool:
    return pickle.dumps(one) == pickle.dumps(other)


def _main(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read("word/document.xml")


def _parse(data: bytes, kept: KeptBlocks):
    return parse_package(data, kept=kept, segments=segments(_main(data)))


@pytest.mark.parametrize("variant", [{}, {"notes": True}, {"sections": True}, {"toc": True, "fields": True},
                                     {"keeps": True}])
def test_a_parse_from_kept_blocks_is_the_parse_of_the_bytes(variant):
    rng = random.Random(41)
    kept = KeptBlocks()
    blocks = report_docs.report(6, seed=7, **variant)
    taken = 0
    for step in range(12):
        data = report_docs.package(blocks)
        mine = _parse(data, kept)
        assert _same(mine, parse_package(data)), f"step {step}"
        taken += kept.hits
        assert len(kept) <= len(segments(_main(data)))
        blocks, _ = report_docs.edit(blocks, rng)
    assert taken > 0


def test_no_object_is_shared_by_two_blocks_of_a_document():
    # The same paragraph three times, and a version that adds a fourth: each its own object.
    one = report_docs.Block("para", "A made-up paragraph said three times.")
    blocks = [report_docs.Block("heading", "Made-up title"), one, one, one]
    kept = KeptBlocks()
    first = _parse(report_docs.package(blocks), kept)
    again = _parse(report_docs.package(blocks + [one]), kept)
    assert kept.hits > 0
    for document in (first, again):
        assert len({id(block) for block in document.body}) == len(document.body)
        runs = [id(run) for block in document.body for run in getattr(block, "runs", ())]
        assert len(set(runs)) == len(runs)
    # A block taken is a copy of the one kept, never the object an earlier document holds.
    assert not {id(block) for block in first.body} & {id(block) for block in again.body}
    assert [block.path for block in again.body] == [block.path for block in parse_package(
        report_docs.package(blocks + [one])).body]


def test_blocks_inside_an_open_field_are_built_each_time():
    blocks = report_docs.report(6, seed=3, toc=True)
    kept = KeptBlocks()
    data = report_docs.package(blocks)
    _parse(data, kept)
    _parse(data, kept)
    # The contents' lines between the field's begin and end are inside it: built again.
    lines = sum(1 for block in blocks if block.kind == "toc")
    assert lines > 2 and kept.built >= lines - 1
    assert _same(_parse(data, kept), parse_package(data))


def _package(colour: str, mapping: str, prefix: str = "w", extra: str = "") -> bytes:
    body = (f'<{prefix}:body><{prefix}:p><{prefix}:r><{prefix}:rPr><{prefix}:color {prefix}:val="000000" '
            f'{prefix}:themeColor="text1"/></{prefix}:rPr><{prefix}:t>Made-up text</{prefix}:t></{prefix}:r>'
            f'</{prefix}:p><{prefix}:sectPr/></{prefix}:body>')
    document = f'<?xml version="1.0"?><{prefix}:document xmlns:{prefix}="{W_NS}"{extra}>{body}</{prefix}:document>'
    theme = ('<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
             'name="Made-up"><a:themeElements><a:clrScheme name="Made-up">'
             f'<a:dk1><a:srgbClr val="{colour}"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
             '<a:dk2><a:srgbClr val="112233"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>'
             '</a:clrScheme></a:themeElements></a:theme>')
    settings = (f'<?xml version="1.0"?><w:settings xmlns:w="{W_NS}"><w:clrSchemeMapping w:t1="{mapping}" '
                'w:bg1="light1"/></w:settings>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    parts = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                               'content-types"><Default Extension="xml" ContentType="application/xml"/>'
                               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
                               'relationships+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                       f'relationships"><Relationship Id="r1" Type="{rel}/officeDocument" '
                       'Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.'
                                        f'org/package/2006/relationships"><Relationship Id="r1" Type="{rel}/theme" '
                                        f'Target="theme/theme1.xml"/><Relationship Id="r2" Type="{rel}/settings" '
                                        'Target="settings.xml"/></Relationships>',
        "word/document.xml": document, "word/theme/theme1.xml": theme, "word/settings.xml": settings}
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
    return buffer.getvalue()


@pytest.mark.parametrize("change", [("445566", "dark1"), ("123456", "dark2"), ("123456", "dark1", "x"),
                                    ("123456", "dark1", "w", ' xmlns:v="urn:made-up"')])
def test_a_block_is_not_taken_where_the_theme_colours_or_namespaces_differ(change):
    kept = KeptBlocks()
    _parse(_package("123456", "dark1"), kept)
    data = _package(*change)
    mine = _parse(data, kept)
    assert kept.hits == 0
    assert _same(mine, parse_package(data))
    colour = mine.body[0].runs[0].properties.color
    assert colour == parse_package(data).body[0].runs[0].properties.color
    # The same bytes again are taken.
    _parse(data, kept)
    assert kept.hits == 1


def test_a_layouts_parse_is_unchanged_by_laying_out_drawing_and_laying_out_again():
    rng = random.Random(5)
    blocks = report_docs.report(8, seed=11, notes=True)
    data = report_docs.package(blocks)
    layout, data, fonts = lay_out_docx(data)
    versions = [(layout, data)]
    for _ in range(5):
        page_svgs(layout, data, fonts=fonts)
        blocks, _ = report_docs.edit(blocks, rng)
        layout, data, fonts = lay_out_docx(report_docs.package(blocks), previous=layout)
        versions.append((layout, data))
    page_svgs(layout, data, fonts=fonts)
    assert sum(layout.state.blocks.hits for layout, _ in versions[1:]) > 0
    for layout, data in versions:
        # The document as parsed for each layout, after every layout and drawing since: the
        # parse of its bytes, every model object as it was made.
        assert parsed_document(layout, data) is layout.state.parsed
        assert _same(layout.state.parsed, parse_package(data))


def test_segments_read_before_are_taken_by_their_bytes():
    rng = random.Random(9)
    blocks = report_docs.report(6, seed=5)
    layout = lay_out_docx(report_docs.package(blocks))[0]
    for _ in range(4):
        blocks, _ = report_docs.edit(blocks, rng)
        data = report_docs.package(blocks)
        before = layout.state.known_segments
        layout = lay_out_docx(data, previous=layout)[0]
        mine = layout.state.segments
        assert mine == segments(_main(data))
        assert sum(1 for segment in mine if any(segment is kept for kept in before.values())) > len(mine) // 2


def test_what_a_layout_keeps_in_its_process_is_not_pickled():
    data = report_docs.package(report_docs.report(3, seed=2))
    layout = lay_out_docx(data)[0]
    assert len(layout.state.blocks) > 0 and layout.state.parsed is not None
    back = pickle.loads(pickle.dumps(layout))
    assert back.state.blocks is None and back.state.parsed is None and back.state.known_segments is None
    assert len(pickle.loads(pickle.dumps(layout.state.blocks))) == 0
    # Without them, the next layout builds every block; with them, it takes them.
    again = lay_out_docx(data, previous=back)[0]
    assert again.state.blocks.hits == 0
    again = lay_out_docx(data, previous=layout)[0]
    assert again.state.blocks.hits > 0
    # Another document's bytes are parsed, not taken.
    other = report_docs.package(report_docs.report(3, seed=4))
    assert _same(parsed_document(layout, other), parse_package(other))
    assert lay_out_docx(data, keep=False)[0].state is None
