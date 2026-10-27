"""Pages and blocks, both ways (round 19, R3): :attr:`Page.blocks`, :meth:`Layout.blocks_on`,
:meth:`Layout.pages_of` and :func:`docx2svg.paths.block_keys`.

The documents are made up (``tools/report_docs.py``): a report of about 16 pages, and
blocks built to cross a page -- a paragraph too long for one, a table of many rows with a
header row Word draws again on each page.
"""

from __future__ import annotations

import pytest

import report_docs
from docx2svg import lay_out_docx
from docx2svg.paths import block_keys, locate
from report_docs import Block

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture(scope="module")
def report():
    blocks = report_docs.report(16)
    data = report_docs.package(blocks)
    return blocks, data, lay_out_docx(data)[0]


@pytest.fixture(scope="module")
def crossing():
    """A paragraph longer than a page, then a table of 70 rows whose first is a header row."""
    long_text = " ".join(f"Sentence {k} of a paragraph that goes on past the foot of its page." for k in range(140))
    rows = (("Name", "Kind", "Note"),) + tuple((f"Row {k}", "plain", f"cell {k}") for k in range(1, 70))
    blocks = [Block("heading", "A long paragraph", para_id="10000001"), Block("para", long_text, para_id="10000002"),
              Block("table", rows=rows), Block("para", "After the table.", para_id="10000003")]
    data = report_docs.package(blocks)
    return blocks, data, lay_out_docx(data)[0]


def test_block_keys_are_every_block_prefix_up_to_the_paragraph():
    assert block_keys("w:body/w:p[3]") == ["w:body/w:p[3]"]
    assert block_keys("w:body/w:p[3]/w:r[2]") == ["w:body/w:p[3]"]
    assert block_keys("w:body/w:tbl[2]/w:tr[3]/w:tc[1]/w:p[2]") == [
        "w:body/w:tbl[2]", "w:body/w:tbl[2]/w:tr[3]", "w:body/w:tbl[2]/w:tr[3]/w:tc[1]",
        "w:body/w:tbl[2]/w:tr[3]/w:tc[1]/w:p[2]"]
    # A text box's paragraph is its anchor's paragraph's, as the body sees it.
    assert block_keys("w:body/w:p[5]/w:r[1]/wp:anchor[1]/wps:txbx/w:txbxContent/w:p[1]") == ["w:body/w:p[5]"]
    assert block_keys("w:body/w:sdt[1]/w:sdtContent/w:p[2]") == ["w:body/w:sdt[1]/w:sdtContent/w:p[2]"]
    assert block_keys("w:hdr/w:p[1]") == []
    assert block_keys("w:footnote[@w:id=2]/w:p[1]") == []


def test_every_key_names_its_element_in_the_main_part(report):
    _, data, layout = report
    for page in layout.pages:
        for key in page.blocks:
            part, element = locate(data, key)
            assert part == "word/document.xml"
            assert element is not None, key
            assert element.tag == W + key.rsplit("/", 1)[1].split("[")[0].split(":")[1], key


def test_every_top_level_block_is_on_a_page_and_in_order(report):
    blocks, _, layout = report
    paragraphs = sum(1 for block in blocks if block.kind != "table")
    tables = sum(1 for block in blocks if block.kind == "table")
    seen = [key for number in range(1, len(layout.pages) + 1) for key in layout.blocks_on(number)]
    expected = {f"w:body/w:p[{k}]" for k in range(1, paragraphs + 1)} | {
        f"w:body/w:tbl[{k}]" for k in range(1, tables + 1)}
    assert set(seen) == expected
    # Document order: a block split across two pages is the last of one and the first of the next.
    order = {key: k for k, key in enumerate(dict.fromkeys(seen))}
    for number in range(1, len(layout.pages) + 1):
        keys = layout.blocks_on(number)
        assert [order[key] for key in keys] == sorted(order[key] for key in keys)
    assert layout.blocks_on(0) == [] and layout.blocks_on(len(layout.pages) + 1) == []


def test_the_two_directions_agree(report):
    blocks, _, layout = report
    for page in layout.pages:
        for key in page.blocks:
            assert page.number + 1 in layout.pages_of(key)
    for key in {key for page in layout.pages for key in page.blocks}:
        pages = layout.pages_of(key)
        assert pages == sorted(pages) and len(set(pages)) == len(pages)
        assert all(key in layout.pages[number - 1].blocks for number in pages)
    # By w14:paraId, any case, the same pages as by path.
    k = 0
    for block in blocks:
        if block.kind == "table":
            continue
        k += 1
        assert layout.pages_of(block.para_id) == layout.pages_of(f"w:body/w:p[{k}]")
        assert layout.pages_of(block.para_id.lower()) == layout.pages_of(block.para_id)


def test_a_block_across_pages_is_on_each(crossing):
    _, _, layout = crossing
    assert len(layout.pages_of("w:body/w:p[2]")) >= 2
    assert layout.pages_of("10000002") == layout.pages_of("w:body/w:p[2]")
    table = layout.pages_of("w:body/w:tbl[1]")
    assert len(table) >= 2
    # The header row is drawn again on every page the table is on; a body row on one.
    assert layout.pages_of("w:body/w:tbl[1]/w:tr[1]") == table
    assert layout.pages_of("w:body/w:tbl[1]/w:tr[1]/w:tc[2]") == table
    assert layout.pages_of("w:body/w:tbl[1]/w:tr[40]") in ([page] for page in table)
    # A run's path is its paragraph's; what is not drawn is on no page.
    assert layout.pages_of("w:body/w:p[3]/w:r[1]") == layout.pages_of("w:body/w:p[3]")
    assert layout.pages_of("w:body/w:p[99]") == []
    assert layout.pages_of("0BADF00D") == []


def test_stories_are_not_the_bodys(report):
    _, _, layout = report
    assert any(line.story for page in layout.pages for line in page.lines)
    assert all(key.startswith("w:body/") for page in layout.pages for key in page.blocks)
