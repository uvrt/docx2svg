"""Laying a document out again from a previous layout (round 19, R1): :func:`lay_out_docx`
with ``previous``, :mod:`docx2svg.relayout`.

**The property**: after any edit, the layout made from the previous one is the layout of
the new bytes -- every page's SVG byte for byte, the warnings and the coverage -- and
``changed_pages`` is exactly the pages whose drawing differs from the previous layout's
page of the same number (their SVG, ``data-docx-*`` attributes aside), with those it did
not have.  Held over seeded random edits (:func:`report_docs.edit`: a paragraph's or a
cell's text changed or grown, paragraphs, rows and page breaks inserted and deleted) of
made-up reports with headings, a list, tables with a header row, side panels text wraps
beside, a header and a ``Page N of M`` footer, each layout made from the one before.
``RELAYOUT_CASES`` (an environment variable) sets how many edits of the small reports
run (CI's default: :data:`CASES`); the 16- and 40-page reports take :data:`LONG_CASES`.

Every document is made up (``tools/report_docs.py``).
"""

from __future__ import annotations

import os
import random
import re
from dataclasses import replace

import pytest

import report_docs
from docx2svg import ConvertOptions, lay_out_docx, page_svgs
from docx2svg.relayout import FALLBACKS

#: Edits of the small reports CI runs; ``RELAYOUT_CASES`` sets another number.
CASES = int(os.environ.get("RELAYOUT_CASES", "96"))
#: Edits of the 16-page and the 40-page report.
LONG_CASES = {16: int(os.environ.get("RELAYOUT_LONG_CASES", "6")), 40: int(os.environ.get("RELAYOUT_LONG_CASES",
                                                                                        "2"))}
#: How many edits are made in a row, each layout from the one before, before starting again.
CHAIN = 8

_DATA = re.compile(r' data-docx-[a-z-]+="[^"]*"')


def _full(data: bytes):
    options = ConvertOptions()
    layout, data, fonts = lay_out_docx(data, options, keep=False)
    return layout, page_svgs(layout, data, fonts=fonts), options


def _again(data: bytes, previous):
    options = ConvertOptions()
    layout, data, fonts = lay_out_docx(data, options, previous=previous)
    return layout, page_svgs(layout, data, fonts=fonts), options


def _check(blocks, previous, previous_svgs, what: str, stats: dict, package: dict | None = None):
    data = report_docs.package(blocks, **(package or {}))
    layout, svgs, options = _again(data, previous)
    full, full_svgs, full_options = _full(data)
    assert len(svgs) == len(full_svgs), what
    for number, (mine, theirs) in enumerate(zip(svgs, full_svgs), start=1):
        assert mine == theirs, f"{what}: page {number} differs from a full layout's"
    assert [(w.code, w.message, w.page) for w in options.warnings] == [
        (w.code, w.message, w.page) for w in full_options.warnings], what
    assert options.coverage.as_dict() == full_options.coverage.as_dict(), what
    differ = [number for number, svg in enumerate(svgs, start=1)
              if number > len(previous_svgs) or _DATA.sub("", svg) != _DATA.sub("", previous_svgs[number - 1])]
    assert layout.changed_pages == differ, what
    assert layout.relaid is None or layout.relaid in FALLBACKS, what
    stats.setdefault(layout.relaid, 0)
    stats[layout.relaid] += 1
    return layout, svgs


def _run(pages: int, seed: int, cases: int, *, notes: bool = False, sections: bool = False, **package) -> dict:
    rng = random.Random(seed)
    stats: dict = {}
    done = 0
    while done < cases:
        blocks = report_docs.report(pages, seed=rng.randrange(1 << 30), notes=notes, sections=sections)
        layout, data, fonts = lay_out_docx(report_docs.package(blocks, **package))
        svgs = page_svgs(layout, data, fonts=fonts)
        for _ in range(min(CHAIN, cases - done)):
            blocks, what = report_docs.edit(blocks, rng)
            layout, svgs = _check(blocks, layout, svgs, f"seed {seed}, {pages} pages, {what}", stats, package)
            done += 1
    return stats


@pytest.mark.parametrize("seed", range(4))
def test_random_edits_of_small_reports_lay_out_as_a_full_layout(seed):
    stats = _run(4, seed, CASES // 4)
    # Most edits are laid out only where they reach.
    assert stats.get(None, 0) >= sum(stats.values()) // 3, stats


def test_random_edits_of_reports_with_footnotes_lay_out_as_a_full_layout():
    # Adding or deleting a footnote's paragraph numbers the notes again: laid out whole.
    stats = _run(6, 7, max(CASES // 8, 4), notes=True)
    assert set(stats) <= {None, "notes", "note-references"}, stats


@pytest.mark.parametrize("variant", [
    {"sections": True},
    {"even_odd": True, "title_page": True},
    {"sections": True, "even_odd": True},
])
def test_random_edits_of_reports_with_sections_and_their_own_headers(variant):
    # Pages that move on keep their headers only where no section starts among them and,
    # where even pages have their own, by an even number: laid out to the end otherwise.
    stats = _run(6, 11 + len(variant), max(CASES // 12, 4), **variant)
    assert set(stats) <= {None, "sections"}, stats


@pytest.mark.parametrize("pages", [16, 40])
def test_random_edits_of_long_reports_lay_out_as_a_full_layout(pages):
    stats = _run(pages, 100 + pages, LONG_CASES[pages])
    assert stats.get(None, 0) >= 1, stats


@pytest.fixture(scope="module")
def report16():
    blocks = report_docs.report(16)
    data = report_docs.package(blocks)
    layout, data, fonts = lay_out_docx(data)
    return blocks, layout, page_svgs(layout, data, fonts=fonts)


def _paragraph(blocks, near: int) -> int:
    return next(k for k in range(near, len(blocks)) if blocks[k].kind == "para")


def test_a_change_within_its_page_lays_out_that_page_only(report16):
    blocks, layout, svgs = report16
    k = _paragraph(blocks, len(blocks) // 2)
    edited = list(blocks)
    edited[k] = replace(edited[k], text=edited[k].text + " Yes.")
    again, again_svgs = _check(edited, layout, svgs, "one word", {})
    assert again.relaid is None
    page = again.pages_of(f"w:body/w:p[{sum(1 for b in blocks[:k + 1] if b.kind != 'table')}]")
    assert set(again.changed_pages) <= set(range(page[0], page[-1] + 2))


def test_no_change_changes_no_page(report16):
    blocks, layout, svgs = report16
    again, _ = _check(blocks, layout, svgs, "nothing", {})
    assert again.changed_pages == [] and again.relaid is None


def test_an_inserted_paragraph_renumbers_the_paths_after_it(report16):
    blocks, layout, svgs = report16
    edited = list(blocks)
    edited.insert(3, report_docs.Block("para", "A short made-up paragraph, inserted.", para_id="7000AAAA"))
    again, again_svgs = _check(edited, layout, svgs, "insert", {})
    assert 'data-docx-path="w:body/w:p[4]"' in "".join(again_svgs)


@pytest.mark.parametrize("change, reason", [
    ("header", "headers-footers"),
    ("break-count", None),
])
def test_a_header_change_is_laid_out_whole_and_a_new_page_is_not(report16, change, reason):
    blocks, layout, svgs = report16
    if change == "header":
        data = report_docs.package(blocks, header="Another made-up header")
    else:
        edited = list(blocks)
        edited.insert(_paragraph(blocks, 5), report_docs.Block("break", para_id="6000BBBB"))
        data = report_docs.package(edited)
    again, again_data, fonts = lay_out_docx(data, previous=layout)
    full, full_data, full_fonts = lay_out_docx(data, keep=False)
    assert page_svgs(again, again_data, fonts=fonts) == page_svgs(full, full_data, fonts=full_fonts)
    if change == "header":
        assert again.relaid == reason
        assert again.changed_pages == list(range(1, len(again.pages) + 1))
    else:
        # A page break adds a page: the pages after it move on, and the footer's NUMPAGES
        # changes on every page -- each kept page has its footer drawn again.
        assert again.relaid is None
        assert again.changed_pages == list(range(1, len(again.pages) + 1))
        assert len(again.pages) == len(layout.pages) + 1


def test_a_layout_from_another_options_is_laid_out_whole(report16):
    blocks, layout, _ = report16
    again = lay_out_docx(report_docs.package(blocks), ConvertOptions(substitute_fonts=False), previous=layout)[0]
    assert again.relaid == "options"


def test_a_layout_kept_nothing_is_laid_out_whole(report16):
    blocks, _, _ = report16
    data = report_docs.package(blocks)
    plain = lay_out_docx(data, keep=False)[0]
    assert plain.state is None
    again = lay_out_docx(data, previous=plain)[0]
    assert again.relaid == "no-state"
    assert again.changed_pages == list(range(1, len(again.pages) + 1))


SECTION_BREAK = ('<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" '
                 'w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>')
PAGE_FIELD = ('<w:p><w:r><w:t xml:space="preserve">This is page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/>'
              '</w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar '
              'w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
              '</w:p>')


@pytest.mark.parametrize("change, reason", [
    ("section", "sections"),
    ("styles", "styles"),
    ("page-field", "page-fields"),
    ("note", "notes"),
])
def test_each_fallback_is_a_full_layout_and_says_why(report16, change, reason):
    blocks, layout, svgs = report16
    edited = list(blocks)
    styles = ""
    if change == "section":
        edited.insert(10, report_docs.Block("raw", SECTION_BREAK))
    elif change == "styles":
        styles = ('<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/>'
                  '<w:basedOn w:val="Normal"/></w:style>')
    elif change == "page-field":
        edited.insert(10, report_docs.Block("raw", PAGE_FIELD))
    else:
        edited.insert(10, report_docs.Block("noted", "A paragraph with a footnote.", note="A made-up note."))
    data = report_docs.package(edited, styles=styles)
    again, again_svgs = _check(edited, layout, svgs, change, {}) if not styles else (None, None)
    if again is None:
        again, again_data, fonts = lay_out_docx(data, previous=layout)
        full, full_data, full_fonts = lay_out_docx(data, keep=False)
        assert page_svgs(again, again_data, fonts=fonts) == page_svgs(full, full_data, fonts=full_fonts)
    assert again.relaid == reason
    # A layout made whole from a previous one is a previous layout as good as any.
    later = list(edited)
    k = _paragraph(later, 30)
    later[k] = replace(later[k], text=later[k].text + " And more.")
    data = report_docs.package(later, styles=styles)
    final = lay_out_docx(data, previous=again)[0]
    full = lay_out_docx(data, keep=False)
    assert page_svgs(final, data, fonts=None) == page_svgs(full[0], full[1], fonts=None)
    assert final.relaid in (None, "page-fields")


def test_paths_are_renumbered_by_the_blocks_before_the_pages_kept(report16):
    # A paragraph inserted near the start and a body-level bookmark at the end: the pages
    # kept after the first are renumbered by the paragraphs before them, the bookmark aside.
    blocks, layout, svgs = report16
    edited = list(blocks)
    edited.insert(2, report_docs.Block("para", "A made-up paragraph.", para_id="7000CCCC"))
    edited.append(report_docs.Block("raw", '<w:bookmarkStart w:id="90" w:name="end"/><w:bookmarkEnd w:id="90"/>'))
    again, again_svgs = _check(edited, layout, svgs, "insert and bookmark", {})
    assert again.relaid is None and len(again.state.laid) < len(again.pages)
