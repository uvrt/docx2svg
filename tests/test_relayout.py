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
Every kind of edit is also made on its own, one after another on one report.

**Two sizes.**  A pull request's CI runs a small seeded set (about a minute per leg): a few
edits per test, every edit kind, every fallback.  ``RELAYOUT_FULL=1`` (an environment
variable) runs the full set -- 96 edits of the small reports, the 16- and the 40-page
report, every fallback with more edits after it -- as the nightly workflow
(``.github/workflows/relayout-full.yml``) does on ubuntu with the newest Python::

    RELAYOUT_FULL=1 python -m pytest -q tests/test_relayout.py

The assertions are the same in both; only the number of edits differs (:data:`SIZES`).
``RELAYOUT_CASES`` and ``RELAYOUT_LONG_CASES`` set the small reports' and the long reports'
numbers outright, in either size.

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

#: Whether the full set runs (``RELAYOUT_FULL=1``; the nightly workflow) or a pull request's.
FULL = os.environ.get("RELAYOUT_FULL", "") not in ("", "0")
#: How many edits each test makes, in the full set and in a pull request's: of the small
#: reports (over the four seeds), of the reports with footnotes and with sections (each
#: variant), of the 16- and the 40-page report, of the report with fields, and of the
#: reports with contents, captions and cross-references (each variant; and of 16 and 40
#: pages).
SIZES = {
    True: {"small": 96, "notes": 12, "sections": 8, 16: 6, 40: 2, "fields": 6, "contents": 24,
           "contents-16": 8, "contents-40": 3, "styleref": 32, "styleref-wrap": 16},
    False: {"small": 8, "notes": 2, "sections": 2, 16: 1, 40: 0, "fields": 2, "contents": 3,
            "contents-16": 0, "contents-40": 0, "styleref": 6, "styleref-wrap": 3},
}[FULL]
#: Edits of the small reports; ``RELAYOUT_CASES`` sets another number.
CASES = int(os.environ.get("RELAYOUT_CASES", SIZES["small"]))
#: Edits of the 16-page and the 40-page report; ``RELAYOUT_LONG_CASES`` sets both.
LONG_CASES = {pages: int(os.environ.get("RELAYOUT_LONG_CASES", SIZES[pages])) for pages in (16, 40)}
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


def _run(pages: int, seed: int, cases: int, *, notes: bool = False, sections: bool = False,
         contents: dict | None = None, kinds: tuple | None = None, **package) -> dict:
    rng = random.Random(seed)
    stats: dict = {}
    done = 0
    while done < cases:
        blocks = report_docs.report(pages, seed=rng.randrange(1 << 30), notes=notes, sections=sections,
                                    **(contents or {}))
        layout, data, fonts = lay_out_docx(report_docs.package(blocks, **package))
        svgs = page_svgs(layout, data, fonts=fonts)
        for _ in range(min(CHAIN, cases - done)):
            blocks, what = report_docs.edit(blocks, rng, kinds=kinds)
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
    stats = _run(6, 7, SIZES["notes"], notes=True)
    assert set(stats) <= {None, "notes", "note-references"}, stats


@pytest.mark.parametrize("variant", [
    {"sections": True},
    {"even_odd": True, "title_page": True},
    {"sections": True, "even_odd": True},
])
def test_random_edits_of_reports_with_sections_and_their_own_headers(variant):
    # Pages that move on keep their headers only where no section starts among them and,
    # where even pages have their own, by an even number: laid out to the end otherwise.
    stats = _run(6, 11 + len(variant), SIZES["sections"], **variant)
    assert set(stats) <= {None, "sections"}, stats


@pytest.mark.parametrize("pages", [16, 40])
def test_random_edits_of_long_reports_lay_out_as_a_full_layout(pages):
    if not LONG_CASES[pages]:
        pytest.skip(f"the {pages}-page report is in the full set only: RELAYOUT_FULL=1")
    stats = _run(pages, 100 + pages, LONG_CASES[pages])
    assert stats.get(None, 0) >= 1, stats


def test_every_kind_of_edit_lays_out_as_a_full_layout():
    # One edit of each kind in turn, each layout made from the one before: whatever the
    # number of random edits, every kind is held to the property.
    rng = random.Random(23)
    blocks = report_docs.report(4, seed=23)
    layout, data, fonts = lay_out_docx(report_docs.package(blocks))
    svgs = page_svgs(layout, data, fonts=fonts)
    stats: dict = {}
    for kind in report_docs.EDIT_KINDS:
        blocks, what = report_docs.edit(blocks, rng, kind)
        assert what.split()[0] == kind, what
        layout, svgs = _check(blocks, layout, svgs, f"each kind, {what}", stats)
    assert stats.get(None, 0) >= sum(stats.values()) // 3, stats


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
    else:
        edited.insert(10, report_docs.Block("noted", "A paragraph with a footnote.", note="A made-up note."))
    data = report_docs.package(edited, styles=styles)
    again, again_svgs = _check(edited, layout, svgs, change, {}) if not styles else (None, None)
    if again is None:
        again, again_data, fonts = lay_out_docx(data, previous=layout)
        full, full_data, full_fonts = lay_out_docx(data, keep=False)
        assert page_svgs(again, again_data, fonts=fonts) == page_svgs(full, full_data, fonts=full_fonts)
    assert again.relaid == reason
    assert again.relaid_cause["reason"] == reason
    # A layout made whole from a previous one is a previous layout as good as any.
    later = list(edited)
    k = _paragraph(later, 30)
    later[k] = replace(later[k], text=later[k].text + " And more.")
    data = report_docs.package(later, styles=styles)
    final = lay_out_docx(data, previous=again)[0]
    full = lay_out_docx(data, keep=False)
    assert page_svgs(final, data, fonts=None) == page_svgs(full[0], full[1], fonts=None)
    assert final.relaid is None and final.relaid_cause is None


def test_paths_are_renumbered_by_the_blocks_before_the_pages_kept(report16):
    # A paragraph inserted near the start and a body-level bookmark at the end: the pages
    # kept after the first are renumbered by the paragraphs before them, the bookmark aside.
    blocks, layout, svgs = report16
    edited = list(blocks)
    edited.insert(2, report_docs.Block("para", "A made-up paragraph.", para_id="7000CCCC"))
    edited.append(report_docs.Block("raw", '<w:bookmarkStart w:id="90" w:name="end"/><w:bookmarkEnd w:id="90"/>'))
    again, again_svgs = _check(edited, layout, svgs, "insert and bookmark", {})
    assert again.relaid is None and len(again.state.laid) < len(again.pages)


FIELDS = [("PAGE", "1"), ("SEQ Figure", "1"), ("NUMPAGES", "1")]


def _with_fields(blocks):
    """``blocks`` with a paragraph of each field in :data:`FIELDS` well into the report,
    each with a cached result a layout replaces."""
    out = list(blocks)
    for k, (instruction, cached) in enumerate(FIELDS):
        out.insert(len(out) * (k + 2) // (len(FIELDS) + 2), report_docs.Block("raw", PAGE_FIELD.replace(
            " PAGE ", f" {instruction} ").replace("<w:t>1</w:t>", f"<w:t>{cached}</w:t>")))
    return out


@pytest.mark.parametrize("name", ["fields", "samplelib/sample-long.docx", "wordto/sample-with-table.docx",
                                  "wordto/sample-with-images.docx", "style-document.docx", "layout-sweep.docx"])
def test_a_layout_that_keeps_its_state_draws_what_one_that_does_not_does(name):
    # What a layout keeps -- the flow's measurements among it -- never changes what it draws:
    # page fields laid out again with their values, SEQ counted, pictures, tables.
    from pathlib import Path

    if name == "fields":
        data = report_docs.package(_with_fields(report_docs.report(8)))
    else:
        data = (Path(__file__).resolve().parent / "fixtures" / name).read_bytes()
    kept, kept_data, kept_fonts = lay_out_docx(data)
    plain, plain_data, plain_fonts = lay_out_docx(data, keep=False)
    assert page_svgs(kept, kept_data, fonts=kept_fonts) == page_svgs(plain, plain_data, fonts=plain_fonts)
    assert kept.state is not None and plain.state is None


def test_edits_of_a_report_with_fields_lay_out_as_a_full_layout():
    blocks = _with_fields(report_docs.report(8))
    layout, data, fonts = lay_out_docx(report_docs.package(blocks))
    svgs = page_svgs(layout, data, fonts=fonts)
    rng = random.Random(5)
    stats: dict = {}
    for _ in range(SIZES["fields"]):
        blocks, what = report_docs.edit(blocks, rng)
        layout, svgs = _check(blocks, layout, svgs, f"fields, {what}", stats)
    # PAGE, SEQ and NUMPAGES in the body: laid out from the edit (round 20).
    assert set(stats) == {None}, stats


# -- round 20: contents, page-number fields, captions and cross-references ----------------------

#: Edits of the reports with contents: the usual mix, and captions and lines of the contents
#: inserted and deleted.
CONTENTS_KINDS = ("text", "text", "cell", "insert", "delete", "row-insert", "row-delete", "break", "unbreak",
                  "grow", *report_docs.FIELD_EDIT_KINDS)
#: The reports with contents: a table of contents with a page break after it and without
#: one; with captions numbered by SEQ, REF to them, PAGEREF and PAGE in the body; both.
CONTENTS = [
    {"toc": True},
    {"toc": True, "toc_break": False},
    {"fields": True},
    {"toc": True, "fields": True},
]


@pytest.mark.parametrize("contents", CONTENTS, ids=["contents", "contents-no-break", "fields", "both"])
def test_random_edits_of_reports_with_contents_lay_out_from_the_edit(contents):
    # Every edit -- a heading moved across a page, lines of the contents added or removed, a
    # caption added or removed (the numbers after it and the REF to them change), a page
    # added or removed -- is laid out from where it reaches, as a full layout lays it out.
    stats = _run(6, 31 + len(contents) * 7 + contents.get("toc_break", True), SIZES["contents"],
                 contents=contents, kinds=CONTENTS_KINDS)
    assert set(stats) == {None}, stats


@pytest.mark.parametrize("pages", [16, 40])
def test_random_edits_of_long_reports_with_contents(pages):
    cases = SIZES[f"contents-{pages}"]
    if not cases:
        pytest.skip(f"the {pages}-page report with contents is in the full set only: RELAYOUT_FULL=1")
    stats = _run(pages, 200 + pages, cases, contents={"toc": True, "fields": pages == 16}, kinds=CONTENTS_KINDS)
    assert set(stats) == {None}, stats


def _text(page) -> str:
    return "\n".join("".join("".join(span.chars) for span in line.spans) for line in page.lines if line.story is None)


@pytest.fixture(scope="module")
def contents16():
    blocks = report_docs.report(16, toc=True)
    data = report_docs.package(blocks)
    layout, data, fonts = lay_out_docx(data)
    return blocks, layout, page_svgs(layout, data, fonts=fonts)


def _heading_page(layout, block) -> int:
    return layout.pages_of(block.para_id)[0]


def test_a_paragraph_that_keeps_the_headings_on_their_pages_redraws_its_page_only(contents16):
    # The production case: a template with contents on page 2, one paragraph lengthened.
    blocks, layout, svgs = contents16
    k = next(k for k in range(len(blocks) // 2, len(blocks)) if blocks[k].kind == "para")
    page = layout.pages_of(blocks[k].para_id)
    edited = list(blocks)
    edited[k] = replace(edited[k], text=edited[k].text + " One more.")
    again, _ = _check(edited, layout, svgs, "one word", {})
    assert again.relaid is None and again.relaid_cause is None
    assert again.changed_pages == page
    assert len(again.state.laid) == 1


def test_a_heading_moved_to_the_next_page_redraws_the_contents_page(contents16):
    blocks, layout, svgs = contents16
    contents = {block.target for block in blocks if block.kind == "toc"}
    # A heading of the contents, mid-report, and the paragraph before it, grown by about
    # thirty lines: the heading moves on to the next page.
    h = next(h for h in range(len(blocks) // 2, len(blocks)) if blocks[h].kind == "heading"
             and blocks[h].mark in contents and blocks[h - 1].kind == "para")
    before = _heading_page(layout, blocks[h])
    edited = list(blocks)
    edited[h - 1] = replace(blocks[h - 1], text=blocks[h - 1].text + " Lantern meadow river stone." * 110)
    again, again_svgs = _check(edited, layout, svgs, "heading moved", {})
    assert again.relaid is None
    assert _heading_page(again, edited[h]) == before + 1
    # The contents page shows the heading's new page number; the pages after the heading's
    # chapter are the previous layout's (the page break before the next chapter).
    assert 2 in again.changed_pages
    line = next(line for line in _text(again.pages[1]).split("\n") if line.startswith(edited[h].text))
    # (the dotted leader is drawn after the line's words and number)
    assert line.rstrip(".") == f"{edited[h].text}{before + 1}", line
    # Laid out again: the contents page, and the heading's pages up to the next chapter's page
    # break (the pages after it are the previous layout's, moved on, their footers drawn again).
    fresh = sorted(k + 1 for k in again.state.fresh)
    after_break = next(again.pages_of(edited[b + 1].para_id)[0] for b in range(h, len(edited))
                       if edited[b].kind == "break")
    assert fresh[0] == 2 and before - 1 <= fresh[1] and fresh[-1] < after_break, fresh
    assert len(again.state.laid) > len(fresh)


def test_a_page_break_after_the_edit_is_where_the_pages_line_up_again(contents16):
    # A paragraph before a chapter's page break grown by a page: the pages after the break
    # are the previous layout's moved on by one, their footers drawn again (Page N of M).
    blocks, layout, svgs = contents16
    b = next(b for b in range(len(blocks) // 2, len(blocks)) if blocks[b].kind == "break" and blocks[b - 1].kind == "para")
    edited = list(blocks)
    edited[b - 1] = replace(blocks[b - 1], text=blocks[b - 1].text + " Lantern meadow river stone harbour." * 80)
    again, _ = _check(edited, layout, svgs, "a page more before a break", {})
    assert again.relaid is None
    assert len(again.pages) == len(layout.pages) + 1
    after_break = again.pages_of(blocks[b + 1].para_id)[0]
    # Laid out again: the contents page (the headings after the break moved on) and the
    # pages up to the break; none after it.
    assert again.state.laid.stop <= after_break
    assert {k + 1 for k in again.state.redrawn} >= set(range(after_break, len(again.pages) + 1))


def test_the_contents_without_a_page_break_after_them(contents16):
    blocks = report_docs.report(8, toc=True, toc_break=False)
    layout, data, fonts = lay_out_docx(report_docs.package(blocks))
    svgs = page_svgs(layout, data, fonts=fonts)
    edited = list(blocks)
    # A line of the contents more: the report after it moves down.
    t = next(t for t, block in enumerate(blocks) if block.kind == "toc" and not block.opens)
    heading = next(block for block in blocks if block.kind == "heading" and block.mark)
    edited.insert(t, report_docs.Block("toc", heading.text, para_id="4000AAAA", target=heading.mark))
    again, _ = _check(edited, layout, svgs, "a line of the contents more", {})
    assert again.relaid is None


def test_a_caption_inserted_numbers_the_captions_after_it_again(contents16):
    blocks = report_docs.report(6, fields=True)
    layout, data, fonts = lay_out_docx(report_docs.package(blocks))
    svgs = page_svgs(layout, data, fonts=fonts)
    k = next(k for k, block in enumerate(blocks) if block.kind == "caption")
    edited = list(blocks)
    edited.insert(k, report_docs.Block("caption", "A made-up chart", para_id="5000AAAA", mark="_Ref9999"))
    again, again_svgs = _check(edited, layout, svgs, "caption inserted", {})
    assert again.relaid is None
    texts = "\n".join(_text(page) for page in again.pages)
    count = sum(1 for block in edited if block.kind == "caption")
    assert f"Table {count}:" in texts and f"Table {count + 1}:" not in texts


def test_editing_the_line_that_opens_the_contents_says_which_field(contents16):
    # The contents' first line opens the TOC field and its closing paragraph closes it: an
    # edit of either is laid out whole, and the cause names the element and the field.
    blocks, layout, svgs = contents16
    t = next(t for t, block in enumerate(blocks) if block.kind == "toc" and block.opens)
    edited = list(blocks)
    edited[t] = replace(blocks[t], text="Another made-up heading")
    again, _ = _check(edited, layout, svgs, "the opening line", {})
    assert again.relaid == "fields-open"
    cause = again.relaid_cause
    assert cause["reason"] == "fields-open"
    assert cause["at"] == f"w:body/w:p[{t + 1}]" and cause["para_id"] == blocks[t].para_id
    assert cause["field"].startswith("TOC \\o")


def test_a_part_change_says_which_part(report16):
    blocks, layout, _ = report16
    again = lay_out_docx(report_docs.package(blocks, header="Another made-up header"), previous=layout)[0]
    assert again.relaid_cause == {"reason": "headers-footers", "part": "word/header1.xml"}
    plain = lay_out_docx(report_docs.package(blocks))[0]
    assert plain.relaid_cause is None and plain.relaid is None


def test_the_last_block_deleted_from_a_report_with_contents():
    # The edit is empty in the new document and at its end: the stretch it makes reaches the
    # previous document's end (a deleted last paragraph once fell back to "body").
    blocks = report_docs.report(4, toc=True, fields=True)
    layout, data, fonts = lay_out_docx(report_docs.package(blocks))
    svgs = page_svgs(layout, data, fonts=fonts)
    again, _ = _check(blocks[:-1], layout, svgs, "the last block deleted", {})
    assert again.relaid is None


# -- round 21: STYLEREF in the header ---------------------------------------------------------

#: The header's fields: the page's first heading and its last (``\\l``).
STYLEREF = ('STYLEREF "heading 1"', "STYLEREF 1 \\l")
#: Edits of a report whose header shows its headings: the usual mix, headings' words
#: changed, and paragraphs made headings and headings paragraphs.
STYLEREF_KINDS = ("text", "insert", "delete", "break", "unbreak", "cell", "row-insert",
                  *report_docs.STYLE_EDIT_KINDS, *report_docs.STYLE_EDIT_KINDS)


def _headings(page) -> list[str]:
    return [line for line in _text_of(page, "header").split("\n") if line]


def _text_of(page, story) -> str:
    return "\n".join("".join("".join(span.chars) for span in line.spans) for line in page.lines if line.story == story)


def test_random_edits_of_reports_with_a_styleref_header_lay_out_from_the_edit():
    # A heading moved to another page, its words changed, a paragraph made a heading or a
    # heading a paragraph: the header of every page showing it is drawn again, the rest kept.
    stats = _run(6, 2101, SIZES["styleref"], kinds=STYLEREF_KINDS, styleref=STYLEREF[:1], header="")
    assert set(stats) == {None}, stats


def test_a_styleref_header_that_wraps_on_some_pages_lays_out_as_a_full_layout():
    # A long title: a long heading's text wraps the header onto a second line on its pages,
    # which moves their body down (measured): laid out whole, and the same as a full layout.
    stats = _run(6, 2102, SIZES["styleref-wrap"], kinds=("grow", "text", "restyle"), styleref=STYLEREF[:1],
                 title="Made-up quarterly report of the river stone lantern orchard meadow")
    assert set(stats) <= {None, "styleref-height"}, stats
    if FULL:
        assert "styleref-height" in stats, stats


def test_a_page_whose_header_wraps_has_its_body_lower():
    # Measured (round 21): Word lays a page out under its own header, a STYLEREF's text on
    # that page included -- a header taller there than its margin moves that page's body only.
    blocks = report_docs.report(6)
    k = next(k for k, block in enumerate(blocks) if block.kind == "heading" and k > len(blocks) // 2)
    # Four lines in the header: past the top margin.
    blocks[k] = replace(blocks[k], text="Part made up with words enough to run the header on past its margin "
                                        "on every page that shows this heading" * 3)
    package = {"styleref": STYLEREF[:1]}
    layout, data, fonts = lay_out_docx(report_docs.package(blocks, **package))
    tops = [page.column[2] for page in layout.pages]
    wrapped = [k for k, page in enumerate(layout.pages) if len([line for line in page.lines
                                                                if line.story == "header"]) > 1]
    assert wrapped and all(tops[k] > min(tops) for k in wrapped)
    assert all(tops[k] == min(tops) for k in range(len(tops)) if k not in wrapped)
    svgs = page_svgs(layout, data, fonts=fonts)
    edited = list(blocks)
    edited[1] = replace(blocks[1], text=blocks[1].text + " One more.")
    again, _ = _check(edited, layout, svgs, "a word more", {}, package)
    assert again.relaid == "styleref-height"


def test_the_header_shows_each_pages_heading_and_follows_an_edit_of_it():
    blocks = report_docs.report(8)
    layout, data, fonts = lay_out_docx(report_docs.package(blocks, styleref=STYLEREF))
    svgs = page_svgs(layout, data, fonts=fonts)
    headings = [(layout.pages_of(block.para_id), block.text) for block in blocks if block.kind == "heading"]
    for index, page in enumerate(layout.pages):
        on = [text for pages, text in headings if index + 1 in pages]
        before = [text for pages, text in headings if pages and pages[0] < index + 1]
        first = on[0] if on else before[-1]
        last = on[-1] if on else before[-1]
        assert _text_of(page, "header") == f"Made-up quarterly report | {first} | {last}", index
    # A heading's words changed: its pages and those after it up to the next heading's are
    # drawn again; nothing is laid out but its own page.
    h = next(k for k in range(len(blocks) // 2, len(blocks)) if blocks[k].kind == "heading"
             and blocks[k].text in _text_of(layout.pages[layout.pages_of(blocks[k].para_id)[0] - 1], "header"))
    edited = list(blocks)
    edited[h] = replace(blocks[h], text="Part made up anew")
    again, _ = _check(edited, layout, svgs, "heading words", {}, {"styleref": STYLEREF})
    assert again.relaid is None
    shown = [k + 1 for k, page in enumerate(again.pages) if "Part made up anew" in _text_of(page, "header")]
    assert shown and set(shown) <= set(again.changed_pages)
