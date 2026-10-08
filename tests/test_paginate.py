"""``docx2svg.paginate`` against Word's pages, offline (ROADMAP.md, "Phase 4 -- measured").

Each probe is regenerated from its generator, byte for byte, and scored against the pages
Word drew its paragraphs' lines on, recorded with the face metrics and advance widths
the model asked for: no Word, no font file.
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_page_fit_probe  # noqa: E402


def test_page_fit_every_case_and_every_page_top():
    """The fit rule, swept by single layout units: every case in every setting, and the
    model's page tops, all equal to Word's."""
    observed, scores = read_page_fit_probe.offline()
    for name, rule in read_page_fit_probe.HYPOTHESES.items():
        agree = sum(rule(read_page_fit_probe.probe.CASES[n], setting) == fit
                    for setting, fits in observed.items() for n, fit in fits.items())
        if name.endswith("(the model)"):
            assert agree == 1740
        else:
            assert agree < 1740, name
    assert sum(len(fits) for fits in observed.values()) == 1740
    for score in scores.values():
        assert (score.tops_matched, score.word_tops) == (636, 636)
        assert (score.paragraphs_matched, score.paragraphs) == (2610, 2610)


import read_keep_probe  # noqa: E402

from docx2svg.paginate import Rules  # noqa: E402


def test_keeps_every_case():
    """Widow and orphan control, keepLines, keepNext and its chains, pageBreakBefore and
    manual breaks: every line of every case on Word's page, in every setting."""
    compared = read_keep_probe.offline()
    assert read_keep_probe.tally(compared) == (939, 939)


def test_keeps_refuted():
    """The alternatives each fail cases Word decides otherwise (ROADMAP.md, Phase 4)."""
    scores = {name: read_keep_probe.tally(read_keep_probe.offline(rules=rules))[0] for name, rules in {
        "widow control off when unstated": Rules(widow_default=False),
        "no widow control": Rules(widow_lines=1),
        "three lines kept": Rules(widow_lines=3),
        "keepNext moves its last line in every mode": Rules(keep_next_moves="line"),
        "keepNext moves the paragraph in every mode": Rules(keep_next_moves="paragraph"),
        "a page break ends a paragraph in every mode": Rules(segment="always"),
        "a page break ends a paragraph in no mode": Rules(segment="never"),
    }.items()}
    assert scores == {
        "widow control off when unstated": 848,
        "no widow control": 751,
        "three lines kept": 798,
        "keepNext moves its last line in every mode": 921,
        "keepNext moves the paragraph in every mode": 933,
        "a page break ends a paragraph in every mode": 930,
        "a page break ends a paragraph in no mode": 936,
    }


import read_footnote_probe  # noqa: E402


def test_footnotes_reserve_their_room():
    """Mode 15: every case of the footnote probe on Word's page -- the separator's line
    and every note line, the note's spacing, widow control splitting a note, and the
    rest of a split note on the next page.  Below 15, the eight cases where Word leaves
    a line's second note to the next page (and one threshold that is not monotone) are
    not modelled."""
    assert read_footnote_probe.offline() == {
        "none": (498, 506), "12": (498, 506), "14": (498, 506), "15": (506, 506), "15-sep20": (506, 506)}


import json  # noqa: E402

import make_header_probe  # noqa: E402
import pages  # noqa: E402
import read_header_probe  # noqa: E402

from docx2svg import paginate, parse_package, vertical  # noqa: E402


def test_headers_and_footers_move_the_margins():
    """A header reaching past the top margin starts the body where it ends (its
    spacing included), a footer past the bottom one ends it there: every section's first
    baseline and line count, and every page top, as Word's."""
    for score in read_header_probe.offline().values():
        assert (score.tops_matched, score.word_tops) == (37, 37)
    data = json.loads(read_header_probe.OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    face = metrics("Calibri")
    for setting in make_header_probe.SETTINGS:
        package = make_header_probe.build(setting)
        document = parse_package(package)
        items = paginate.flow(document, advances, metrics)
        for number, case in enumerate(make_header_probe.CASES):
            top, bottom = paginate._geometry(document, number, items, True)
            lines = int((bottom - top) / vertical.natural_height_px(face, 22))
            observed = data["documents"][setting]["first"][str(number)]
            if case.name == "header, negative top":
                # The header overlaps the first body line, which Quartz merges with it.
                assert observed == [344.0, 50] and lines == 51
                continue
            assert [vertical.baseline_px(top, face, 22), lines] == observed, case.name


import read_pagination_probe  # noqa: E402


def test_every_paragraph_on_words_page():
    """Phase 4's "done when": three page sizes by three margin sets (mode 15), the same
    body in three more settings, and a document of four sections -- every page top
    and every paragraph's every line where Word put it, from the file alone."""
    for name, score in read_pagination_probe.offline().items():
        assert score.tops_matched == score.word_tops, name
        assert score.paragraphs_matched == score.paragraphs == 260, name
        assert score.line_count_differs == 0, name
        assert score.stopped is None, name


import read_picture_probe  # noqa: E402


def test_a_picture_line_is_the_pictures_height():
    """An inline picture's line: its height in whole twips, nothing below it, the mark's
    natural height as a floor below mode 15, an auto multiple's extra over the mark --
    the baseline after every picture of every setting."""
    for setting, (exact, count, wrong) in read_picture_probe.offline().items():
        assert (exact, count) == (200, 200), (setting, wrong[:3])


import read_section_probe  # noqa: E402


def test_sections_start_pages_as_words_do():
    """``oddPage`` / ``evenPage`` add a blank page when the parity is wrong; a continuous
    section that changes the page size starts a new page: every section's every line on
    Word's page, in three settings."""
    for score in read_section_probe.offline().values():
        assert (score.tops_matched, score.word_tops, score.paragraphs_matched, score.paragraphs) == (15, 15, 13, 13)
