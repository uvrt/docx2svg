"""What a paragraph at the top of a page keeps of its space before, without Word.

``tests/fixtures/page-top-observations.json`` holds the lines Word drew for every
document of ``tools/make_page_top_probe.py`` (recorded by ``tools/read_page_top_probe.py
--record``); the documents are regenerated from the generator.  Every baseline goes
through ``baselines.predict``, so this holds the model's own page-top rule
(``docx2svg.vertical.keeps_space_before_at_page_top``) to Word.  ROADMAP.md, "The page
top -- measured", has the tables.
"""

from __future__ import annotations

import functools
import json
from pathlib import Path

import pytest

import baselines
import make_page_top_probe as probe
import read_page_top_probe as reader

DATA = json.loads((Path(__file__).parent / "fixtures" / "page-top-observations.json").read_text(encoding="utf-8"))


#: Computed on first use, not when the module is collected: every pytest-xdist worker
#: collects every module, and only the one that runs these tests needs them (the module's
#: tests share one worker: tests/conftest.py, ``LAZY_MODULES``).
@functools.cache
def _scored():
    return reader.offline(DATA)


#: (exact, scored) per compatibility setting and kind of page top, every line of the page.
EXPECTED_BY_KIND = {
    "document start": (32, 32),
    "section start": (679, 679),
    "manual break": (240, 240),
    "natural": (713, 713),
}
#: ``pageBreakBefore`` keeps its space before in every setting but mode 15.
EXPECTED_PAGE_BREAK_BEFORE = {setting: (360, 360) for setting in probe.SETTINGS}


def test_the_recording_is_of_these_documents():
    assert set(DATA["documents"]) == {p.name for p in probe.PROBES}
    for p, results in _scored():
        # Every page top is a paragraph the generator put there, and every line is scored.
        assert all(r.status in ("exact", "miss") for r in results), p.name


@pytest.mark.parametrize("setting", sorted(probe.SETTINGS))
def test_page_tops_by_kind(setting):
    table = reader.tabulate(_scored())
    got = {kind: tuple(table[(setting, kind)]) for kind in EXPECTED_BY_KIND}
    assert got == EXPECTED_BY_KIND
    assert tuple(table[(setting, "pageBreakBefore")]) == EXPECTED_PAGE_BREAK_BEFORE[setting]


def test_the_first_paragraph_of_every_section_keeps_its_space_before():
    """The document's first paragraph and every ``nextPage`` section's, in all seven
    settings: 224 + 4,753 lines, every one exact."""
    table = reader.tabulate(_scored())
    for kind in ("document start", "section start"):
        exact = sum(v[0] for (s, k), v in table.items() if k == kind)
        scored = sum(v[1] for (s, k), v in table.items() if k == kind)
        assert exact == scored


def test_dropping_the_space_before_at_every_page_top_is_refuted(monkeypatch):
    """The model before the probe: every section start low by its space before."""
    monkeypatch.setattr(baselines, "keeps_space_before_at_page_top", lambda **_: False)
    table = reader.tabulate(reader.offline(DATA))
    assert sum(v[0] for (s, k), v in table.items() if k == "document start") == 28
    assert sum(v[0] for (s, k), v in table.items() if k == "section start") == 49
    assert sum(v[0] for (s, k), v in table.items() if k == "pageBreakBefore") == 402


def test_page_break_before_by_compatibility_mode(monkeypatch):
    """Kept below mode 15 and when no mode is stated, dropped in 15: ignoring the mode
    either way misses every ``pageBreakBefore`` page of the other side but the ones
    whose space is under half a pixel."""
    from docx2svg.vertical import keeps_space_before_at_page_top as rule

    for mode, wrong in ((None, "15"), (15, "14")):
        monkeypatch.setattr(baselines, "keeps_space_before_at_page_top",
                            lambda mode=mode, **kw: rule(**{**kw, "compatibility_mode": mode}))
        table = reader.tabulate(reader.offline(DATA))
        assert tuple(table[(wrong, "pageBreakBefore")]) == (7, 360)


def test_a_dropped_space_before_leaving_the_box_is_refuted(monkeypatch):
    """Dropped at a page top, the space before still hangs above the margin in the first
    line's box.  Leaving it out of the box misses 7 manual-break page tops per setting
    (12, 17, 36, 41, 65, 89, 113 twips: Calibri 11 pt drawn at 343, not 344) and one
    natural one."""
    from fractions import Fraction

    from docx2svg.vertical import paragraph_gap_px

    def outside(before, *, keeps, previous_after=0):
        return paragraph_gap_px(0, before) if keeps else (Fraction(0), Fraction(0))

    monkeypatch.setattr(baselines, "page_top_gap_px", outside)
    table = reader.tabulate(reader.offline(DATA))
    for setting in probe.SETTINGS:
        assert tuple(table[(setting, "manual break")]) == (233, 240)
        assert tuple(table[(setting, "natural")]) == (712, 713)


import make_page_top_collapse_probe  # noqa: E402
import read_page_top_collapse_probe  # noqa: E402


def test_a_kept_space_before_collapses_with_the_space_after_above_the_break():
    """``pageBreakBefore`` (below mode 15) and a section's first paragraph keep their
    space before at a page top, less the space after of the paragraph before the break:
    every line of every setting (ROADMAP.md, "Phase 4 -- measured")."""
    table = read_page_top_collapse_probe.tabulate(read_page_top_collapse_probe.offline())
    assert sum(v[0] for v in table.values()) == sum(v[1] for v in table.values()) == 724
    assert set(make_page_top_collapse_probe.SETTINGS) == {s for s, _ in table}
