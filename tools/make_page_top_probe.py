#!/usr/bin/env python3
"""What a paragraph at the top of a page keeps of its space before, by how it got there.

The model drops the space before at every page top: measured after a manual page break
(``samplelib``, compatibility mode 14) and seen once after a natural break.  Nine real
documents then showed two exceptions it cannot see (ROADMAP.md, "Nine more real
documents"): **the document's first paragraph keeps it** (every ``wordto`` document,
mode 14), and **a ``w:pageBreakBefore`` paragraph keeps it in modes 12 and 14 but not
15** (``filesamples/sample1``, re-exported with only the mode changed).  Each document
shows one side of each rule.  This probe separates the ways a paragraph comes to start
a page, in one document per compatibility setting and starting value:

* ``document start`` -- the document's first paragraph, space before ``s``;
* ``manual break`` -- the paragraph after one holding only ``w:br w:type="page"``;
* ``pageBreakBefore`` -- a paragraph with ``w:pageBreakBefore``;
* ``section start`` -- the first paragraph of a ``nextPage`` section (not the
  document's), so "first in a section" and "first in the document" come apart;
* ``natural`` -- whichever of 130 one-line filler paragraphs, each with space before,
  Word's own pagination puts at the top of a page.

The space before is swept: ``s`` over eight values for the document start, and 1..120
twips (in the eight documents of a setting together) for every other kind, so the
fractional pixel of what is kept covers [0, 1).  Every other paragraph is plain.  The
settings are no ``settings.xml`` (every earlier probe), one with no ``w:compat``, an
empty ``w:compat`` (what Word 12 writes), modes 12, 14 and 15 stated, and mode 14 with
``w:suppressSpBfAfterPgBrk``.

Calibri 11 pt, single spacing, one line per paragraph.  Measured by
``read_page_top_probe.py`` through ``baselines.predict``: which page each line is on
comes from Word; every baseline is predicted.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
#: The document start's space before, one document each.
STARTS = (1, 7, 13, 50, 100, 173, 240, 480)
#: Pages of each kind per document; eight documents cover 1..120 twips.
PER_DOCUMENT = 15
FILLERS = 130

#: name -> (compatibilityMode or None, legacy compat elements, write settings.xml, empty w:compat)
SETTINGS = {
    "none": (None, "", False, False),
    "nocompat": (None, "", True, False),
    "compat-empty": (None, "", True, True),
    "12": (12, "", True, False),
    "14": (14, "", True, False),
    "15": (15, "", True, False),
    "14-suppress": (14, "<w:suppressSpBfAfterPgBrk/>", True, False),
}


@dataclass(frozen=True)
class Probe:
    setting: str
    index: int

    @property
    def start(self) -> int:
        return STARTS[self.index]

    @property
    def name(self) -> str:
        return f"page-top-{self.setting}-{self.index}"


PROBES = tuple(Probe(setting, index) for setting in SETTINGS for index in range(len(STARTS)))


def _p(text: str, *, before: int = 0, page_break_before: bool = False, sect: str = "") -> str:
    props = {"spacing": {"before": before, "after": 0, "line": 240, "lineRule": "auto"}}
    if page_break_before:
        props["pageBreakBefore"] = True
    return wml.paragraph(wml.run(text), mark={}, sect=sect, **props)


def blocks(probe: Probe) -> list[tuple[str, str]]:
    """``(kind, w:p)`` for every paragraph, in order."""
    out = [("document start", _p(f"Start {probe.start}", before=probe.start)),
           ("plain", _p("Plain a")), ("plain", _p("Plain b"))]
    for j in range(PER_DOCUMENT):
        v = 1 + probe.index * PER_DOCUMENT + j
        out += [
            ("break", '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                      '</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'),
            ("manual break", _p(f"Break {v}", before=v)),
            ("plain", _p(f"After break {v}")),
            ("pageBreakBefore", _p(f"Pbb {v}", before=v, page_break_before=True)),
            ("plain", _p(f"After pbb {v}")),
            ("plain", _p(f"End of section {v}", sect=probe_docx.section())),
            ("section start", _p(f"Section {v}", before=v)),
            ("plain", _p(f"After section {v}")),
        ]
    for k in range(FILLERS):
        out.append(("natural", _p(f"Filler {k}", before=(k * 37) % 120 + 1)))
    return out


def kinds(probe: Probe) -> list[str]:
    return [kind for kind, _ in blocks(probe)]


def build(probe: Probe) -> bytes:
    mode, legacy, write_settings, empty = SETTINGS[probe.setting]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    extra = ()
    if write_settings:
        extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode, compat_elements=legacy,
                                   empty_compat=empty),)
    return probe_docx.package("".join(p for _, p in blocks(probe)), styles=styles, extra_parts=extra)
