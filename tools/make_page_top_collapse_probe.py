#!/usr/bin/env python3
"""Whether a space before kept at a page top collapses with the space after above the break.

``make_page_top_probe.py`` found that a ``pageBreakBefore`` paragraph (below mode 15) and
a section's first paragraph keep their space before at the top of a page -- measured
after paragraphs with no space after.  The pagination probe then drew such a paragraph
40 twips low, not 240, after one with 200 after: kept, but collapsed against the space
after at the bottom of the page before, as if no break were between them.  This probe
separates the two: the paragraph before the break has ``after`` twips of space after
(0, 40, 120, 200, 360), the page top ``before`` (1, 50, 100, 150, 240, 480), in both
ways of keeping it (``pageBreakBefore``, and the first paragraph of a ``nextPage``
section), in each compatibility setting (none, 12, 14, 15).  Calibri 11 pt, one line per
paragraph; scored through ``baselines.predict``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
AFTERS = (0, 40, 120, 200, 360)
BEFORES = (1, 50, 100, 150, 240, 480)
KINDS = ("pageBreakBefore", "section start")
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}


@dataclass(frozen=True)
class Case:
    kind: str
    after: int
    before: int


CASES = tuple(Case(kind, a, b) for kind in KINDS for a in AFTERS for b in BEFORES)


def _p(text: str, *, before: int = 0, after: int = 0, sect: str = "", **props) -> str:
    return wml.paragraph(wml.run(text), mark={}, sect=sect,
                         spacing={"before": before, "after": after, "line": 240, "lineRule": "auto"}, **props)


def blocks() -> list[tuple[str, str]]:
    """``(kind, w:p)`` for every paragraph: ``plain``, ``above`` (before the break) or
    the case's kind (the page top)."""
    out = [("plain", _p("Start"))]
    for number, case in enumerate(CASES):
        sect = probe_docx.section() if case.kind == "section start" else ""
        out.append(("above", _p(f"Above {number}", after=case.after, sect=sect)))
        out.append((case.kind, _p(f"Top {number}", before=case.before,
                                  **({"pageBreakBefore": True} if case.kind == "pageBreakBefore" else {}))))
        out.append(("plain", _p(f"Below {number}")))
    return out


def kinds() -> list[str]:
    return [kind for kind, _ in blocks()]


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=extra)
