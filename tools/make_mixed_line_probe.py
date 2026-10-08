#!/usr/bin/env python3
"""How tall is a line that mixes faces or sizes?

Phase 2 measured one face and one size per line.  The first real-world documents
(``tests/fixtures/samplelib/``) put a SymbolMT bullet in front of Cambria text under an
``auto`` 276 line rule, and the bullet lines came out ~2.5 px taller than Cambria's and
~0.25 px *shorter* than SymbolMT's -- neither "the tallest item" nor "largest ascent plus
largest descent".  This probe isolates it: one group per page, 22 single-line paragraphs (every group fits one page),
so the pitch is pinned to 1/21 px by the first and last baselines.

Groups vary one thing each: whether the second face arrives as a list label or as an
inline run, whether it is taller above (Symbol, 2059 over Cambria's 1946) or bigger
(Calibri 16 pt), and the line rule (``auto`` 240 / 276 / 360, ``exact``, ``atLeast``).

Throwaway: exported through ``tools/oracle.py``.  Measure with
``python tools/read_mixed_line_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

LINES = 22
CAMBRIA = {"ascii": "Cambria", "hAnsi": "Cambria", "eastAsia": "Cambria", "cs": "Cambria"}
SYMBOL = {"ascii": "Symbol", "hAnsi": "Symbol", "hint": "default"}


@dataclass(frozen=True)
class Group:
    name: str
    #: ``label`` (numbering with a Symbol bullet), ``symbol`` (inline Symbol run),
    #: ``big`` (inline Calibri run at ``big`` half points), or ``plain``.
    extra: str
    rule: str = "auto"
    line: int = 276
    big: int = 32


GROUPS = tuple(
    Group(f"{extra}-{rule}{line}", extra, rule, line)
    for rule, line in (("auto", 240), ("auto", 276), ("auto", 360), ("exact", 300), ("atLeast", 250))
    for extra in ("plain", "label", "symbol", "big")
)


def build() -> bytes:
    body = []
    for group in GROUPS:
        for index in range(LINES):
            runs = wml.run("Hxample", rFonts=CAMBRIA, sz=22)
            if group.extra == "symbol":
                runs += wml.run("", rFonts=SYMBOL, sz=22)
            elif group.extra == "big":
                runs += wml.run("Hx", rFonts={"ascii": "Calibri", "hAnsi": "Calibri"}, sz=group.big)
            props = {
                "pageBreakBefore": index == 0,
                "spacing": {"before": 0, "after": 0, "line": group.line, "lineRule": group.rule},
            }
            if group.extra == "label":
                props["numPr"] = '<w:ilvl w:val="0"/><w:numId w:val="1"/>'
            body.append(wml.paragraph(runs, mark={"rFonts": CAMBRIA, "sz": 22}, **props))
    numbering = wml.numbering_part([
        ("bullet", "", {"ind": {"left": 720, "hanging": 360}}, {"rFonts": SYMBOL}),
    ])
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": CAMBRIA, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package("".join(body), styles=styles, extra_parts=(numbering,))
