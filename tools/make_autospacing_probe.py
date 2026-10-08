#!/usr/bin/env python3
"""What ``w:beforeAutospacing`` / ``w:afterAutospacing`` make of a paragraph's spacing.

``filesamples/sample2`` (an HTML-style heading, both flags on, 100 twips stated) is
exact only if both spaces are read as 14 pt (280 twips) whatever is stated (ROADMAP.md,
"Nine more real documents", finding 3).  One document shows one stated value, one size
and one neighbour.  This probe varies each, one page per case, every page led by a
plain ``w:pageBreakBefore`` anchor with no spacing:

* ``stated`` -- each flag alone with the stated value 0, 100, 400 or 1000 twips (so the
  flag is seen above, equal to and below 280), at 8, 11, 20 and 36 pt, single, ``auto``
  1.5 and ``exact`` 40 pt: is it 14 pt, or does it scale with the size or the line?
* ``neighbour`` -- an autospaced paragraph after a space after of 0, 100, 279, 280, 281,
  400 or 1000 twips, and before a space before of the same: does it collapse with its
  neighbour as a stated 280 would?
* ``consecutive`` -- three paragraphs with both flags; ``list`` -- three list items
  with both flags (HTML drops the margins between list items); ``list edges`` -- an
  autospaced paragraph, an item of one list, an item of another, an autospaced
  paragraph; ``list stated`` -- list items with stated spacing only; ``contextual`` --
  two of a style with ``w:contextualSpacing`` and both flags;
* ``collapse`` -- space after 100 against space before 100 with neither, one or the
  other flag: does a setting that changes autospacing change how spacing collapses?
* page tops -- the document's first paragraph, a section's first, a paragraph after a
  manual break and a ``pageBreakBefore`` paragraph, each with ``beforeAutospacing``.

In four settings: no ``settings.xml``, modes 14 and 15, and mode 15 with
``w:doNotUseHTMLParagraphAutoSpacing``.  Calibri, one line per paragraph, list labels in
Calibri so a label adds nothing a line of text does not.  Measured by
``read_autospacing_probe.py`` through ``baselines.predict``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
STATED = (0, 100, 400, 1000)
SIZES = (16, 22, 40, 72)
RULES = (("auto", 240), ("auto", 360), ("exact", 800))
NEIGHBOURS = (0, 100, 279, 280, 281, 400, 1000)

#: name -> (compatibilityMode or None, legacy compat elements, write settings.xml)
SETTINGS = {
    "none": (None, "", False),
    "14": (14, "", True),
    "15": (15, "", True),
    "15-noHTML": (15, "<w:doNotUseHTMLParagraphAutoSpacing/>", True),
}


@dataclass(frozen=True)
class Probe:
    setting: str

    @property
    def name(self) -> str:
        return f"autospacing-{self.setting}"


PROBES = tuple(Probe(setting) for setting in SETTINGS)


def _p(text: str, *, hp: int = 22, before: int = 0, after: int = 0, before_auto: bool = False,
       after_auto: bool = False, rule: str = "auto", line: int = 240, **props) -> str:
    spacing = {"before": before, "after": after, "line": line, "lineRule": rule}
    if before_auto:
        spacing["beforeAutospacing"] = 1
    if after_auto:
        spacing["afterAutospacing"] = 1
    size = {"sz": hp, "szCs": hp}
    return wml.paragraph(wml.run(text, **size), mark=size, spacing=spacing, **props)


def _anchor(tag: str) -> str:
    return _p(f"Anchor {tag}", pageBreakBefore=True)


def blocks() -> list[tuple[str, str]]:
    """``(case, w:p)`` for every paragraph, in order; a page's case is its first's."""
    out: list[tuple[str, str]] = []

    def page(case: str, paragraphs: list[str], *, anchor: bool = True) -> None:
        if anchor:
            out.append((case, _anchor(case)))
        out.extend((case, p) for p in paragraphs)

    page("document start", [_p("Start auto", before=100, before_auto=True), _p("After start")],
         anchor=False)
    for s in STATED:
        for hp in SIZES:
            for rule, line in RULES:
                tag = f"{s}-{hp}-{rule}{line}"
                page("stated", [
                    _p(f"Before {tag}", hp=hp, before=s, before_auto=True, rule=rule, line=line),
                    _p(f"Plain {tag}"),
                    _p(f"After {tag}", hp=hp, after=s, after_auto=True, rule=rule, line=line),
                    _p(f"Below {tag}"),
                ])
    for n in NEIGHBOURS:
        page("neighbour", [
            _p(f"Spaced after {n}", after=n),
            _p(f"Auto before {n}", before_auto=True),
            _p(f"Auto after {n}", after_auto=True),
            _p(f"Spaced before {n}", before=n),
        ])
    page("consecutive", [_p(f"Both {k}", before=100, after=100, before_auto=True, after_auto=True)
                         for k in range(3)] + [_p("After both")])
    numbered = {"numPr": '<w:ilvl w:val="0"/><w:numId w:val="1"/>'}
    page("list", [_p(f"Item {k}", before=100, after=100, before_auto=True, after_auto=True, **numbered)
                  for k in range(3)] + [_p("After list")])
    other = {"numPr": '<w:ilvl w:val="0"/><w:numId w:val="2"/>'}
    page("list edges", [
        _p("Auto above list", before_auto=True, after_auto=True),
        _p("Edge 0", before_auto=True, after_auto=True, **numbered),
        _p("Edge 1", before_auto=True, after_auto=True, **other),
        _p("Auto below list", before_auto=True, after_auto=True),
    ])
    page("list stated", [_p(f"Stated item {k}", before=100 + 60 * k, after=100, **numbered)
                         for k in range(3)] + [_p("After stated list")])
    page("collapse", [
        _p("Collapse a", after=100), _p("Collapse b", before=100),
        _p("Collapse c", after=100, after_auto=True), _p("Collapse d", before=100),
        _p("Collapse e", after=100), _p("Collapse f", before=100, before_auto=True),
        _p("Collapse g"),
    ])
    page("contextual", [_p(f"Context {k}", before=100, after=100, before_auto=True, after_auto=True,
                           pStyle="Ctx") for k in range(2)] + [_p("After context")])
    out.append(("section start", _p("End of section", sect=probe_docx.section())))
    out.append(("section start", _p("Section auto", before=100, before_auto=True)))
    out.append(("section start", _p("After section")))
    out.append(("manual break", '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
                                ' w:lineRule="auto"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'))
    out.append(("manual break", _p("Break auto", before=100, before_auto=True)))
    out.append(("manual break", _p("After break")))
    out.append(("pageBreakBefore", _p("Pbb auto", before=100, before_auto=True, pageBreakBefore=True)))
    out.append(("pageBreakBefore", _p("After pbb")))
    return out


def kinds(probe: Probe) -> list[str]:
    return [kind for kind, _ in blocks()]


def build(probe: Probe) -> bytes:
    mode, legacy, write_settings = SETTINGS[probe.setting]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "Ctx", based_on="Normal", ppr_={"contextualSpacing": True})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    name, content_type, rel, xml = wml.numbering_part(
        [("decimal", "%1.", {"ind": {"left": 720, "hanging": 360}}, {"rFonts": FACE})])
    # A second list over the same definition, so "another list" is a different numId.
    xml = xml.replace("</w:numbering>", '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    extra = [(name, content_type, rel, xml)]
    if write_settings:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode, compat_elements=legacy))
    return probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=tuple(extra))
