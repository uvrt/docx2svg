#!/usr/bin/env python3
"""Where the baseline sits in a line that carries space the text does not use.

Phase 2 reproduced every baseline of lines whose pitch is all text (single spacing,
``exact``, ``atLeast``, ``auto`` below one) and left three 1 px residuals, all about
space *around* the text: ``auto`` multiples above one (the extra goes below the text), a
paragraph with a bottom border (the border goes below), and lines at or near a rounding
tie after paragraph spacing.  The documents said the three might be one: Word appears to
round a line's baseline inside a **line box** that includes the paragraph's space
before (on its first line), its space after and bottom border (on its last line) and the
multiple's extra -- not inside the text's pitch alone.

This probe varies each of those, one at a time, over enough values that the fractional
pixel of every quantity covers [0, 1):

* ``after``   -- ``w:spacing/@w:after`` 1..120 twips on single-spaced lines;
* ``before``  -- ``w:spacing/@w:before`` 1..120 twips;
* ``border``  -- ``w:pBdr/w:bottom`` with ``w:space`` 0..31 pt and four widths, each
  bordered paragraph followed by a plain one (identical borders on consecutive
  paragraphs would merge into one group, which is a different measurement);
* ``exact``   -- ``w:line`` 240..480 twips ``exact``: Phase 2 could not tell how far the
  rounding slack moves an ``exact`` baseline (lambda anywhere in [0, 2)), and
  ``style-document.docx`` says it is at most 1/2;
* ``multiple`` -- ``auto`` ``w:line`` 241..480 in steps of 1, two paragraphs each: the
  probe Phase 2 named for its unsettled rule (Times New Roman 14.5 pt first);
* ``multiple-after`` -- ``auto`` 276 (every Word 2010 Normal) with space after 1..120;
* ``combo-after`` / ``combo-before`` / ``combo-both`` -- every other line rule (``exact``,
  ``atLeast`` above the text, ``auto`` below and above one) with space after, space
  before, or both (1..48 twips), so no rule is only ever seen with a bare line box.

Every paragraph is one line of ``Hx`` whose mark has the same face and size.  Each page
begins with a plain single-spaced anchor paragraph carrying ``w:pageBreakBefore`` and no
spacing, so no page starts with space before (whose fate at a page top is only measured
after a manual break) and every page restarts at the top margin; the pages are filled
to at most 85% so nothing breaks naturally.

Exported through ``tools/oracle.py``; measured by ``read_line_box_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

#: Usable column of an A4 page with 1-inch margins, in device px.
_COLUMN_PX = (16838 - 2 * 1440) / 4.8
_FILL = 0.85
#: Border widths (eighths of a point) the border sweep cycles through.
BORDER_WIDTHS = (2, 6, 12, 27)


@dataclass(frozen=True)
class Para:
    rule: str = "auto"
    line: int = 240
    before: int = 0
    after: int = 0
    #: (width in eighths of a point, space in points), or ``None``.
    border: tuple[int, int] | None = None
    #: The page's first paragraph: plain, and starts the page.
    anchor: bool = False


@dataclass(frozen=True)
class Probe:
    sweep: str
    face: str
    half_points: int

    @property
    def name(self) -> str:
        return f"line-box-{self.sweep}-{self.face.replace(' ', '').lower()}-{self.half_points}"


PROBES = (
    tuple(Probe("after", face, hp) for face in ("Calibri", "Times New Roman", "Arial", "Cambria")
          for hp in (13, 22, 29, 40))
    + tuple(Probe("before", face, hp) for face in ("Calibri", "Times New Roman", "Cambria")
            for hp in (13, 22, 29, 40))
    + tuple(Probe("border", face, hp) for face in ("Calibri", "Times New Roman")
            for hp in (22, 29, 40))
    + tuple(Probe("exact", face, hp) for face, hp in
            (("Calibri", 22), ("Courier New", 20), ("Times New Roman", 22), ("Calibri", 13)))
    + tuple(Probe("multiple", face, hp) for face, hp in
            (("Times New Roman", 29), ("Arial", 29), ("Cambria", 29), ("Calibri", 40),
             ("Calibri", 22), ("Cambria", 22)))
    + tuple(Probe("multiple-after", face, hp) for face, hp in (("Cambria", 22), ("Calibri", 22)))
    + tuple(Probe(f"combo-{mode}", face, hp) for mode in ("after", "before", "both")
            for face, hp in (("Calibri", 22), ("Times New Roman", 29)))
)

#: The line rules the ``combo`` sweeps pair with paragraph spacing.
COMBO_RULES = (("exact", 263), ("exact", 290), ("exact", 333), ("atLeast", 350),
               ("atLeast", 411), ("auto", 200), ("auto", 220), ("auto", 276), ("auto", 360))


def paragraphs(probe: Probe) -> list[Para]:
    """The swept paragraphs, in order, without anchors."""
    if probe.sweep == "after":
        return [Para(after=v) for v in range(1, 121)]
    if probe.sweep == "before":
        return [Para(before=v) for v in range(1, 121)]
    if probe.sweep == "border":
        out = []
        for space in range(32):
            for width in BORDER_WIDTHS:
                out += [Para(border=(width, space)), Para()]
        return out
    if probe.sweep == "exact":
        return [Para("exact", v) for v in range(240, 481)]
    if probe.sweep == "multiple":
        return [Para("auto", v) for v in range(241, 481) for _ in range(2)]
    if probe.sweep == "multiple-after":
        return [Para("auto", 276, after=v) for v in range(1, 121)]
    if probe.sweep.startswith("combo-"):
        mode = probe.sweep[len("combo-"):]
        return [Para(rule, line, before=v if mode in ("before", "both") else 0,
                     after=v if mode in ("after", "both") else 0)
                for rule, line in COMBO_RULES for v in range(1, 49)]
    raise ValueError(probe.sweep)


def _estimate_px(probe: Probe, para: Para) -> float:
    natural = 1.3 * probe.half_points / 2 * 300 / 72
    if para.rule == "auto":
        pitch = natural * para.line / 240
    elif para.rule == "exact":
        pitch = para.line / 4.8
    else:
        pitch = max(natural, para.line / 4.8)
    extra = (para.before + para.after) / 4.8
    if para.border:
        width, space = para.border
        extra += (space + width / 8) * 300 / 72
    return pitch + extra


def pages(probe: Probe) -> list[list[Para]]:
    """Paragraphs per page, each page led by an anchor; the reader relies on this split."""
    out: list[list[Para]] = []
    current: list[Para] = []
    used = 0.0
    swept = paragraphs(probe)
    index = 0
    while index < len(swept):
        if not current:
            current = [Para(anchor=True)]
            used = _estimate_px(probe, current[0])
        # Keep a bordered paragraph and the plain one after it on the same page.
        take = 2 if swept[index].border else 1
        height = sum(_estimate_px(probe, p) for p in swept[index:index + take])
        if used + height > _FILL * _COLUMN_PX and len(current) > 1:
            out.append(current)
            current = []
            continue
        current += swept[index:index + take]
        used += height
        index += take
    if current:
        out.append(current)
    return out


def _fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def build(probe: Probe) -> bytes:
    rpr = {"rFonts": _fonts(probe.face), "sz": probe.half_points, "szCs": probe.half_points}
    body = []
    for page in pages(probe):
        for para in page:
            props = {
                "spacing": {"before": para.before, "after": para.after, "line": para.line,
                            "lineRule": para.rule},
            }
            if para.anchor:
                props["pageBreakBefore"] = True
            if para.border:
                width, space = para.border
                props["pBdr"] = (f'<w:bottom w:val="single" w:sz="{width}" w:space="{space}"'
                                 ' w:color="000000"/>')
            body.append(wml.paragraph(wml.run("Hx", **rpr), mark=rpr, **props))
    return probe_docx.package("".join(body), face=probe.face, half_points=probe.half_points)
