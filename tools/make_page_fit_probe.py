#!/usr/bin/env python3
"""What "fits" means at the bottom of a page: a line's box swept through the margin by
single layout units.

Phase 3 measured the horizontal budget by composing a line that ends δ units short of
the edge; this is the same measurement turned through 90°.  Every case is one page: a
plain anchor paragraph with ``w:pageBreakBefore``, filler lines, and a **candidate**
paragraph whose box ends ``slack`` layout units (1/4096 pt) above the bottom margin --
negative is below it.  Whether the candidate is drawn on its anchor's page or on the
next says whether it fitted.

The filler is composed so that every slack is exact to the unit: ``n`` Calibri 11 pt
lines (one paragraph of one run, each line an ``f`` and a ``w:br``) (55,000 units each: (1950 + 550) x 22), one Times New Roman line of ``a`` half
points (2355 a: 1825 + 443 + 87), one Calibri line of ``b`` half points (2500 b) and one
``exact`` line of ``t`` twips (``t x 204.8`` rounded half up) -- every one of them
reproduced to the pixel by the vertical model.  The page is 11,900 x 16,840 twips with
1,440 margins, so the text area is exactly 13,960 twips = 2,859,008 units.

The candidates (``FAMILIES``), each swept over ``DELTAS`` about one or two thresholds:

* ``text`` -- a plain Calibri 11 pt line; about the bottom of its pitch.
* ``after`` -- with 240 twips of space after: about the pitch's bottom and about the
  bottom of the space after (does the space after have to fit?).
* ``border`` -- with a bottom border (``w:sz`` 12, ``w:space`` 4 pt: 110 twips): about
  the pitch's bottom and the border's.
* ``multiple`` -- ``auto`` 360 (1.5 lines): the extra goes below the text; about the
  bottom of the text and of the pitch.
* ``exact`` / ``exact-small`` -- ``exact`` 400 and 200 twips (taller and shorter than
  the text); about the pitch's bottom.
* ``atleast`` -- ``atLeast`` 400 twips (the extra goes above the text).
* ``before`` -- with 240 twips of space before (above the text, counted by the stack).
* ``second`` -- the second line of a two-line paragraph (a ``w:br``), widow control off:
  about the second line's bottom.
* ``tnr`` -- Times New Roman 14.5 pt.

A second round, after the first showed that the border counts and the multiple's extra
does not: ``multiple-border``, ``multiple-after``, ``multiple-small`` (``auto`` 200, whose
text reaches below its pitch), the first and the second line of two-line paragraphs
under ``auto`` 360 and with a border (``multiple-first``, ``multiple-second``,
``border-first``, ``border-second``), and ``border-after``.

One document per compatibility setting: no ``settings.xml``, and modes 12, 14 and 15.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass

import probe_docx
import wml

WIDTH, HEIGHT, MARGIN = 11900, 16840, 1440
AVAILABLE = (HEIGHT - 2 * MARGIN) * 1024 // 5  # 13,960 twips in 1/4096 pt, exactly

CALIBRI = 2500  # (ascent + descent + lineGap) of Calibri, per half point, in units
TNR = 2355      # Times New Roman (the macOS system copy Word lays out with)
LINE = CALIBRI * 22

DELTAS = (-1024, -205, -41, -9, -3, -2, -1, 0, 1, 2, 3, 9, 41, 205, 1024)


def twips_units(twips: int) -> int:
    """A twip length in 1/4096 pt, rounded half up (Phase 2, 2.1)."""
    return (twips * 2048 + 5) // 10


@dataclass(frozen=True)
class Family:
    name: str
    #: The candidate's paragraph properties (``wml.paragraph`` keywords).
    props: tuple
    #: Height of the candidate's box from the top of its pitch to the bottom of its
    #: pitch (the slack is measured from there), and the space above its pitch.
    pitch: int
    above: int = 0
    #: Offsets of the thresholds swept, from the bottom of the pitch.
    thresholds: tuple = (0,)
    face: str = "Calibri"
    half_points: int = 22
    #: A paragraph of ``lines`` lines (split by ``w:br``, each ``pitch`` tall), of which
    #: line ``candidate`` is the one swept.
    lines: int = 1
    candidate: int = 0


def _spacing(before=0, after=0, line=240, rule="auto") -> dict:
    return {"before": before, "after": after, "line": line, "lineRule": rule}


BORDER = '<w:bottom w:val="single" w:sz="12" w:space="4" w:color="000000"/>'
#: What that border adds below the paragraph: 4 pt of space and 12/8 pt of line, in whole twips.
BORDER_UNITS = twips_units(80) + twips_units(30)
MULTIPLE = LINE * 3 // 2

FAMILIES = (
    Family("text", (), LINE),
    Family("after", (("spacing", _spacing(after=240)),), LINE, thresholds=(0, twips_units(240))),
    Family("border", (("pBdr", BORDER),), LINE, thresholds=(0, BORDER_UNITS)),
    Family("multiple", (("spacing", _spacing(line=360)),), MULTIPLE, thresholds=(-(LINE // 2), 0)),
    Family("exact", (("spacing", _spacing(line=400, rule="exact")),), twips_units(400)),
    Family("exact-small", (("spacing", _spacing(line=200, rule="exact")),), twips_units(200)),
    Family("atleast", (("spacing", _spacing(line=400, rule="atLeast")),), twips_units(400)),
    Family("before", (("spacing", _spacing(before=240)),), LINE, above=twips_units(240)),
    Family("second", (("widowControl", False),), LINE, lines=2, candidate=1),
    Family("tnr", (), TNR * 29, face="Times New Roman", half_points=29),
    # The second round, after the first showed that a bottom border counts and an auto
    # multiple's extra does not: how the two combine, where in a paragraph they apply,
    # and a multiple below one (its text reaches below its pitch).
    Family("multiple-border", (("spacing", _spacing(line=360)), ("pBdr", BORDER)), MULTIPLE,
           thresholds=(-(LINE // 2) + BORDER_UNITS, BORDER_UNITS)),
    Family("multiple-after", (("spacing", _spacing(line=360, after=240)),), MULTIPLE,
           thresholds=(-(LINE // 2), 0)),
    Family("multiple-small", (("spacing", _spacing(line=200)),), (LINE * 200 + 120) // 240,
           thresholds=(0, LINE - (LINE * 200 + 120) // 240)),
    Family("multiple-first", (("spacing", _spacing(line=360)), ("widowControl", False)), MULTIPLE,
           thresholds=(-(LINE // 2), 0), lines=2, candidate=0),
    Family("multiple-second", (("spacing", _spacing(line=360)), ("widowControl", False)), MULTIPLE,
           thresholds=(-(LINE // 2), 0), lines=2, candidate=1),
    Family("border-first", (("pBdr", BORDER), ("widowControl", False)), LINE,
           thresholds=(0, BORDER_UNITS), lines=2, candidate=0),
    Family("border-second", (("pBdr", BORDER), ("widowControl", False)), LINE,
           thresholds=(0, BORDER_UNITS), lines=2, candidate=1),
    Family("border-after", (("pBdr", BORDER), ("spacing", _spacing(after=240))), LINE,
           thresholds=(BORDER_UNITS, BORDER_UNITS + twips_units(240))),
)

#: name -> compatibilityMode (``None``: no settings part at all).
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}


@dataclass(frozen=True)
class Case:
    family: Family
    threshold: int
    delta: int

    @property
    def slack(self) -> int:
        """Units from the bottom of the candidate's pitch up to the bottom margin."""
        return self.threshold + self.delta

    @property
    def key(self) -> str:
        return f"{self.family.name} {self.threshold:+d} {self.delta:+d}"


CASES = tuple(Case(f, t, d) for f in FAMILIES for t in f.thresholds for d in DELTAS)


@functools.lru_cache(maxsize=None)
def compose(total: int) -> tuple[int, int, int, int]:
    """``(n, a, b, t)`` with ``n x 55000 + 2355 a + 2500 b + twips_units(t) == total``."""
    for n in range(total // LINE, -1, -1):
        for a in range(16, 61):
            for b in range(16, 61):
                rest = total - n * LINE - TNR * a - CALIBRI * b
                if rest < twips_units(100):
                    continue
                t = (rest * 5 + 512) // 1024
                if 100 <= t <= 2000 and twips_units(t) == rest:
                    return n, a, b, t
    raise ValueError(f"cannot compose {total} units")


def _face(name: str) -> dict:
    return {"ascii": name, "hAnsi": name, "eastAsia": name, "cs": name}


def _p(text: str, *, face="Calibri", half_points=22, **props) -> str:
    props.setdefault("spacing", _spacing())
    run_props = {"rFonts": _face(face), "sz": half_points, "szCs": half_points}
    return wml.paragraph(wml.run(text, **run_props), mark=run_props, **props)


def case_blocks(number: int, case: Case) -> list[tuple[str, str]]:
    """``(role, w:p)`` for one case's page."""
    family = case.family
    height = family.above + family.pitch * (family.candidate + 1)
    n, a, b, t = compose(AVAILABLE - LINE - height - case.slack)
    out = [("anchor", _p(f"Case {number} anchor", pageBreakBefore=True))]
    if n:
        # The n plain lines are one paragraph, broken by w:br: the same pitches as n
        # paragraphs, a tenth of the paragraphs to lay out offline.
        run_props = {"rFonts": _face("Calibri"), "sz": 22, "szCs": 22}
        lines = "<w:br/>".join('<w:t xml:space="preserve">f</w:t>' for _ in range(n))
        out.append(("filler", wml.paragraph(f"<w:r>{wml.rpr(**run_props)}{lines}</w:r>",
                                            mark=run_props, spacing=_spacing())))
    out.append(("filler", _p(f"f{number}.a", face="Times New Roman", half_points=a)))
    out.append(("filler", _p(f"f{number}.b", half_points=b)))
    out.append(("filler", _p(f"f{number}.t", spacing=_spacing(line=t, rule="exact"))))
    props = dict(family.props)
    run_props = {"rFonts": _face(family.face), "sz": family.half_points, "szCs": family.half_points}
    texts = [f"Case {number} {'cand' if k == family.candidate else f'line{k}'}" for k in range(family.lines)]
    brk = f"<w:r>{wml.rpr(**run_props)}<w:br/></w:r>"
    runs = brk.join(wml.run(text, **run_props) for text in texts)
    props.setdefault("spacing", _spacing())
    out.append(("candidate", wml.paragraph(runs, mark=run_props, **props)))
    return out


@functools.lru_cache(maxsize=1)
def blocks() -> list[tuple[str, str]]:
    out = []
    for number, case in enumerate(CASES):
        out += case_blocks(number, case)
    return out


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": _face("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": _spacing()},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    section = (
        "<w:sectPr>"
        f'<w:pgSz w:w="{WIDTH}" w:h="{HEIGHT}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    )
    return probe_docx.package("".join(p for _, p in blocks()), final_section=section, styles=styles,
                              extra_parts=extra)
