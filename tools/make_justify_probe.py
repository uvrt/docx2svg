#!/usr/bin/env python3
"""How far past the edge a justified line may run in compatibility mode 15: the δ sweep.

Phase 3.1 measured the line budget as exact, inclusive and without slack, in Word's unit
(1/4096 pt), and it holds in every mode for text that is not justified.  In mode 15 Word
fits more words on a *justified* line (Phase 4.6, 5.10: 65 paragraphs of the pagination
probe broke into fewer lines; the render probe's lines ran up to 24.8% of their space
width past the column): it compresses the line's spaces.  This probe measures by how
much, the way 3.1 measured the budget.

Every case is one justified (``both``) paragraph whose first line under test is ``k``
copies of a fixed word, spaced, then a composed word ``W2`` -- so the line has ``k``
spaces before ``W2`` -- then more words, so the line is never the paragraph's last.  The
column is chosen, in multiples of 5 twips, so that the line through ``W2`` is **δ units
wider than the budget**: δ = ``round(f x S)`` for ``S`` the width of its ``k`` spaces and
``f`` swept from 0 to 0.40 by 0.025 (the nearest composable δ to it: its actual δ is
recorded).  Under the
non-compressing budget the line never holds ``W2`` (δ > 0 but for f = 0); how far Word
lets it is the measurement.

Faces Calibri, Times New Roman, Arial and Georgia at 11 and 14.5 pt; ``k`` = 1, 2, 3,
5, 8 and 12.  Documents: mode 15 justified (the measurement), and three controls that
must not compress -- mode 15 left-aligned, mode 14 justified and no ``settings.xml``
justified.  A second round (:func:`fine_cases`, documents ``fine-15-both`` and
``fine-15-left``) places cases at single units about the edge the first found, and asks
which spaces count.  Reader: ``read_justify_probe.py``.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass

import make_wrap_budget_probe as budget
import probe_docx as w
import wml

FACES = ("Calibri", "Times New Roman", "Arial", "Georgia")
SIZES = (22, 29)
SPACES = (1, 2, 3, 5, 8, 12)
FRACTIONS = tuple(i / 40 for i in range(17))  # 0 .. 0.40
W1 = "Hnio"
TAIL = " Hnnn Hnnn Hnnn Hnnn Hnnn Hnnn Hnnn Hnnn"
#: document -> (compatibility mode, alignment)
SETTINGS = {"15-both": (15, "both"), "15-left": (15, "left"), "14-both": (14, "both"), "none-both": (None, "both")}
#: The second round's documents (``fine_cases``): mode 15, justified and left-aligned.
FINE_SETTINGS = {"fine-15-both": (15, "both"), "fine-15-left": (15, "left")}


@dataclass(frozen=True)
class Case:
    face: str
    half_points: int
    spaces: int
    target: float  # f, the fraction of the spaces' width aimed at
    delta: int  # the line through W2, less the budget, in 1/4096 pt (> 0: over the edge)
    ind_right: int
    text: str

    @property
    def space_units(self) -> int:
        return budget.WIDTHS[self.face][" "] * self.half_points * self.spaces


#: The fixed word, and others tried in its place where no column composes the case.
FIRST_WORDS = (W1, "Hnie", "Hnia", "Hnit", "Hnir", "Hnid", "Hnil")


def _case(face: str, hp: int, spaces: int, delta: int, target: float) -> Case | None:
    if face not in budget._SUMS:
        budget._SUMS[face] = budget._sums(face)
    sums = budget._SUMS[face]
    for first in FIRST_WORDS:
        fixed = budget._units(face, (first + " ") * spaces, hp)
        for col in range(8800, 1200, -5):
            want = col * 1024 // 5 + delta - fixed
            if want > 0 and want % hp == 0 and want // hp < len(sums) and sums[want // hp]:
                word = sums[want // hp]
                if len(word) < 3:
                    continue
                text = (first + " ") * spaces + word + TAIL
                return Case(face, hp, spaces, target, delta, budget.COLUMN - col, text)
    return None


@functools.lru_cache(maxsize=None)
def cases() -> tuple[Case, ...]:
    out = []
    for face in FACES:
        for hp in SIZES:
            for spaces in SPACES:
                space_units = budget.WIDTHS[face][" "] * hp * spaces
                seen = set()
                deltas = [(round(f * space_units), f) for f in FRACTIONS]
                for delta, f in deltas:
                    if hp % 2 == 0 and delta % 2:
                        delta += 1
                    if delta in seen:
                        continue
                    seen.add(delta)
                    case = None
                    # The nearest composable δ on the same side (of 0 and of the target).
                    for step in range(0, 60):
                        for candidate in ((delta + step, delta - step) if delta else (delta + step,)):
                            if hp % 2 == 0 and candidate % 2:
                                continue
                            case = _case(face, hp, spaces, candidate,
                                         f if f is not None else candidate / space_units)
                            if case is not None:
                                break
                        if case is not None:
                            break
                    if case is None:
                        raise RuntimeError(f"no case for {face} {hp} {spaces} {delta}")
                    out.append(case)
    return tuple(out)


@dataclass(frozen=True)
class FineCase:
    """A second-round case: the line's words as ``(text, half points)`` runs, ``W2`` last."""

    family: str
    face: str
    runs: tuple  # ((text, half points), ...), the line under test then the tail
    spaces: tuple  # the half points of every space before W2
    delta: int
    ind_right: int

    @property
    def space_units(self) -> int:
        return sum(budget.WIDTHS[self.face][" "] * hp for hp in self.spaces)

    @property
    def text(self) -> str:
        return "".join(text for text, _ in self.runs)


def _fine(family: str, face: str, lead: list, hp: int, delta: int) -> FineCase | None:
    """``lead``: the ``(text, half points)`` runs before ``W2``, whose size is ``hp``."""
    if face not in budget._SUMS:
        budget._SUMS[face] = budget._sums(face)
    sums = budget._SUMS[face]
    fixed = sum(budget._units(face, text, size) for text, size in lead)
    spaces = tuple(size for text, size in lead for c in text if c == " ")
    for col in range(8800, 1200, -5):
        want = col * 1024 // 5 + delta - fixed
        if want > 0 and want % hp == 0 and want // hp < len(sums) and sums[want // hp] and len(sums[want // hp]) >= 3:
            runs = tuple(lead) + ((sums[want // hp] + TAIL, hp),)
            return FineCase(family, face, runs, spaces, delta, budget.COLUMN - col)
    return None


def _fine_group(family: str, face: str, lead: list, hp: int, deltas) -> list:
    out, seen = [], set()
    for delta in deltas:
        for step in range(0, 60):
            found = None
            for candidate in (delta + step, delta - step):
                if candidate in seen:
                    continue
                found = _fine(family, face, lead, hp, candidate)
                if found is not None:
                    break
            if found is not None:
                seen.add(found.delta)
                out.append(found)
                break
    return out


@functools.lru_cache(maxsize=None)
def fine_cases() -> tuple:
    """The second round, about the edge the first found -- a quarter of the spaces' width:

    * ``plain`` -- the first round's lines, δ at single units about a quarter of the spaces'
      width (even units at 11 pt, where only even widths exist);
    * ``mixed`` -- half the spaces at 11 pt and half at 14.5 pt, ``W2`` at 14.5;
    * ``double`` -- two spaces between words.
    """
    out = []
    for face in FACES:
        for hp in SIZES:
            for spaces in SPACES:
                quarter = budget.WIDTHS[face][" "] * hp * spaces / 4
                lead = [((W1 + " ") * spaces, hp)]
                out += _fine_group("plain", face, lead, hp, [int(quarter) + d for d in range(-3, 4)])
    for face in ("Calibri", "Times New Roman"):
        space = budget.WIDTHS[face][" "]
        for pairs in (1, 2, 3):
            lead = [((W1 + " ") * pairs, 22), ((W1 + " ") * pairs, 29)]
            quarter = space * (22 + 29) * pairs / 4
            out += _fine_group("mixed", face, lead, 29, [int(quarter) + d for d in range(-3, 4)])
        for spaces in (2, 4):
            lead = [((W1 + "  ") * spaces, 22)]
            single = space * 22 * spaces
            out += _fine_group("double", face, lead, 22,
                               [round(f * single) for f in (0.2, 0.25, 0.3, 0.4, 0.45, 0.5, 0.55)])
    return tuple(out)


@functools.lru_cache(maxsize=None)
def third_cases() -> tuple:
    """The third round: is the edge also a fraction of ``W2`` itself?  The render probe's
    first justified line held ``is`` where ``the`` needed 14% of its spaces -- well inside
    a quarter -- but 50.3% of ``the``.  So ``W2`` is short (1,500, 2,200 and 3,000 font
    units wide) after 24 two-letter words, whose spaces' quarter is more than the overrun,
    and δ is a fraction of ``W2``'s own width: 0.30 to 0.70.  The line's text ends at ``W2``'s end when δ = 0 is
    subtracted, so δ says how much of ``W2`` sticks out past the edge."""
    out = []
    for face, hp in (("Calibri", 22), ("Calibri", 29), ("Times New Roman", 22), ("Arial", 22)):
        if face not in budget._SUMS:
            budget._SUMS[face] = budget._sums(face)
        sums = budget._SUMS[face]
        lead_text = "Hn " * 24
        fixed = budget._units(face, lead_text, hp)
        for target in (1500, 2200, 3000):  # W2's width, font units
            for f in (0.30, 0.40, 0.45, 0.48, 0.50, 0.52, 0.55, 0.60, 0.70):
                found = None
                for width in sorted(range(target - 400, target + 400), key=lambda v: abs(v - target)):
                    if width >= len(sums) or not sums[width] or len(sums[width]) < 2:
                        continue
                    w2 = width * hp
                    for d in range(0, 40):
                        for delta in (round(f * w2) + d, round(f * w2) - d):
                            if (fixed + w2 - delta) % 1024 == 0:
                                col = (fixed + w2 - delta) // 1024 * 5
                                if 1200 < col <= 8800:
                                    found = FineCase("short", face, ((lead_text, hp), (sums[width] + TAIL, hp)),
                                                     (hp,) * 24, delta, budget.COLUMN - col)
                                    break
                        if found:
                            break
                    if found:
                        break
                if found:
                    out.append(found)
    return tuple(out)


THIRD_SETTINGS = {"third-15-both": (15, "both")}


def build_third(name: str) -> bytes:
    mode, jc = THIRD_SETTINGS[name]
    extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    body = "".join(_fine_paragraph(case, jc) for case in third_cases())
    return w.package(body, final_section=budget.section(), extra_parts=extra)


def _fine_paragraph(case: FineCase, jc: str) -> str:
    fonts = {"ascii": case.face, "hAnsi": case.face, "eastAsia": case.face, "cs": case.face}
    runs = "".join(w.run(text, case.face, hp) for text, hp in case.runs)
    last = case.runs[-1][1]
    mark = {"rFonts": fonts, "sz": last, "szCs": last}
    return wml.paragraph(runs, mark=mark, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"},
                         ind={"left": 0, "right": case.ind_right}, jc=jc)


def build_fine(name: str) -> bytes:
    mode, jc = FINE_SETTINGS[name]
    extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    body = "".join(_fine_paragraph(case, jc) for case in fine_cases())
    return w.package(body, final_section=budget.section(), extra_parts=extra)


def _paragraph(case: Case, jc: str) -> str:
    run = w.run(case.text, case.face, case.half_points)
    fonts = {"ascii": case.face, "hAnsi": case.face, "eastAsia": case.face, "cs": case.face}
    mark = {"rFonts": fonts, "sz": case.half_points, "szCs": case.half_points}
    return wml.paragraph(run, mark=mark, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"},
                         ind={"left": 0, "right": case.ind_right}, jc=jc)


def build(name: str) -> bytes:
    mode, jc = SETTINGS[name]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    body = "".join(_paragraph(case, jc) for case in cases())
    return w.package(body, final_section=budget.section(), extra_parts=extra)


if __name__ == "__main__":
    print(len(cases()), "cases a document")
    for name in SETTINGS:
        print(name, len(build(name)), "bytes")
    print(len(fine_cases()), "second-round cases")
    for name in FINE_SETTINGS:
        print(name, len(build_fine(name)), "bytes")
    print(len(third_cases()), "third-round cases")
