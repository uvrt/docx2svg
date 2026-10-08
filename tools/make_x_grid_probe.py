#!/usr/bin/env python3
"""Is the horizontal axis quantised?  A probe built to say *what* snaps, not just *whether*.

Every vertical position Word 16.106 exports lands on the 1/300-inch device grid.  The
Phase 0 indent ladder came back at 72.0000, 90.0000 ... pt -- which are *also* whole
device pixels, so it could not distinguish "Word snaps x" from "those indents happened to
be on the grid".  This probe can, and it separates three candidate granularities:

* **Line start** -- ``w:ind/@w:left`` and ``@w:firstLine`` at every twip from 0 to 24 (a
  device pixel is 4.8 twips, so five whole pixels in twip steps) plus a few primes.  If x
  snaps to the device grid the start moves in 4.8-twip jumps; if it snaps to twips it
  moves in 1/4.8 px steps; if it does not snap it moves by exactly ``t / 4.8`` px.
* **Tab stops** -- ``w:tab/@w:pos`` at 1440 + 0..12 twips, and an off-grid prime.
* **Glyph advances within a line** -- ten identical glyphs at sizes whose advance is a
  fractional number of pixels, twice: once as one run (only the paragraph mark's
  position is exact), and once with every glyph its own run in alternating colours.  A
  colour change forces Quartz to open a new text object, and a new text object carries
  its pen position to 1e-4 px -- so every glyph's x is read exactly.  If each glyph
  snapped, the increments would be whole pixels; if only the line start snapped, the
  increments would be the unrounded advance.  Comparing the one-run total with the
  per-run total also says whether a *run boundary* rounds.

Three faces (Calibri, Arial, Times New Roman) because a rule that happens to hold for
one face's advance widths should not be mistaken for a rule.

Throwaway: the ``.docx`` goes straight to the oracle and is not committed.  Regenerate
and measure with ``python tools/read_x_grid_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx as w

FACES = ("Calibri", "Arial", "Times New Roman")
INDENTS = tuple(range(0, 25)) + (37, 53, 97, 113, 131, 1447)
FIRST_LINES = (7, 13, 37)
TAB_OFFSETS = tuple(range(0, 13)) + (37,)
ADVANCE_HALF_POINTS = (13, 21, 23)  # 6.5, 10.5, 11.5 pt: fractional px advances
ADVANCE_GLYPHS = ("H", "i")
REPEAT = 10


@dataclass(frozen=True)
class Item:
    kind: str  # "indent" | "firstline" | "tab" | "adv-one" | "adv-each" | "center" | "right"
    face: str
    half_points: int
    value: int = 0  # twips for indent / tab, unused otherwise
    glyph: str = "H"


def items() -> list[Item]:
    out: list[Item] = []
    out += [Item("indent", "Calibri", 22, t) for t in INDENTS]
    out += [Item("firstline", "Calibri", 22, t) for t in FIRST_LINES]
    out += [Item("tab", "Calibri", 22, 1440 + k) for k in TAB_OFFSETS]
    for face in FACES:
        for hp in ADVANCE_HALF_POINTS:
            for glyph in ADVANCE_GLYPHS:
                out.append(Item("adv-one", face, hp, glyph=glyph))
                out.append(Item("adv-each", face, hp, glyph=glyph))
    for face in FACES:
        out.append(Item("center", face, 23))
        out.append(Item("right", face, 23))
    return out


def _paragraph(item: Item) -> str:
    face, hp = item.face, item.half_points
    mark = w.rpr(face, hp)
    if item.kind == "indent":
        return w.paragraph(w.run("H", face, hp), indent_left=item.value, mark_rpr=mark)
    if item.kind == "firstline":
        return w.paragraph(w.run("H", face, hp), first_line=item.value, mark_rpr=mark)
    if item.kind == "tab":
        runs = w.run("H", face, hp) + w.run("\t", face, hp) + w.run("H", face, hp)
        return w.paragraph(runs, tabs=(item.value,), mark_rpr=mark)
    if item.kind == "adv-one":
        return w.paragraph(w.run(item.glyph * REPEAT, face, hp), mark_rpr=mark)
    if item.kind == "adv-each":
        runs = "".join(
            w.run(item.glyph, face, hp, color="C00000" if i % 2 else "000000")
            for i in range(REPEAT)
        )
        return w.paragraph(runs, mark_rpr=mark)
    if item.kind in ("center", "right"):
        return w.paragraph(w.run("HiHiH", face, hp), jc=item.kind, mark_rpr=mark)
    raise ValueError(item.kind)


def build() -> bytes:
    return w.package("".join(_paragraph(item) for item in items()))


if __name__ == "__main__":
    import sys
    from pathlib import Path

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("x-grid-probe.docx")
    target.write_bytes(build())
    print(f"wrote {target}")
