#!/usr/bin/env python3
"""Is the horizontal axis quantised?  Reads ``make_x_grid_probe.py``'s export and says.

Every position is a pen position from the PDF content stream (``quartz_pdf``), in device
px, exact to the 1e-4 px Quartz prints.  Three candidate grids are scored for each
authored indent and tab stop -- the device pixel, the twip, and Word's 1/4096 pt layout
unit -- and for glyph advances the per-glyph increments are compared with the unrounded
advance and with a whole pixel.

Measured on Word 16.106 (ROADMAP.md, Phase 2): x is **not** snapped to device pixels.
Indents and tab stops land on the authored twip value rounded to 1/4096 pt (residual
<= 6e-5 px, the print precision); glyph advances are the unrounded
``advance/upm x size`` (exact in 1/4096 pt for 2048-upm faces at half-point sizes); a
run boundary rounds nothing; only the *drawn* font size snaps (46 px for 11 pt).

Usage::

    python tools/read_x_grid_probe.py
"""

from __future__ import annotations

import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_x_grid_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg.vertical import twips_to_px  # noqa: E402

MARGIN_PX = 300  # 1440 twips


def main() -> int:
    pdf = oracle.export(probe.build(), name="x-grid-probe")
    lines = quartz_pdf.lines(quartz_pdf.read(pdf))
    items = probe.items()
    if len(lines) != len(items):
        raise SystemExit(f"{len(lines)} lines for {len(items)} probe paragraphs")

    # The authored values are whole twips, so "unrounded" is also "on the twip grid".
    worst = {"device px": 0.0, "unrounded": 0.0, "1/4096 pt": 0.0}
    for item, line in zip(items, lines):
        if item.kind in ("indent", "firstline"):
            twips, x = item.value, line.runs[0].x - MARGIN_PX
        elif item.kind == "tab":
            twips, x = item.value, line.runs[2].x - MARGIN_PX
        else:
            continue
        exact = twips / 4.8
        worst["unrounded"] = max(worst["unrounded"], abs(x - exact))
        worst["device px"] = max(worst["device px"], abs(x - round(exact)))
        worst["1/4096 pt"] = max(worst["1/4096 pt"], abs(x - float(twips_to_px(twips))))
    print("indents and tab stops at 0..24 twips and primes, worst residual per grid:")
    for grid, residual in worst.items():
        print(f"  {grid:10} {residual:.5f} px")

    print("glyph advances (one run vs one run per glyph):")
    for item, line in zip(items, lines):
        if item.kind != "adv-one":
            continue
        total_one = line.runs[-1].x - MARGIN_PX
        each = lines[items.index(item) + 1]
        xs = [run.x for run in each.runs]
        steps = [b - a for a, b in zip(xs, xs[1:-1])]
        spread = max(steps) - min(steps)
        whole = all(abs(s - round(s)) < 1e-3 for s in steps)
        print(
            f"  {item.face:16} {item.half_points / 2:5} pt {item.glyph}: step {steps[0]:.4f} px"
            f" (spread {spread:.4f}, whole pixels: {whole}); totals {total_one:.4f} vs"
            f" {xs[-1] - MARGIN_PX:.4f}"
        )
    drawn = {(it.face, it.half_points): ln.runs[0].size_px for it, ln in zip(items, lines)}
    print("drawn size (Tm scale) vs authored size in px:")
    for (face, hp), size in sorted(drawn.items()):
        print(f"  {face:16} {hp / 2:5} pt: authored {hp / 2 * 300 / 72:.3f} px, drawn {size:g} px")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
