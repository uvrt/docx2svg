#!/usr/bin/env python3
"""Measure ``make_border_probe.py``: bordered lines, and the lines after them.

For each group, predicts every baseline with ``docx2svg.vertical``: a bottom border
takes ``space + width`` (``border_px``) below the bordered line, and it is part of that
line's **box**, so the bordered line rounds with space below its text
(``baseline_in_box``).  Prints the per-line residual separately for the bordered line
itself and for the plain line after it.  Groups that spilled onto a second page are
skipped (reported).

``--no-line-box`` rounds the bordered line in its own pitch, as Phase 2 did: 108 / 132.

Usage::

    python tools/read_border_probe.py
"""

from __future__ import annotations

import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import face_metrics  # noqa: E402
import make_border_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg.vertical import (  # noqa: E402
    baseline_in_box, baseline_px, border_px, face_line_box, natural_height_px, twips_to_px,
)


def main() -> int:
    in_box = "--no-line-box" not in sys.argv[1:]
    pages = quartz_pdf.read(oracle.export(probe.build(), name="border-probe"))
    calibri = face_metrics.metrics("Calibri")
    page_index = 0
    totals = {"bordered": [0, 0], "after": [0, 0]}
    for group in probe.GROUPS:
        lines = [round(line.y) for line in quartz_pdf.lines([pages[page_index]])]
        page_index += 1
        if len(lines) < 2 * probe.PAIRS:
            # The rest of the group went to the next page; skip it and that page.
            print(f"  {group}: spills onto a second page, skipped")
            page_index += 1
            continue
        h = natural_height_px(calibri, group.half_points)
        extra = border_px(group.width, group.space)
        bordered = face_line_box(calibri, group.half_points, space_after=extra if in_box else Fraction(0))
        top = twips_to_px(1440)
        residual = {"bordered": [], "after": []}
        for index in range(probe.PAIRS):
            residual["bordered"].append(lines[2 * index] - baseline_in_box(top, bordered))
            top += h + extra
            residual["after"].append(lines[2 * index + 1] - baseline_px(top, calibri, group.half_points))
            top += h
        for key, values in residual.items():
            totals[key][0] += values.count(0)
            totals[key][1] += len(values)
        print(f"  {group}: bordered {residual['bordered']}  after {residual['after']}")
    for key, (exact, count) in totals.items():
        print(f"{key:9} lines exact: {exact} / {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
