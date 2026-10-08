#!/usr/bin/env python3
"""Read the autofit probe (``make_table_autofit_probe.py``): how wide Word made each
case's first column, against its content.

For each case: the first cell's text start (x0), the widest line Word drew in it (its
pen end less x0: the content), the second cell's text start (x1), and ``x1 - x0 -
content`` -- what Word put around the content: with the column sized to its content,
the two cells' margins and the border between them.  Autofit is not modelled
(ROADMAP.md, "Tables -- measured", stage 7); this records what Word does.

Usage::

    python tools/read_table_autofit_probe.py
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_table_autofit_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402


def columns(pdf) -> list[tuple[str, float, float, float]]:
    """Per case: ``(name, x0, content, x1)`` in device px."""
    runs = [run for page in quartz_pdf.read(pdf) for run in sorted(page.runs, key=lambda r: (r.y, r.x))
            if run.text.strip()]
    starts = [k for k, run in enumerate(runs) if run.text.startswith("Case ")]
    out = []
    for number, (name, _, _) in enumerate(probe.CASES):
        region = runs[starts[number] + 1:starts[number + 1] if number + 1 < len(starts) else len(runs)]
        second = next(run for run in region if run.text.startswith(f"{number}Kx"))
        x0 = next(run for run in region if run.text.startswith(f"{number}K") and run is not second).x
        lines: dict[float, float] = {}
        for run in region:
            if run.x >= x0 - 0.5 and run.x < second.x - 0.5:
                lines[run.y] = max(lines.get(run.y, 0.0), run.end)
        content = max(end - x0 for end in lines.values())
        out.append((name, x0, content, second.x))
    return out


def main() -> int:
    for setting in probe.SETTINGS:
        pdf = oracle.export(probe.build(setting), name=f"table-autofit-{setting}")
        print(f"== {setting}")
        for name, x0, content, x1 in columns(pdf):
            print(f"  {name:16} x0 {x0:7.1f}  content {content:7.1f}  x1 {x1:7.1f}  around {x1 - x0 - content:6.1f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
