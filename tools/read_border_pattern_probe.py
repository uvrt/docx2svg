#!/usr/bin/env python3
"""Measure ``make_border_pattern_probe.py``: the pattern Word cuts a dotted, dashed or
double border line into, side by side with the model's (solid) line.

For every border line the model draws, Word's fills of that line's colour inside it are
read as runs along the line; the lengths of the inner runs (away from the corners) and
their period are printed per style and width.  Recorded in ROADMAP.md, Phase 5.15; not
adopted -- where a pattern starts along a side is not settled.

Usage::

    python tools/read_border_pattern_probe.py [SETTING]
"""

from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_border_pattern_probe as probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402


def runs_along(rule, fills: list, page: int) -> list[tuple[int, int]]:
    """Word's runs of ``rule``'s colour along it, px from its start."""
    horizontal = rule.width > rule.height
    x0, y0 = int(rule.x), int(rule.y)
    length = int(rule.width if horizontal else rule.height)
    covered = [False] * length
    for f in fills:
        if f[0] != page or f[5] != (rule.color or "000000").upper():
            continue
        a, b = (round(f[1]) - x0, round(f[3]) - x0) if horizontal else (round(f[2]) - y0, round(f[4]) - y0)
        across = (round(f[2]) < y0 + rule.height and round(f[4]) > y0) if horizontal else \
            (round(f[1]) < x0 + rule.width and round(f[3]) > x0)
        if across:
            for i in range(max(a, 0), min(b, length)):
                covered[i] = True
    out, start = [], None
    for i, c in enumerate(covered + [False]):
        if c and start is None:
            start = i
        elif not c and start is not None:
            out.append((start, i))
            start = None
    return out


def main(argv: list[str]) -> int:
    import oracle
    from docx2svg import _render

    setting = argv[1] if len(argv) > 1 else "none"
    data = probe.build(setting)
    fills = read_render.word_fills(oracle.export(data, name=f"border-pattern-{setting}"))
    _, layout, _, _ = _render(data, render_record.options(render_record.RecordingFonts(data)))
    patterns: dict = defaultdict(Counter)
    for index, page in enumerate(layout.pages):
        for rule in page.rules:
            if rule.style in (None, "single") or max(rule.width, rule.height) < 200:
                continue
            runs = runs_along(rule, fills, index)[1:-1]
            thick = int(min(rule.width, rule.height))
            for (a, b), (c, _) in zip(runs, runs[1:]):
                patterns[(rule.style, thick)][(b - a, c - a)] += 1
    for (style, thick), counter in sorted(patterns.items()):
        on, period = counter.most_common(1)[0][0]
        print(f"{style:13} {thick:3} px: on {on}, period {period} ({counter.most_common(1)[0][1]} of "
              f"{sum(counter.values())} inner runs)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
