#!/usr/bin/env python3
"""Measure ``make_para_border_probe.py``: where Word draws each paragraph's bottom border,
against the model and the rules tried for it.

Each bordered paragraph's red fill is read off Word's filled rectangles near where the
model draws its bottom border (its row ``y0`` and its height); the model's own row, and
the rows the rules below would give from the model's exact line geometry, are scored
against it.  Not adopted: none of them is exact (ROADMAP.md, Phase 5.15).

``--record`` writes Word's rows to ``tests/fixtures/para-border-observations.json``;
``tests/test_para_border.py`` holds the model's count to it.

Usage::

    python tools/read_para_border_probe.py [--record]
"""

from __future__ import annotations

import argparse
import math
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_para_border_probe as probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "para-border-observations.json"


def model_rows(layout) -> list:
    """``[page, y0, height]`` of every bottom border the model draws, in document order."""
    return [[index, int(rule.y), int(rule.height)] for index, page in enumerate(layout.pages)
            for rule in page.rules if rule.kind == "border-bottom"]


def word_rows(ours: list, fills: list) -> list:
    """Word's ``[y0, height]`` of the red fill nearest each of the model's borders."""
    out = []
    for page, y0, _ in ours:
        near = [f for f in fills if f[0] == page and f[5] == probe.RED and abs(f[2] - y0) < 8]
        best = min(near, key=lambda f: abs(f[2] - y0), default=None)
        out.append(None if best is None else [round(best[2]), round(best[4] - best[2])])
    return out


def geometry(layout) -> list:
    """Per bordered paragraph, the exact geometry the rules tried use: its last line's
    top and pitch (``Line.top``, ``Line.pitch``), in document order."""
    out = []
    for page in layout.pages:
        for line in page.lines:
            if "Bordered" in "".join("".join(span.chars) for span in line.spans):
                out.append((line.top, line.pitch))
    return out


def rules(case: dict, top: Fraction, pitch: Fraction) -> dict:
    """The rows two of the rules tried give (ROADMAP.md, 5.15's table)."""
    from docx2svg.vertical import round_half_up, twips_to_px

    space = twips_to_px(case["space"] * 20)
    return {"round(top + pitch + space)": round_half_up(top + pitch + space),
            "ceil(top + pitch) + floor(space)": math.ceil(top + pitch) + math.floor(space)}


def main(argv: list[str]) -> int:
    import oracle
    from docx2svg import _render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for setting in probe.SETTINGS:
        name = f"para-border-{setting}"
        data = probe.build(setting)
        fills = read_render.word_fills(oracle.export(data, name=name))
        fonts = render_record.RecordingFonts(data, faces)
        _, layout, _, _ = _render(data, render_record.options(fonts))
        ours = model_rows(layout)
        word = word_rows(ours, fills)
        scores = {"model": sum(w is not None and w[0] == o[1] for o, w in zip(ours, word))}
        for case, (top, pitch), w in zip(probe.CASES, geometry(layout), word):
            for rule, row in rules(case, top, pitch).items():
                scores[rule] = scores.get(rule, 0) + (w is not None and w[0] == row)
        print(f"{name:18} bottom borders where Word drew them, of {len(ours)}: "
              + ", ".join(f"{k} {v}" for k, v in scores.items()))
        recorded[name] = word
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("Where Word 16.106 drew the bottom border of each bordered paragraph of "
                       "tools/make_para_border_probe.py: [y0, height], device px, in document order "
                       "(tools/read_para_border_probe.py). Measurements only; regenerate with --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
