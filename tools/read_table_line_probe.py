#!/usr/bin/env python3
"""Measure ``make_table_line_probe.py``: every border line and every cell's shading, to the
pixel, Word's against the model's.

For each table cell the model draws, three things are read off Word's filled rectangles
at the place the model drew them (``observe``):

* ``v`` -- each vertical line down the row (the cell's left one, and the row's last right
  one): the run of border-coloured pixels across the row's middle nearest the model's;
* ``s`` -- the cell's shading: the widest fill of its colour through the cell's middle,
  ``[x0, x1, y0, y1]``;
* ``h`` -- each horizontal line across the cell (the band above the row, and below the
  last row): the run of border-coloured pixels down the cell's middle nearest the
  model's.

The same function read over the model's own rectangles gives the model's numbers, so the
two are compared key by key.  ``--record`` writes Word's numbers to
``tests/fixtures/table-line-observations.json``; ``tests/test_table_lines.py`` holds the
model to them offline.

Usage::

    python tools/read_table_line_probe.py [--record] [-v]
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_table_line_probe as probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "table-line-observations.json"
BORDER = probe.BORDER
SHADING = probe.SHADING


def model_fills(layout) -> list:
    """The model's rectangles in :func:`read_render.word_fills`' form."""
    return [[index, float(rule.x), float(rule.y), float(rule.x + rule.width), float(rule.y + rule.height),
             (rule.color or "000000").upper()]
            for index, page in enumerate(layout.pages) for rule in page.rules]


def _run(fills: list, page: int, colour: str, along: str, at: float, centre: float, reach: float = 15):
    """The contiguous run of ``colour`` pixels across ``at`` (a y for ``along="x"``, an x
    for ``along="y"``) nearest ``centre``, within ``reach`` of it; ``None`` if there is none."""
    cells: set[int] = set()
    for f in fills:
        if f[0] != page or f[5] != colour:
            continue
        x0, y0, x1, y1 = round(f[1]), round(f[2]), round(f[3]), round(f[4])
        if along == "x" and y0 <= at < y1:
            cells.update(range(max(x0, int(centre - reach)), min(x1, int(centre + reach) + 1)))
        elif along == "y" and x0 <= at < x1:
            cells.update(range(max(y0, int(centre - reach)), min(y1, int(centre + reach) + 1)))
    runs: list[list[int]] = []
    for c in sorted(cells):
        if runs and c == runs[-1][1]:
            runs[-1][1] = c + 1
        else:
            runs.append([c, c + 1])
    if not runs:
        return None
    return min(runs, key=lambda r: (abs((r[0] + r[1]) / 2 - centre), r[0]))


def observe(layout, fills: list) -> dict:
    """``{key: [...]}`` read off ``fills`` at every place the model draws a table rule (the
    module docstring)."""
    out: dict = {}
    for index, page in enumerate(layout.pages):
        by_path: dict[str, list] = {}
        for rule in page.rules:
            if rule.kind in ("table-border", "cell-shading"):
                by_path.setdefault(rule.path, []).append(rule)
        for path, rules in by_path.items():
            shading = [r for r in rules if r.kind == "cell-shading" and r.height >= 20]
            vertical = sorted((r for r in rules if r.kind == "table-border" and r.height >= 20 and r.width < 40),
                              key=lambda r: r.x)
            horizontal = sorted((r for r in rules if r.kind == "table-border" and r.width >= 40), key=lambda r: r.y)
            for n, r in enumerate(vertical):
                found = _run(fills, index, BORDER, "x", float(r.y + r.height / 2), float(r.x + r.width / 2))
                out[f"{index}|{path}|v{n}"] = found
            for n, r in enumerate(horizontal):
                found = _run(fills, index, BORDER, "y", float(r.x + r.width / 2), float(r.y + r.height / 2))
                out[f"{index}|{path}|h{n}"] = found
            for r in shading:
                mx, my = float(r.x + r.width / 2), float(r.y + r.height / 2)
                found = [f for f in fills if f[0] == index and f[5] == SHADING and f[1] <= mx < f[3] and f[2] <= my < f[4]]
                best = max(found, key=lambda f: (f[3] - f[1], f[4] - f[2]), default=None)
                out[f"{index}|{path}|s"] = None if best is None else [round(v) for v in best[1:5]]
    return out


def keys_digest(keys: list) -> str:
    """What a recording is aligned by: the SHA-256 of the model's keys, in order."""
    import hashlib

    return hashlib.sha256("\n".join(keys).encode()).hexdigest()[:16]


def compare(word: dict, ours: dict) -> Counter:
    """Per kind (``v``, ``h``, ``s``): how many keys agree and how many are read at all."""
    out: Counter = Counter()
    for key, value in ours.items():
        kind = key.rsplit("|", 1)[1][0]
        out[f"{kind} scored"] += 1
        out[f"{kind} agree"] += word.get(key) == value
    return out


def main(argv: list[str]) -> int:
    import oracle
    from docx2svg import _render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for setting in probe.SETTINGS:
        name = f"table-line-{setting}"
        data = probe.build(setting)
        fills = read_render.word_fills(oracle.export(data, name=name))
        fonts = render_record.RecordingFonts(data, faces)
        _, layout, _, _ = _render(data, render_record.options(fonts))
        word = observe(layout, fills)
        ours = observe(layout, model_fills(layout))
        score = compare(word, ours)
        print(f"{name:24} " + "  ".join(f"{k} {score[k + ' agree']} / {score[k + ' scored']}"
                                          for k in ("v", "h", "s")))
        if args.verbose:
            for key, value in ours.items():
                if word.get(key) != value:
                    print(f"    {key}: Word {word.get(key)}, model {value}")
        keys = sorted(ours)
        recorded[name] = {"keys": keys_digest(keys), "word": [word.get(key) for key in keys]}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("Word 16.106's table lines and shading for tools/make_table_line_probe.py, read off its "
                       "filled rectangles where the model draws each (tools/read_table_line_probe.py.observe): "
                       "per document, in the sorted order of the model's keys ('page|cell path|v<n>' a vertical "
                       "line's [x0, x1) across the row's middle, 'h<n>' a horizontal line's [y0, y1) down the "
                       "cell's middle, 's' the cell's shading [x0, x1, y0, y1]; device px), with the keys' "
                       "digest. Measurements only; regenerate with --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
