#!/usr/bin/env python3
"""Score ``make_wrap_tables_probe.py``: every glyph against Word's, where a square-wrapped text
box stands beside several tables and lines run on to the next page.

Records, with ``--record``, Word's text objects and every face number the renderer asked
for in ``tests/fixtures/wrap-tables-observations.json``; ``tests/test_wrap_tables.py`` holds the
renderer to them offline.

Usage::

    python tools/read_wrap_tables_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "wrap-tables-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_wrap_tables_probe

    return [(f"wraptables-{setting}", make_wrap_tables_probe.build(setting)) for setting in make_wrap_tables_probe.SETTINGS]


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for name, data in documents():
        pdf = oracle.export(data, name=name)
        objects = read_render.word_objects(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, _rects, warnings = read_render.score(name, data, fonts, objects, [])
        print(result.row())
        if args.verbose:
            for problem in result.problems[:40]:
                print("   ", problem.replace("|", ""))
        recorded[name] = {"objects": objects, "glyphs": read_story_probe.row(result), "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_wrap_tables_probe.py -- every text "
                       "object, as render-observations.json records them -- the renderer's glyph scores, and every "
                       "face number it asked for. Measurements only; regenerate with tools/read_wrap_tables_probe.py "
                       "--record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
