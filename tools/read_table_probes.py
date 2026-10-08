#!/usr/bin/env python3
"""Score the table probes: every glyph of every cell against where Word drew it.

The table probes (``make_table_*_probe.py``, ROADMAP.md "Tables -- measured") are laid
out and drawn by the library from the file alone -- its grid, cell boxes, line breaker,
row heights and pagination -- and compared glyph for glyph with Word's PDF by the
glyph-position check (``glyphs.compare``): a cell's text start and every advance on
Word's pen position, every baseline on Word's device-pixel row, in Word's face and
size.  A cell's first glyph is a text object's start, so its x is held exactly, and a
right-aligned word ends where the cell's line may end, so the line budget is held too.
Word's filled rectangles are compared as well (``read_render.rect_score``).

With ``--record``, everything Word drew that they compare against, and every face number
the renderer asked for, is written to ``tests/fixtures/table-observations.json``, from
which ``tests/test_tables.py`` holds the model to all of it offline.

Usage::

    python tools/read_table_probes.py            # export (cached), compare, summarise
    python tools/read_table_probes.py -v         # ... and list every disagreement
    python tools/read_table_probes.py --record   # ... and write the recording
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "table-observations.json"


def probes() -> list:
    """The table probe generators, in stage order."""
    import make_table_border_probe
    import make_table_content_probe
    import make_table_geometry_probe
    import make_table_merge_probe
    import make_table_pages_probe
    import make_table_row_probe
    import make_table_style_size_probe

    return [make_table_geometry_probe, make_table_content_probe, make_table_style_size_probe, make_table_row_probe,
            make_table_merge_probe, make_table_border_probe, make_table_pages_probe]


def documents() -> list[tuple[str, bytes]]:
    """``(name, .docx bytes)`` of every table probe document."""
    out = []
    for module in probes():
        stem = module.__name__.replace("make_", "").replace("_probe", "").replace("_", "-")
        for setting in module.SETTINGS:
            out.append((f"{stem}-{setting}", module.build(setting)))
    return out


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
        objects, fills = read_render.word_objects(pdf), read_render.word_fills(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, rects, warnings = read_render.score(name, data, fonts, objects, fills)
        print(result.row())
        print(f"{'':34} rectangles: model {rects['all'][0]} px, Word {rects['all'][1]} px, model only "
              f"{rects['all'][2]}, Word only {rects['all'][3]}; warnings {warnings}")
        if args.verbose:
            for problem in result.problems[:40]:
                print("   ", problem.replace("|", ""))
        recorded[name] = {"objects": objects, "fills": fills, "glyphs": read_render.row(result), "rects": rects,
                          "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each table probe document (tools/make_table_*_probe.py) -- "
                       "every text object ([page, baseline px, /BaseFont, drawn size px, text, pen x in 1e-4 "
                       "px, each next glyph's step in 1e-3 px]) and every filled rectangle ([page, x0, y0, x1, "
                       "y1, colour], device px) -- the model's scores against them, and every face number the "
                       "renderer asked for. Measurements only; regenerate with tools/read_table_probes.py "
                       "--record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
