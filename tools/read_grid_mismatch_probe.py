#!/usr/bin/env python3
"""Score the grid-mismatch probe (``make_grid_mismatch_probe.py``): rows with more cells
than ``w:tblGrid`` has, and fewer, every glyph against where Word drew it
(``read_render.score``).

With ``--record``, what Word drew is written to
``tests/fixtures/grid-mismatch-observations.json``, from which ``tests/test_grid_mismatch.py``
holds the model to it offline.

Usage::

    python tools/read_grid_mismatch_probe.py            # export (cached), compare, summarise
    python tools/read_grid_mismatch_probe.py -v         # ... and list every disagreement
    python tools/read_grid_mismatch_probe.py --record   # ... and write the recording
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

OBSERVATIONS = read_render.FIXTURES / "grid-mismatch-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_grid_mismatch_probe

    return [(f"grid-mismatch-{setting}", make_grid_mismatch_probe.build(setting))
            for setting in make_grid_mismatch_probe.SETTINGS]


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
            for problem in result.problems[:60]:
                print("   ", problem.replace("|", ""))
        recorded[name] = {"objects": objects, "fills": fills, "glyphs": read_render.row(result), "rects": rects,
                          "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for the grid-mismatch probe documents "
                       "(tools/make_grid_mismatch_probe.py) -- every text object and filled rectangle, as "
                       "read_table_probes.py records them -- the model's scores against them, and every face "
                       "number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_grid_mismatch_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
