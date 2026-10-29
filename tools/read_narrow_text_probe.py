#!/usr/bin/env python3
"""Score ``make_narrow_text_probe.py``: every glyph against Word's, for paragraphs whose
indents leave their text little or no width -- in a fixed-layout table, in the body and
in autofit tables.

Records, with ``--record``, Word's text objects and every face number the renderer asked
for in ``tests/fixtures/narrow-text-observations.json``; ``tests/test_narrow_text.py``
holds the renderer to them offline, and to the scores of the rule the probe refutes
(:func:`refuted`).

Usage::

    python tools/read_narrow_text_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import contextlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "narrow-text-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_narrow_text_probe

    return [(f"narrow-{setting}", make_narrow_text_probe.build(setting))
            for setting in make_narrow_text_probe.SETTINGS]


@contextlib.contextmanager
def _refused():
    """Mode 15: an autofit table as wide as the column refused (laid out on its grid)."""
    from docx2svg import table

    original = table.column_wide
    table.column_wide = lambda *args: False
    try:
        yield
    finally:
        table.column_wide = original


def refuted() -> list[tuple[str, object]]:
    """``(what, context manager)`` for each rule the probe refutes."""
    return [("mode 15: an autofit table as wide as the column laid out on its grid", _refused)]


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
        print(f"{'':34} warnings {warnings}")
        if args.verbose:
            for problem in result.problems[:40]:
                print("   ", problem.replace("|", ""))
        for what, variant in refuted():
            with variant():
                other, _rects, _warnings = read_render.score(name, data, fonts, objects, [])
            print(f"   refuted, {what}: {other.matched} matched, x {other.x_agree}, baseline {other.y_agree}, "
                  f"extra {other.extra}, not drawn {other.not_drawn}")
        recorded[name] = {"objects": objects, "glyphs": read_story_probe.row(result), "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_narrow_text_probe.py -- every text "
                       "object, as render-observations.json records them -- the renderer's glyph scores, and "
                       "every face number it asked for. Measurements only; regenerate with "
                       "tools/read_narrow_text_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
