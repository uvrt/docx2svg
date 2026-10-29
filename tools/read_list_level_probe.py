#!/usr/bin/env python3
"""Score ``make_list_level_probe.py``: every glyph against Word's -- each list item's label
(or none), its text and where both start -- for paragraphs whose list level comes from a
direct ``w:ilvl`` alone, lies past the definition's levels, or is tied to another style.

Records, with ``--record``, Word's text objects and every face number the renderer asked
for in ``tests/fixtures/list-level-observations.json``; ``tests/test_list_level.py`` holds
the renderer to them offline, and to the scores of the rules the probe refutes
(:func:`refuted`).

Usage::

    python tools/read_list_level_probe.py [-v] [--record]
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

OBSERVATIONS = read_render.FIXTURES / "list-level-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_list_level_probe

    return [(f"list-level-{setting}", make_list_level_probe.build(setting))
            for setting in make_list_level_probe.SETTINGS]


@contextlib.contextmanager
def _variant(name: str, replacement):
    """``docx2svg.resolve.cascade`` with one of its rules (a function) replaced."""
    from docx2svg.resolve import cascade

    original = getattr(cascade, name)
    setattr(cascade, name, replacement)
    try:
        yield
    finally:
        setattr(cascade, name, original)


def refuted() -> list[tuple[str, object]]:
    """``(what, context manager)`` for each rule the probe refutes."""
    return [
        ("a direct w:ilvl alone puts the list's indent below the style's",
         lambda: _variant("_numbered_directly", lambda direct: "numPr.numId" in direct)),
        ("a level past a single-level definition, or not in a multilevel one, or tied to another style, "
         "taken as named", lambda: _variant("_list_level", lambda *args, **kwargs: None)),
    ]


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
        for what, variant in refuted():
            with variant():
                other, _rects, _warnings = read_render.score(name, data, fonts, objects, [])
            print(f"   refuted, {what}: {other.matched} matched, x {other.x_agree}, baseline {other.y_agree}, "
                  f"extra {other.extra}, not drawn {other.not_drawn}")
        recorded[name] = {"objects": objects, "glyphs": read_story_probe.row(result), "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_list_level_probe.py -- every text "
                       "object, as render-observations.json records them -- the renderer's glyph scores, and "
                       "every face number it asked for. Measurements only; regenerate with "
                       "tools/read_list_level_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
