#!/usr/bin/env python3
"""Score ``make_xref_field_probe.py``: every glyph against Word's, where ``PAGEREF``, ``REF``
and ``SEQ`` fields are computed.

Records, with ``--record``, Word's text objects and every face number the renderer asked
for in ``tests/fixtures/xref-field-observations.json``; ``tests/test_xref_fields.py`` holds the
renderer to them offline.

Usage::

    python tools/read_xref_field_probe.py [-v] [--record]
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

OBSERVATIONS = read_render.FIXTURES / "xref-field-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_xref_field_probe

    return [(f"xref-{setting}", make_xref_field_probe.build(setting)) for setting in make_xref_field_probe.SETTINGS]


def refuted() -> list[tuple[str, tuple]]:
    """``(what, mock.patch.object arguments)`` of each rule the probe refutes; the recorder
    lays the probe out under each too, so the faces they ask for are recorded."""
    import dataclasses

    from docx2svg import fields
    from docx2svg.parse import document as parse_document

    reference_runs, field_text = fields.reference_runs, fields.field_text

    def instruction_format(paragraphs, run, marks):
        return tuple(dataclasses.replace(piece, properties=run.properties)
                     for piece in reference_runs(paragraphs, run, marks))

    def referencing_format(instruction, keyword, **numbers):
        if keyword == "PAGEREF":
            numbers["section_format"] = None  # the decimal section the field is in
        return field_text(instruction, keyword, **numbers)

    return [
        ("before: PAGEREF, REF and SEQ drawn as cached",
         (parse_document, "COMPUTED_FIELDS", frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES"}))),
        ("REF in its instruction's format, as PAGE is", (fields, "reference_runs", instruction_format)),
        ("PAGEREF in the format of the section the field is in", (fields, "field_text", referencing_format)),
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
        recorded[name] = {"objects": objects, "glyphs": read_story_probe.row(result), "warnings": warnings}
        from unittest import mock

        for what, patch in refuted():
            with mock.patch.object(*patch):
                other, _rects, _warnings = read_render.score(name, data, fonts, objects, [])
            print(f"   refuted, {what}: {other.matched} matched, x {other.x_agree}, baseline {other.y_agree}, "
                  f"face {other.face_agree}")
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_xref_field_probe.py -- every text "
                       "object, as render-observations.json records them -- the renderer's glyph scores, and every "
                       "face number it asked for. Measurements only; regenerate with tools/read_xref_field_probe.py "
                       "--record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
