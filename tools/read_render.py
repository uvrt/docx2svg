#!/usr/bin/env python3
"""Score the renderer against Word: every glyph (``glyphs.py``) and every filled rectangle.

Two instruments run over the committed documents and the generated render probes
(``make_render_probe.py``), and with ``--record`` everything Word drew that they compare
against is written to ``tests/fixtures/render-observations.json`` -- Word's text objects
(their pen positions and glyphs) and its filled rectangles, and the numbers of every face
the renderer asked for (``render_record.py``) -- so ``tests/test_render.py`` holds the
renderer to all of it offline, without Word or a font file.

* **Glyphs**: see ``glyphs.py``.  Every matched glyph's pen x (a text object's start
  exactly, an advance inside one within Quartz's encoding), drawn baseline, face and size.
* **Rectangles**: per page and colour, the device pixels the model's rules cover against
  those Word's fills cover (underlines, strikes, highlights, shading, paragraph borders),
  as ``model only`` / ``Word only`` pixel counts; ``0 / 0`` is an exact match.

Usage::

    python tools/read_render.py            # live: export (cached), compare, summarise
    python tools/read_render.py --record   # ... and write the recording
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import render_record  # noqa: E402

REPO = HERE.parent
FIXTURES = REPO / "tests" / "fixtures"
OBSERVATIONS = FIXTURES / "render-observations.json"

#: The committed documents, by the name the recording keeps them under.
COMMITTED = (
    "layout-sweep.docx", "style-document.docx",
    "samplelib/sample-blank.docx", "samplelib/sample-long.docx", "samplelib/sample-resume.docx",
    "samplelib/sample-simple.docx",
    "wordto/sample-1page.docx", "wordto/sample-5pages.docx", "wordto/sample-10pages.docx",
    "wordto/sample-with-images.docx", "wordto/sample-with-table.docx",
)


def documents() -> list[tuple[str, bytes]]:
    """``(name, .docx bytes)``: the committed documents, then the generated probes."""
    import make_render_probe

    out = [(name, (FIXTURES / name).read_bytes()) for name in COMMITTED]
    out += [(f"render-{setting}", make_render_probe.build(setting)) for setting in make_render_probe.SETTINGS]
    return out


def export(name: str, data: bytes) -> Path:
    import oracle

    if name == "layout-sweep.docx":
        return oracle.ORACLE_DIR / "layout-sweep.pdf"  # Phase 0's export, reused by name
    return oracle.export(data, name=Path(name).stem)


# -- what Word drew, compactly ------------------------------------------------------------


def word_objects(pdf: Path) -> list:
    """Every text object: ``[page, baseline, /BaseFont, drawn size, text, x0, steps]`` with
    ``x0`` in 1e-4 px and each following glyph's step from the one before in 1e-3 px."""
    import quartz_pdf

    out = []
    for line in quartz_pdf.lines(quartz_pdf.read(pdf)):
        for run in line.runs:
            steps = [round((b - a) * 1000) for a, b in zip(run.xs, run.xs[1:])]
            out.append([run.page, round(run.y, 3), run.font, run.size_px, run.text, round(run.x * 10000), steps])
    return out


def glyphs_of(objects: list) -> list[glyphs.Glyph]:
    """:func:`glyphs.word_glyphs` from a recording."""
    import baselines

    out = []
    for number, (page, y, font, size, text, x0, steps) in enumerate(objects):
        family, bold, italic = baselines._drawn_face(font)
        x = x0 / 10000
        xs = [x]
        for step in steps:
            x += step / 1000
            xs.append(x)
        for k, (char, gx) in enumerate(zip(text, xs)):
            if not char.isspace():
                out.append(glyphs.Glyph(page, char, gx, y, family, size, bold, italic, y, k == 0, number + 1, k))
    return out


def word_fills(pdf: Path) -> list:
    import quartz_pdf

    return [[f.page, round(f.x0, 2), round(f.y0, 2), round(f.x1, 2), round(f.y1, 2), f.color]
            for f in quartz_pdf.fills(pdf)]


# -- comparing ------------------------------------------------------------------------------


def rect_score(layout, fills: list) -> dict:
    """Per page and colour, ``[model px, Word px, model only, Word only]`` over the pages
    the model drew rules on; and the totals under ``"all"``."""
    import numpy as np

    by_page = defaultdict(list)
    for fill in fills:
        by_page[fill[0]].append(fill)
    out: dict = {}
    total = [0, 0, 0, 0]
    for index, page in enumerate(layout.pages):
        colors = sorted({(rule.color or "000000").upper() for rule in page.rules})
        for color in colors:
            ours = np.zeros((page.height_px, page.width_px), bool)
            theirs = np.zeros_like(ours)
            for rule in page.rules:
                if (rule.color or "000000").upper() == color:
                    ours[int(rule.y):int(rule.y + rule.height), int(rule.x):int(rule.x + rule.width)] = True
            for fill in by_page[index]:
                if fill[5] == color:
                    theirs[round(fill[2]):round(fill[4]), round(fill[1]):round(fill[3])] = True
            row = [int(ours.sum()), int(theirs.sum()), int((ours & ~theirs).sum()), int((theirs & ~ours).sum())]
            out[f"{index + 1}/{color}"] = row
            total = [a + b for a, b in zip(total, row)]
    out["all"] = total
    return out


def score(name: str, data: bytes, fonts, objects: list, fills: list) -> tuple[glyphs.Score, dict, list[str]]:
    """Render ``data`` measuring with ``fonts``; check the SVG against the layout; compare
    with Word's objects and fills.  Returns the glyph score, the rectangle score and the
    warning codes."""
    from docx2svg import _render

    options = render_record.options(fonts)
    documents_, layout, _data, _fonts = _render(data, options)
    ours = glyphs.svg_glyphs(documents_)
    exact = glyphs.layout_glyphs(layout, options.glyph_size)
    assert [(g.page, g.char) for g in ours] == [(g.page, g.char) for g in exact], f"{name}: the SVG lost a glyph"
    for a, b in zip(ours, exact):
        assert abs(a.x - float(b.x)) <= 5e-5 + 1e-9 and a.y == float(b.y), (name, a, b)
    result = glyphs.compare(name, glyphs.exact_floats(exact), glyphs_of(objects))
    return result, rect_score(layout, fills), sorted({w.code for w in options.warnings})


def row(result: glyphs.Score) -> list[int]:
    """What a recording pins of a glyph score."""
    return [result.matched, result.model, result.word, result.exact_agree, result.exact_scored,
            result.step_agree, result.step_scored, result.y_agree, result.face_agree, result.size_agree,
            result.extra, result.not_drawn]


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for name, data in documents():
        pdf = export(name, data)
        objects, fills = word_objects(pdf), word_fills(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, rects, warnings = score(name, data, fonts, objects, fills)
        print(result.row())
        print(f"{'':34} rectangles: model {rects['all'][0]} px, Word {rects['all'][1]} px, model only "
              f"{rects['all'][2]}, Word only {rects['all'][3]}; warnings {warnings}")
        if args.verbose:
            for problem in result.problems[:30]:
                print("   ", problem.replace("|", ""))
        recorded[name] = {"objects": objects, "fills": fills, "glyphs": row(result), "rects": rects,
                          "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each committed document and render probe -- every text "
                       "object ([page, baseline px, /BaseFont, drawn size px, text, pen x in 1e-4 px, each "
                       "next glyph's step in 1e-3 px]) and every filled rectangle ([page, x0, y0, x1, y1, "
                       "colour], device px) -- the renderer's scores against them, and every face number "
                       "the renderer asked for. Measurements only; regenerate with tools/read_render.py "
                       "--record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
