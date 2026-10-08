#!/usr/bin/env python3
"""Score floating drawings against Word: ``make_anchor_probe.py`` (where an anchor goes,
and in which layer), ``make_drawing_probe.py`` (what a shape, a text box or a group
paints), ``make_story_anchor_probe.py`` (a header's and a footer's drawings) and
``make_wrap_anchor_probe.py`` (text above and below a ``wrapTopAndBottom`` drawing) and
``make_wrap_stack_probe.py`` (such drawings in consecutive paragraphs).

Word's PDF gives, per page:

* every **picture** it drew -- its box (``pymupdf``'s image info, device px) and whether
  it is painted before the page's first glyph (behind the text) or after its last (in
  front) -- matched in paint order with the model's picture primitives
  (:meth:`docx2svg.layout.Page.layers`), so a box agrees only where the order does too;
* every **filled rectangle** (``quartz_pdf.fills``): a rectangle shape's fill, compared
  with the model's filled rectangles (paths of four corners) of the same colour;
* every **glyph**: the body's, which a floating drawing must not move, and a text box's
  (``read_render.score``, as the committed documents are).

A box agrees when every edge is within :data:`TOLERANCE` px, the precision Quartz writes
a matrix to.  ``--record`` writes Word's side to ``tests/fixtures/anchor-observations.json``
and ``tests/test_anchors.py`` holds the model to it offline.

Usage::

    python tools/read_anchor_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "anchor-observations.json"
TOLERANCE = 0.03


def documents() -> list[tuple[str, bytes]]:
    import make_anchor_probe
    import make_drawing_probe
    import make_story_anchor_probe
    import make_wrap_anchor_probe
    import make_wrap_stack_probe

    out = [(f"anchor-{s}", make_anchor_probe.build(s)) for s in make_anchor_probe.SETTINGS]
    out += [(f"drawing-{s}", make_drawing_probe.build(s)) for s in make_drawing_probe.SETTINGS]
    out += [(f"story-anchor-{s}", make_story_anchor_probe.build(s)) for s in make_story_anchor_probe.SETTINGS]
    out += [(f"wrap-anchor-{s}", make_wrap_anchor_probe.build(s)) for s in make_wrap_anchor_probe.SETTINGS]
    out += [(f"wrap-stack-{s}", make_wrap_stack_probe.build(s)) for s in make_wrap_stack_probe.SETTINGS]
    return out


def word_images(pdf: Path) -> list:
    """``[page, x0, y0, x1, y1, layer]`` of every picture Word drew, in paint order;
    ``layer`` is ``behind`` (before the page's first glyph), ``front`` (after its last)
    or ``between``."""
    import pymupdf

    out = []
    with pymupdf.open(str(pdf)) as document:
        for index, page in enumerate(document):
            log = page.get_bboxlog()
            texts = [k for k, (kind, _) in enumerate(log) if kind == "fill-text"]
            for k, (kind, box) in enumerate(log):
                if kind != "fill-image":
                    continue
                layer = ("behind" if not texts or k < texts[0] else "front" if k > texts[-1] else "between")
                out.append([index] + [round(v * 300 / 72, 3) for v in box] + [layer])
    return out


def model_images(layout) -> list:
    """The model's pictures of floating drawings, in paint order (:meth:`Page.paint`), as
    :func:`word_images`: each ``behind``, ``front`` or ``between`` the page's glyphs."""
    out = []
    for index, page in enumerate(layout.pages):
        sequence: list = []  # "text" or an image row, in paint order
        for step in page.paint():
            if step[0] == "lines":
                sequence += ["text" for line in step[1] if any(span.chars for span in line.spans)]
            elif step[0] == "floats":
                for placed in step[1]:
                    for p in placed.primitives:
                        if p.kind == "image":
                            sequence.append([index, float(p.x), float(p.y), float(p.x + p.width),
                                             float(p.y + p.height)])
                    sequence += ["text" for line in placed.lines if any(span.chars for span in line.spans)]
        texts = [k for k, item in enumerate(sequence) if item == "text"]
        for k, item in enumerate(sequence):
            if item != "text":
                layer = ("behind" if not texts or k < texts[0] else "front" if k > texts[-1] else "between")
                out.append(item + [layer])
    return out


_NUMBER = re.compile(r"-?\d+(?:\.\d+)?")


def model_fills(layout) -> list:
    """``[page, x0, y0, x1, y1, colour]`` of every floating drawing's filled path that is
    an axis-aligned rectangle (four corners), untransformed."""
    out = []
    for index, page in enumerate(layout.pages):
        for placed in page.floats:
            for p in placed.primitives:
                if p.kind != "path" or not p.fill or p.transform or not re.fullmatch(r"M[^MLZ]+(L[^MLZ]+){3}Z", p.d):
                    continue
                values = [float(v) for v in _NUMBER.findall(p.d)]
                xs, ys = values[0::2], values[1::2]
                if len({round(v, 3) for v in xs}) == 2 and len({round(v, 3) for v in ys}) == 2:
                    out.append([index, min(xs), min(ys), max(xs), max(ys), p.fill.upper()])
    return out


def compare_boxes(word: list, ours: list, *, keyed: bool = False) -> tuple[int, int, list]:
    """``(agreeing, scored, disagreements)``: per page, Word's boxes against the model's
    in order (``keyed``: each of Word's against the model's of the same colour)."""
    pages = sorted({b[0] for b in word} | {b[0] for b in ours})
    agree = scored = 0
    problems = []
    for page in pages:
        a = [b for b in word if b[0] == page]
        b = [m for m in ours if m[0] == page]
        if keyed:
            for w in a:
                scored += 1
                if any(m[5] == w[5] and max(abs(x - y) for x, y in zip(w[1:5], m[1:5])) <= TOLERANCE for m in b):
                    agree += 1
                else:
                    problems.append((page, w, [m for m in b if m[5] == w[5]][:2]))
            continue
        if len(a) != len(b):
            scored += max(len(a), len(b))
            problems.append((page, "count", a, b))
            continue
        for w, m in zip(a, b):
            scored += 1
            if max(abs(x - y) for x, y in zip(w[1:5], m[1:5])) <= TOLERANCE and w[5] == m[5]:
                agree += 1
            else:
                problems.append((page, w, [round(v, 3) if isinstance(v, float) else v for v in m]))
    return agree, scored, problems


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    """Every score of one document against Word's recorded side."""
    from docx2svg import _render

    glyph_score, rects, warnings = read_render.score(name, data, fonts, recorded["objects"], recorded["fills"])
    _documents, layout, _data, _fonts = _render(data, render_record.options(fonts))
    image_fills = [f for f in recorded["fills"] if f[0] < len(layout.pages)]
    images = compare_boxes(recorded["images"], model_images(layout))
    fills = compare_boxes([f for f in image_fills if any(f[0] == m[0] for m in model_fills(layout))],
                          model_fills(layout), keyed=True)
    return {"glyphs": read_render.row(glyph_score), "images": list(images[:2]), "fills": list(fills[:2]),
            "warnings": warnings, "problems": [*images[2], *fills[2]], "result": glyph_score}


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("names", nargs="*")
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--limit", type=int, default=30)
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    try:
        for name, data in documents():
            if args.names and name not in args.names:
                continue
            pdf = oracle.export(data, name=name)
            word = {"objects": read_render.word_objects(pdf), "fills": read_render.word_fills(pdf),
                    "images": word_images(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} pictures where Word drew them, in its order and layer: {result['images'][0]} / "
                  f"{result['images'][1]}; rectangle fills: {result['fills'][0]} / {result['fills'][1]}; "
                  f"warnings {result['warnings']}")
            if args.verbose:
                for problem in result["result"].problems[:args.limit]:
                    print("   ", problem.replace("|", ""))
                for problem in result["problems"][:args.limit]:
                    print("   ", problem)
            recorded[name] = {**word, "glyphs": result["glyphs"], "images_score": result["images"],
                              "fills_score": result["fills"], "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_anchor_probe.py and tools/make_drawing_probe.py: every "
                       "text object and filled rectangle (as render-observations.json records them), every picture "
                       "([page, x0, y0, x1, y1, layer], device px, in paint order; layer: painted before the page's "
                       "first glyph, after its last, or between), the renderer's scores against them, and every "
                       "face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_anchor_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
