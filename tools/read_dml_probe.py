#!/usr/bin/env python3
"""Score DrawingML's paint against Word: ``make_dml_probe.py``'s colours, gradients,
dashes, patterns and theme style references, as Word's PDF draws them and as the shared
renderers (``ooxml_common.drawingml``, under its ``WORD`` rules) draw them here.

Word's PDF gives, per case (a page each):

* ``colour`` -- every swatch's fill colour, in reading order (``pymupdf``'s drawings);
* ``gradient`` -- every shading Word writes: its type, its coordinates in the shape's
  own frame (EMU from its top left: Quartz draws the shading through the shape's matrix),
  and its colour at 0, 1/4, 1/2, 3/4 and 1 (the profile's values; a two-stop gradient is
  written in a linear profile, see ``ooxml_common.drawingml.fill``);
* ``dash`` -- every stroked line's width, dash array and cap, and for a line Word drew
  as filled outlines (a round or square cap) the extents of its first dashes;
* ``pattern`` -- each pattern's tile origin and step;
* ``style`` -- the fill and outline colours and widths a theme reference gives;
* ``auto`` -- the colour Word draws a text box's automatic text in.

The model's side is read off the layout (its primitives and the page's definitions) in
the same terms, and :func:`compare` counts what agrees.  ``--record`` writes Word's side
and the scores to ``tests/fixtures/dml-observations.json``; ``tests/test_dml.py`` holds the
model to it offline.

Usage::

    python tools/read_dml_probe.py [-v] [--record]
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

OBSERVATIONS = read_render.FIXTURES / "dml-observations.json"
#: EMU per device px.
EMU = 914400 / 300
#: How near a gradient's coordinates must be (EMU: 0.05 device px).
TOLERANCE_EMU = 160
_NUMBER = re.compile(r"-?\d+(?:\.\d+)?(?:e-?\d+)?")


def documents() -> list[tuple[str, bytes]]:
    import make_dml_probe

    return [(f"dml-{s}", make_dml_probe.build(s)) for s in make_dml_probe.SETTINGS]


def _hex(rgb) -> str:
    return "".join(f"{round(v * 255):02X}" for v in rgb)


def word_side(pdf: Path) -> list[dict]:
    """What Word drew for each case (module docstring)."""
    import make_dml_probe
    import pymupdf

    out = []
    with pymupdf.open(str(pdf)) as document:
        for index, case in enumerate(make_dml_probe.CASES):
            page = document[index]
            drawings = [d for d in page.get_drawings() if d["rect"].y0 > 50]
            row: dict = {"family": case.family, "note": case.note}
            if case.family in ("colour", "style"):
                filled = sorted((d for d in drawings if d.get("fill") is not None and d["rect"].width > 20),
                                key=lambda d: (round(d["rect"].y0), d["rect"].x0))
                row["fills"] = [_hex(d["fill"]) for d in filled]
                row["strokes"] = [[_hex(d["color"]), round(d["width"], 3)] for d in drawings
                                  if d.get("color") is not None and d.get("width")]
            if case.family == "gradient":
                row["shadings"] = _shadings(document, page)
            if case.family == "dash":
                # In drawing order: a stroked line, or a line Word drew as filled dashes.
                row["lines"] = [["s", round(d["width"], 3), _dashes(d.get("dashes")), d["lineCap"][0]]
                                if d["type"] == "s" else ["f", _spans(d)]
                                for d in drawings if d["type"] == "s" or d["rect"].height > 0.5]
            if case.family == "pattern":
                row["tiles"] = _tiles(document, page)
            if case.family == "auto":
                row["text"] = [f"{span['color']:06X}" for block in page.get_text("dict")["blocks"]
                               for text_line in block.get("lines", []) for span in text_line["spans"]
                               if span["text"].startswith("Auto")]
            out.append(row)
    return out


def _dashes(value) -> list[float]:
    return [float(v) for v in _NUMBER.findall((value or "").split("]")[0])]


def _spans(drawing) -> list[list[float]]:
    """The x extents of a flattened dashed line's first dashes (pt)."""
    if drawing["items"][0][0] == "re":
        spans = sorted((item[1].x0, item[1].x1) for item in drawing["items"] if item[0] == "re")
    else:
        spans, points, first_y = [], [], drawing["items"][0][1].y
        for item in drawing["items"]:
            points += [p.x for p in item[1:] if hasattr(p, "x")]
            if item[0] == "c" and len(points) >= 8 and abs(item[-1].x - points[0]) < 1e-6 \
                    and abs(item[-1].y - first_y) < 5:
                spans.append((min(points), max(points)))
                points = []
        spans.sort()
    return [[round(a, 3), round(b, 3)] for a, b in spans[:6]]


def _shadings(document, page) -> list[dict]:
    out = []
    kind, value = document.xref_get_key(page.xref, "Resources/Shading")
    for name, ref in re.findall(r"/(\w+) (\d+) 0 R", value if kind != "null" else ""):
        body = document.xref_object(int(ref), compressed=True)
        function = int(re.search(r"/Function (\d+) 0 R", body).group(1))
        samples = document.xref_stream(function)
        count = int(re.search(r"/Size\[(\d+)\]", document.xref_object(function, compressed=True)).group(1))
        width = len(samples) // count
        colour_space = document.xref_object(int(re.search(r"/ColorSpace (\d+) 0 R", body).group(1)), compressed=True)
        profile = document.xref_stream(int(re.search(r"(\d+) 0 R", colour_space).group(1)))
        out.append({
            "name": name,
            "type": int(re.search(r"/ShadingType (\d)", body).group(1)),
            "coords": [float(v) for v in _NUMBER.findall(re.search(r"/Coords\[([^\]]*)\]", body).group(1))],
            "linear": b"Generic HDR" in profile,
            "samples": [samples[k * width:(k + 1) * width].hex().upper()
                        for k in (round(t * (count - 1)) for t in (0, 0.25, 0.5, 0.75, 1))],
        })
    return out


def _tiles(document, page) -> list[list[float]]:
    out = []
    kind, value = document.xref_get_key(page.xref, "Resources/Pattern")
    for _name, ref in re.findall(r"/(\w+) (\d+) 0 R", value if kind != "null" else ""):
        body = document.xref_object(int(ref), compressed=True)
        matrix = [float(v) for v in _NUMBER.findall(re.search(r"/Matrix\[([^\]]*)\]", body).group(1))]
        step = float(re.search(r"/XStep ([\d.]+)", body).group(1))
        out.append([matrix[4], round(page.rect.height - matrix[5], 3), step * matrix[0]])
    return out


# -- the model's side -----------------------------------------------------------------


def model_side(layout) -> list[dict]:
    """The model's paint for each case, in :func:`word_side`'s terms."""
    import make_dml_probe

    out = []
    for index, case in enumerate(make_dml_probe.CASES):
        page = layout.pages[index]
        found_defs = getattr(page, "defs", None)
        defs = {re.search(r'id="([^"]+)"', d).group(1): d for d in (found_defs.defs if found_defs else [])}
        paths = [p for placed in page.floats for p in placed.primitives if p.kind == "path"]
        box = page.floats[0] if page.floats else None
        row: dict = {}
        if case.family in ("colour", "style"):
            filled = sorted((p for p in paths if p.fill), key=lambda p: _first_point(p.d)[::-1])
            row["fills"] = [p.fill for p in filled]
            row["strokes"] = [[p.stroke, round(float(p.stroke_width) * 72 / 300, 3)] for p in paths if p.stroke]
        if case.family == "gradient" and box is not None:
            row["gradients"] = []
            for p in paths:
                found = re.search(r'url\(#([^)]+)\)', getattr(p, "attrs", ""))
                if found and found.group(1) in defs:
                    row["gradients"].append(_gradient_emu(defs[found.group(1)], box))
        if case.family == "dash":
            row["strokes"] = [_stroke(getattr(p, "attrs", "") or f'stroke-width="{float(p.stroke_width)}"')
                              for p in paths if p.stroke]
        if case.family == "pattern":
            row["tiles"] = [_pattern(defs[k]) for k in defs if defs[k].startswith("<pattern")]
        if case.family == "auto":
            row["text"] = [(span.color or "000000").upper() for placed in page.floats for text_line in placed.lines
                           for span in text_line.spans if "".join(span.chars).startswith("Auto")]
        out.append(row)
    return out


def _first_point(d: str) -> tuple[float, float]:
    values = [float(v) for v in _NUMBER.findall(d)[:2]]
    return values[0], values[1]


def _gradient_emu(definition: str, box) -> dict:
    """A gradient's geometry in EMU from the drawing's top left, as Word's shading."""
    values = {k: float(v) for k, v in re.findall(r' (x1|y1|x2|y2|cx|cy|r|fx|fy)="([^"%]+)"', definition.split(">")[0])}
    def to_emu(key, origin):
        return (values[key] - float(origin)) * EMU
    if definition.startswith("<linearGradient"):
        return {"type": 2, "coords": [to_emu("x1", box.x), to_emu("y1", box.y), to_emu("x2", box.x),
                                      to_emu("y2", box.y)]}
    if "cx" in values and "fx" in values and "gradientUnits" in definition:
        return {"type": 3, "coords": [to_emu("fx", box.x), to_emu("fy", box.y), 0, to_emu("cx", box.x),
                                      to_emu("cy", box.y), values["r"] * EMU]}
    return {"type": 0, "coords": []}


def _stroke(attrs: str) -> list:
    width = float(re.search(r'stroke-width="([^"]+)"', attrs).group(1)) * 72 / 300
    dashes = re.search(r'stroke-dasharray="([^"]+)"', attrs)
    cap = re.search(r'stroke-linecap="([^"]+)"', attrs)
    return [round(width, 3), [round(float(v) * 72 / 300, 3) for v in dashes.group(1).split()] if dashes else [],
            {"butt": 0, "round": 1, "square": 2}.get(cap.group(1) if cap else "butt")]


def _pattern(definition: str) -> list[float]:
    width = float(re.search(r' width="([^"]+)"', definition).group(1))
    x = float((re.search(r' x="([^"]+)"', definition.split(">")[0]) or [0, 0])[1])
    y = float((re.search(r' y="([^"]+)"', definition.split(">")[0]) or [0, 0])[1])
    return [x * 72 / 300, y * 72 / 300, width * 72 / 300]


# -- comparison -----------------------------------------------------------------------


def compare(word: list[dict], model: list[dict]) -> dict:
    """What agrees, per family: ``[agree, compared]``, and the disagreements."""
    import make_dml_probe

    scores: dict = {}
    problems: list[str] = []

    def count(family: str, ok: bool, what: str) -> None:
        agree, total = scores.get(family, [0, 0])
        scores[family] = [agree + int(ok), total + 1]
        if not ok:
            problems.append(f"{family}: {what}")

    for index, (w, m) in enumerate(zip(word, model)):
        family, note = w["family"], w["note"]
        if family in ("colour", "style"):
            for k, colour in enumerate(w["fills"]):
                ours = m["fills"][k] if k < len(m.get("fills", [])) else None
                count(family, ours == colour, f"case {index} ({note}) fill {k}: Word {colour}, model {ours}")
            if family == "style":
                for k, (colour, width) in enumerate(w["strokes"]):
                    ours = m["strokes"][k] if k < len(m["strokes"]) else None
                    count(family, ours is not None and ours[0] == colour and abs(ours[1] - width) < 0.01,
                          f"case {index} ({note}) outline {k}: Word {colour} {width}, model {ours}")
        if family == "gradient":
            ours = m.get("gradients") or []
            main = _main_shading(w["shadings"], make_dml_probe.CASES[index])
            if main is None:
                count(family, not ours or ours[0]["type"] == 0, f"case {index} ({note}): Word drew a picture")
                continue
            coords = main["coords"]
            if main["type"] == 2 and _reversed(main["samples"][0]):
                # Word may write the run the other way with its colours reversed: the
                # probe's gradients all start red, so a run starting blue is reversed.
                coords = coords[2:] + coords[:2]
            ok = bool(ours) and ours[0]["type"] == main["type"] and all(
                abs(a - b) <= TOLERANCE_EMU for a, b in zip(ours[0]["coords"], coords))
            count(family, ok, f"case {index} ({note}): Word {main['type']} {main['coords']}, "
                              f"model {ours[0] if ours else None}")
        if family == "dash":
            ours = m.get("strokes", [])
            for k, line in enumerate(w["lines"]):
                mine = ours[k] if k < len(ours) else None
                if line[0] == "s":
                    ok = mine is not None and abs(mine[0] - line[1]) < 0.01 and \
                        [round(v, 2) for v in mine[1]] == [round(v, 2) for v in line[2]]
                else:
                    measured = _measured(line[1])
                    shown = _visual(mine) if mine is not None else []
                    ok = bool(measured) and len(shown) == len(measured) and \
                        all(abs(a - b) < 0.02 for a, b in zip(shown, measured))
                count(family, ok, f"case {index} ({note}) line {k}: Word {line}, model {mine}")
        if family == "auto":
            for k, colour in enumerate(w["text"]):
                ours = m["text"][k] if k < len(m["text"]) else None
                count(family, ours == colour, f"case {index} ({note}) text {k}: Word {colour}, model {ours}")
        if family == "pattern":
            for tile in w["tiles"]:
                # Registered to the page's corner: the origin on the 8 pt lattice.
                count(family, abs(tile[0] / 8 - round(tile[0] / 8)) < 1e-3 and abs(tile[1] / 8 - round(tile[1] / 8)) < 1e-3
                      and bool(m.get("tiles"))
                      and all(t[:2] == [0.0, 0.0] and abs(t[2] - 8) < 1e-3 for t in m["tiles"]),
                      f"case {index} ({note}): Word's tile at {tile}, model's {m.get('tiles')}")
    return {"scores": scores, "problems": problems}


def _visual(stroke: list) -> list[float]:
    """The dashes and gaps a stroke shows (pt), from its second dash on, for two turns of
    its pattern: a round cap reaches half a width past each end of a dash."""
    width, dashes, cap = stroke
    if len(dashes) < 2:
        return []
    grow = width if cap == 1 else 0
    shown = [round(value + grow if k % 2 == 0 else value - grow, 2) for k, value in enumerate(dashes)]
    turn = (shown[2:] + shown[:2]) * 3
    return turn[:6]


def _measured(spans: list) -> list[float]:
    """The same of a line Word drew as filled dashes, from its second dash (the first
    carries the line's start cap)."""
    out = []
    for k in range(1, min(len(spans) - 1, 4)):
        out += [round(spans[k][1] - spans[k][0], 2), round(spans[k + 1][0] - spans[k][1], 2)]
    return out[:6]


def _reversed(sample: str) -> bool:
    return int(sample[4:6], 16) > int(sample[0:2], 16)


def _main_shading(shadings: list[dict], case) -> dict | None:
    """The shading that spans the shape.  Word pads a linear one with mirrored copies
    before and after it; the one whose run holds the box's centre is it."""
    if not shadings:
        return None
    width, height = case.cx // 635 * 635, case.cy // 635 * 635
    for shading in shadings:
        if shading["type"] != 2:
            return shading
        x1, y1, x2, y2 = shading["coords"]
        length = (x2 - x1) ** 2 + (y2 - y1) ** 2
        t = ((width / 2 - x1) * (x2 - x1) + (height / 2 - y1) * (y2 - y1)) / length if length else -1
        if 0 <= t <= 1:
            return shading
    return shadings[0]


def score(data: bytes, fonts, word: list[dict]) -> dict:
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    return compare(word, model_side(layout))


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    try:
        for name, data in documents():
            pdf = oracle.export(data, name=name)
            word = word_side(pdf)
            result = score(data, render_record.RecordingFonts(data, faces), word)
            print(name, " ".join(f"{family} {a}/{b}" for family, (a, b) in sorted(result["scores"].items())))
            if args.verbose:
                for problem in result["problems"]:
                    print("   ", problem)
            recorded[name] = {"word": word, "scores": result["scores"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_dml_probe.py: per case its fill colours, its "
                       "gradients' shadings (type, coordinates in the shape's frame in EMU, colour samples), its "
                       "strokes' widths, dash arrays and caps, its patterns' tiles, and the renderer's scores "
                       "against them, and every face number the renderer asked for. Measurements only; "
                       "regenerate with tools/read_dml_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
