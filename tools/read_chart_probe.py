#!/usr/bin/env python3
"""Score charts against Word: ``make_chart_probe.py``'s charts as Word's PDF draws them and
as docx2svg draws them (``ooxml-common``'s chart layout and scene renderers, under Word's
rules).

Word's side, per case (a page each), read by PyMuPDF:

* ``texts`` -- every span of the chart's text (the probe's own text is Georgia and the
  chart's never is): the string, the face's PostScript name, the size, the advance box's
  left and right and the baseline (pt);
* ``fills`` -- every filled shape: its colour and box (a bar, a legend key, a slice, the
  chart's and the plot area's backgrounds); a path of several rectangles is one fill each;
* ``strokes`` -- every stroked straight segment: its colour, width and ends (gridlines,
  axes, tick marks, series lines, outlines), and a stroked curve as its box.

The model's side is read off the layout -- the chart's SVG fragment (:mod:`docx2svg.chart`)
and the transform that places it -- in the same terms, by :func:`model_side`, which needs
nothing but the standard library.  :func:`compare` pairs them:

* a text agrees when one of the model's with the same string has the same face, its size
  is Word's (Word's PDF writes the size on the device grid: 10 pt is 10.08), and its
  anchor -- the left, middle or right of the advance box, as the model's ``text-anchor``
  says -- and baseline are within :data:`TOLERANCE`;
* a fill when one of the model's has its colour and every edge within :data:`TOLERANCE`;
* a stroke when one of the model's has its colour, its width to 0.05 pt and both ends
  within :data:`TOLERANCE`;

and whatever the model draws that pairs with nothing of Word's counts against it too: a
family's score is what agreed out of Word's items and the model's left over.

``--record`` writes Word's side and the scores to ``tests/fixtures/chart-observations.json``,
which ``tests/test_chart.py`` holds the model to offline.

Usage::

    python tools/read_chart_probe.py [-v] [--record] [--case N]
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from pathlib import Path
from xml.etree import ElementTree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "chart-observations.json"
#: How near a position must be, pt (two device px).
TOLERANCE = 0.5
#: Points per device px.
PT = 72 / 300
_NUMBER = re.compile(r"-?\d+(?:\.\d+)?(?:[eE]-?\d+)?")
SVG = "{http://www.w3.org/2000/svg}"
#: Bullet characters, which Word's PDF writes in one span with their paragraph's text.
BULLETS = "\u2022\u25aa\u2013\u00a7"
#: The face the probe's own text is in, which no chart text is.
BODY_FACE = "Georgia"


def documents() -> list[tuple[str, bytes]]:
    import make_chart_probe

    return [(f"chart-{s}", make_chart_probe.build(s)) for s in make_chart_probe.SETTINGS]


def _hex(rgb) -> str:
    return "".join(f"{round(v * 255):02X}" for v in rgb)


def _r(value: float) -> float:
    return round(float(value), 3)


# -- Word's side ----------------------------------------------------------------------


def word_side(pdf: Path) -> list[dict]:
    """What Word drew for each case (module docstring)."""
    import make_chart_probe
    import pymupdf

    with pymupdf.open(str(pdf)) as document:
        return [word_page(document[index]) for index in range(len(make_chart_probe.CASES))]


def word_page(page) -> dict:
    """One page's drawing that is not the probe's own text, in the module docstring's
    terms."""
    row: dict = {"texts": [], "fills": [], "strokes": []}
    for block in page.get_text("rawdict")["blocks"]:
        for text_line in block.get("lines", []):
            for span in text_line["spans"]:
                text = "".join(c["c"] for c in span["chars"]).strip()
                if not text or span["font"].startswith(BODY_FACE):
                    continue
                direction = text_line["dir"]
                angle = _r(math.degrees(math.atan2(-direction[1], direction[0])))
                # A bulleted paragraph's span starts with its bullet, drawn apart from its
                # text: two items, as a renderer that writes them apart has them.
                pieces = [span["chars"]]
                if span["chars"][0]["c"] in BULLETS and len(span["chars"]) > 1:
                    pieces = [span["chars"][:1], span["chars"][1:]]
                for piece in pieces:
                    chars = [c for c in piece if c["c"].strip()]
                    if not chars:
                        continue
                    row["texts"].append(["".join(c["c"] for c in piece).strip(), span["font"], _r(span["size"]),
                                         _r(chars[0]["bbox"][0]), _r(chars[-1]["bbox"][2]),
                                         _r(span["origin"][1]), angle])
    for drawing in page.get_drawings():
        kind = drawing["type"]
        if "f" in kind and drawing.get("fill") is not None:
            colour = _hex(drawing["fill"])
            for box in _boxes(drawing):
                row["fills"].append([colour] + box)
        if "s" in kind and drawing.get("color") is not None:
            colour, width = _hex(drawing["color"]), _r(drawing["width"] or 0)
            for segment in _segments(drawing):
                row["strokes"].append([colour, width] + segment)
    return row


def _bezier(p0, p1, p2, p3, steps: int = 24) -> list[tuple[float, float]]:
    """Points along a cubic, for its tight box (a box of the control points is not one)."""
    out = []
    for k in range(steps + 1):
        t = k / steps
        u = 1 - t
        out.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                    u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return out


def _tight(items) -> list[float]:
    points = []
    for item in items:
        if item[0] == "c":
            points += _bezier(*[(p.x, p.y) for p in item[1:5]])
        elif item[0] == "l":
            points += [(item[1].x, item[1].y), (item[2].x, item[2].y)]
        elif item[0] == "re":
            points += [(item[1].x0, item[1].y0), (item[1].x1, item[1].y1)]
        elif item[0] == "qu":
            points += [(item[1].rect.x0, item[1].rect.y0), (item[1].rect.x1, item[1].rect.y1)]
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return [_r(min(xs)), _r(min(ys)), _r(max(xs)), _r(max(ys))]


def _boxes(drawing) -> list[list[float]]:
    """A filled path's shapes: each rectangle on its own, else the path's tight box."""
    items = drawing["items"]
    if all(item[0] == "re" for item in items):
        return [[_r(item[1].x0), _r(item[1].y0), _r(item[1].x1), _r(item[1].y1)] for item in items]
    if all(item[0] == "qu" for item in items):
        return [[_r(item[1].rect.x0), _r(item[1].rect.y0), _r(item[1].rect.x1), _r(item[1].rect.y1)]
                for item in items]
    return [_tight(items)]


def _segments(drawing) -> list[list[float]]:
    """A stroked path's straight segments, and a curved one's box as ``[x0, y0, x1, y1,
    "box"]``."""
    items = drawing["items"]
    if any(item[0] == "c" for item in items):
        return [_tight(items) + ["box"]]
    out = []
    for item in items:
        if item[0] == "l":
            out.append([_r(item[1].x), _r(item[1].y), _r(item[2].x), _r(item[2].y)])
        elif item[0] == "re":
            r = item[1]
            corners = [(r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)]
            out += [[_r(a[0]), _r(a[1]), _r(b[0]), _r(b[1])] for a, b in zip(corners, corners[1:] + corners[:1])]
        elif item[0] == "qu":
            q = item[1]
            corners = [(q.ul.x, q.ul.y), (q.ur.x, q.ur.y), (q.lr.x, q.lr.y), (q.ll.x, q.ll.y)]
            out += [[_r(a[0]), _r(a[1]), _r(b[0]), _r(b[1])] for a, b in zip(corners, corners[1:] + corners[:1])]
    return out


# -- the model's side -----------------------------------------------------------------


def _multiply(a, b):
    return (a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3],
            a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5])


def _transform(text: str | None):
    """An SVG ``transform`` as a matrix ``(a, b, c, d, e, f)``."""
    matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for name, args in re.findall(r"(\w+)\s*\(([^)]*)\)", text or ""):
        values = [float(v) for v in _NUMBER.findall(args)]
        if name == "translate":
            step = (1, 0, 0, 1, values[0], values[1] if len(values) > 1 else 0)
        elif name == "scale":
            step = (values[0], 0, 0, values[1] if len(values) > 1 else values[0], 0, 0)
        elif name == "rotate":
            angle = math.radians(values[0])
            cos, sin = math.cos(angle), math.sin(angle)
            step = (cos, sin, -sin, cos, 0, 0)
            if len(values) == 3:
                cx, cy = values[1], values[2]
                step = _multiply(_multiply((1, 0, 0, 1, cx, cy), step), (1, 0, 0, 1, -cx, -cy))
        elif name == "matrix":
            step = tuple(values)
        else:
            continue
        matrix = _multiply(matrix, step)
    return matrix


def _apply(matrix, x: float, y: float) -> tuple[float, float]:
    return matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]


def _scale_of(matrix) -> float:
    return math.sqrt(abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]))


def _paint(element, inherited: dict) -> dict:
    out = dict(inherited)
    for key in ("fill", "stroke", "stroke-width", "font-size", "font-family", "text-anchor", "font-weight"):
        if element.get(key) is not None:
            out[key] = element.get(key)
    return out


def _colour(value: str | None) -> str | None:
    if not value or value == "none" or not value.startswith("#"):
        return None
    return value[1:].upper()


def _path_points(d: str) -> tuple[list[list[tuple[float, float]]], bool]:
    """A path's subpaths as points (absolute ``M L H V C Z`` only, as the renderers write),
    and whether it has a curve."""
    tokens = re.findall(r"[MLHVCZmlhvcz]|-?\d+(?:\.\d+)?(?:[eE]-?\d+)?", d)
    subpaths: list[list[tuple[float, float]]] = []
    current: list[tuple[float, float]] = []
    command, curved, k = None, False, 0
    x = y = 0.0
    while k < len(tokens):
        token = tokens[k]
        if token.isalpha():
            command = token
            k += 1
            if command in "Zz":
                if current:
                    current.append(current[0])
                continue
            continue
        if command in ("M", "L"):
            x, y = float(tokens[k]), float(tokens[k + 1])
            k += 2
            if command == "M":
                if current:
                    subpaths.append(current)
                current = [(x, y)]
                command = "L"
            else:
                current.append((x, y))
        elif command == "H":
            x = float(tokens[k])
            k += 1
            current.append((x, y))
        elif command == "V":
            y = float(tokens[k])
            k += 1
            current.append((x, y))
        elif command == "C":
            curved = True
            points = [(float(tokens[k + j]), float(tokens[k + j + 1])) for j in (0, 2, 4)]
            k += 6
            current += _bezier((x, y), *points)[1:]
            x, y = points[-1]
        else:
            k += 1
    if current:
        subpaths.append(current)
    return subpaths, curved


def model_chart(markup: str, transform: str) -> dict:
    """The chart fragment ``markup`` placed by ``transform`` (device px), as texts, fills
    and strokes in pt (module docstring)."""
    root = ElementTree.fromstring(f'<g xmlns="http://www.w3.org/2000/svg" transform="{transform}">{markup}</g>')
    out: dict = {"texts": [], "fills": [], "strokes": []}
    page = (PT, 0, 0, PT, 0, 0)

    def walk(element, matrix, paint):
        matrix = _multiply(matrix, _transform(element.get("transform")))
        paint = _paint(element, paint)
        tag = element.tag.replace(SVG, "")
        if tag == "text":
            _text(element, matrix, paint)
            return
        points_list, curved, closed = [], False, False
        if tag == "rect":
            x, y = float(element.get("x", 0)), float(element.get("y", 0))
            w, h = float(element.get("width", 0)), float(element.get("height", 0))
            points_list = [[(x, y), (x + w, y), (x + w, y + h), (x, y + h), (x, y)]]
            closed = True
        elif tag == "line":
            points_list = [[(float(element.get("x1", 0)), float(element.get("y1", 0))),
                            (float(element.get("x2", 0)), float(element.get("y2", 0)))]]
        elif tag == "path":
            points_list, curved = _path_points(element.get("d", ""))
            closed = True
        elif tag in ("polygon", "polyline"):
            values = [float(v) for v in _NUMBER.findall(element.get("points", ""))]
            points = list(zip(values[0::2], values[1::2]))
            points_list = [points + (points[:1] if tag == "polygon" else [])]
            closed = True
        elif tag in ("circle", "ellipse"):
            cx, cy = float(element.get("cx", 0)), float(element.get("cy", 0))
            rx = float(element.get("r", element.get("rx", 0)))
            ry = float(element.get("r", element.get("ry", 0)))
            points_list = [[(cx + rx * math.cos(k * math.pi / 32), cy + ry * math.sin(k * math.pi / 32))
                            for k in range(65)]]
            curved, closed = True, True
        if points_list:
            placed = [[_apply(matrix, *p) for p in points] for points in points_list]
            fill = _colour(paint.get("fill", "#000000"))
            if fill and closed and tag != "line":
                for points in placed:
                    xs, ys = [p[0] for p in points], [p[1] for p in points]
                    out["fills"].append([fill, _r(min(xs)), _r(min(ys)), _r(max(xs)), _r(max(ys))])
            stroke = _colour(paint.get("stroke"))
            if stroke:
                width = _r(float(paint.get("stroke-width", 1)) * _scale_of(matrix))
                for points in placed:
                    if curved:
                        xs, ys = [p[0] for p in points], [p[1] for p in points]
                        out["strokes"].append([stroke, width, _r(min(xs)), _r(min(ys)), _r(max(xs)), _r(max(ys)),
                                               "box"])
                        continue
                    for a, b in zip(points, points[1:]):
                        if a != b:
                            out["strokes"].append([stroke, width, _r(a[0]), _r(a[1]), _r(b[0]), _r(b[1])])
        for item in element:
            walk(item, matrix, paint)

    def _text(element, matrix, paint):
        base_x, base_y = float(element.get("x", 0)), float(element.get("y", 0))
        y = base_y
        for span in list(element) or [element]:
            here = _paint(span, paint)
            if span.get("dy"):
                y += float(span.get("dy"))
            x = float(span.get("x", base_x))
            text = (span.text or "").strip()
            if not text:
                continue
            ax, ay = _apply(matrix, x, y)
            angle = math.degrees(math.atan2(-matrix[1], matrix[0]))
            size = float(here.get("font-size", 16)) * _scale_of(matrix)
            family = (here.get("font-family") or "").split(",")[0].strip().strip("'\"")
            out["texts"].append([text, family, _r(size), here.get("text-anchor", "start"), _r(ax), _r(ay), _r(angle),
                                 here.get("font-weight") == "bold"])

    walk(root, page, {})
    return out


def model_side(layout, what: str = "chart") -> list[dict]:
    """The model's charts (or diagrams: ``what``) for each case, in :func:`word_side`'s
    terms."""
    out = []
    for page in layout.pages:
        row: dict = {"texts": [], "fills": [], "strokes": []}
        for placed in page.floats:
            for primitive in placed.primitives:
                if primitive.kind == "markup" and primitive.what == what:
                    found = model_chart(primitive.markup, primitive.transform)
                    for key in row:
                        row[key] += found[key]
        out.append(row)
    return out


# -- comparing ------------------------------------------------------------------------


def family_of(postscript: str) -> tuple[str, bool]:
    """A PDF font name as ``(family, bold)``: ``ArialMT`` -> Arial, ``Aptos-Bold`` ->
    Aptos, bold; ``Aptos-Display`` -> Aptos Display."""
    name = postscript.split("+", 1)[-1]
    name = re.sub(r"(PSMT|MT|PS)$", "", name)
    bold = bool(re.search(r"-(Bold|Black|Heavy)", name))
    name = re.sub(r"-(Bold|Italic|BoldItalic|Regular|Roman)$", "", name)
    name = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", name.replace("-", " "))
    return name, bold


def device(size: float) -> float:
    """A size in pt as Word's PDF writes it: on the device grid, 300 an inch."""
    return round(size / PT) * PT


def _near(a: float, b: float, tolerance: float = TOLERANCE) -> bool:
    return abs(a - b) <= tolerance + 1e-9


def compare(word: dict, model: dict, tolerance: float = TOLERANCE) -> tuple[dict, list]:
    """``{family: [agree, compared]}`` and what disagreed, for one case."""
    scores: dict = {}
    problems: list = []
    # Texts: each of Word's against the nearest unused one of the model's with its string.
    agree, used = 0, set()
    for text, font, size, left, right, baseline, angle in word["texts"]:
        family, bold = family_of(font)
        candidates = [(k, t) for k, t in enumerate(model["texts"]) if k not in used and t[0] == text]
        if not candidates:
            problems.append(("text", text, "not drawn"))
            continue

        def anchor(t):
            return left if t[3] == "start" else right if t[3] == "end" else (left + right) / 2

        k, best = min(candidates, key=lambda c: abs(anchor(c[1]) - c[1][4]) + abs(baseline - c[1][5]))
        used.add(k)
        ok_face = best[1].replace(" ", "").lower() == family.replace(" ", "").lower() and best[7] == bold
        ok_size = abs(device(best[2]) - size) < 0.06
        dx, dy = best[4] - anchor(best), best[5] - baseline
        if ok_face and ok_size and _near(dx, 0, tolerance) and _near(dy, 0, tolerance):
            agree += 1
        else:
            problems.append(("text", text, f"{family}{' bold' if bold else ''} {size} vs {best[1]}"
                             f"{' bold' if best[7] else ''} {_r(device(best[2]))}, dx {dx:+.2f} dy {dy:+.2f}"))
    extra = [t for k, t in enumerate(model["texts"]) if k not in used]
    problems += [("text", t[0], "drawn, and not by Word") for t in extra]
    scores["text"] = [agree, len(word["texts"]) + len(extra)]
    for key, width in (("fills", 1), ("strokes", 2)):
        agree, used = 0, set()
        for item in word[key]:
            colour = item[0]
            ends = item[width:]
            best, best_k, best_off = None, None, None
            for k, other in enumerate(model[key]):
                if k in used or not _same_colour(other[0], colour):
                    continue
                if key == "strokes" and abs(other[1] - item[1]) > 0.05:
                    continue
                if len(other[width:]) != len(ends):
                    continue
                off = _offset(ends, other[width:], key == "strokes")
                if best_off is None or off < best_off:
                    best, best_k, best_off = other, k, off
            if best is not None and best_off <= tolerance + 1e-9:
                used.add(best_k)
                agree += 1
            else:
                problems.append((key, item, "nearest " + (f"{best} off {best_off:.2f}" if best else "none")))
        extra = [other for k, other in enumerate(model[key]) if k not in used]
        problems += [(key, other, "drawn, and not by Word") for other in extra]
        scores[key] = [agree, len(word[key]) + len(extra)]
    return scores, problems


def _same_colour(a: str, b: str) -> bool:
    return all(abs(int(a[k:k + 2], 16) - int(b[k:k + 2], 16)) <= 1 for k in (0, 2, 4))


def _offset(a, b, unordered: bool) -> float:
    numbers_a = [v for v in a if not isinstance(v, str)]
    numbers_b = [v for v in b if not isinstance(v, str)]
    straight = max(abs(x - y) for x, y in zip(numbers_a, numbers_b))
    if unordered and len(numbers_a) == 4:
        swapped = numbers_b[2:] + numbers_b[:2]
        straight = min(straight, max(abs(x - y) for x, y in zip(numbers_a, swapped)))
    return straight


def score(data: bytes, fonts, word: list[dict]) -> dict:
    """The model's charts against ``word`` (:func:`word_side`), per case and in total."""
    import render_record
    from docx2svg import _render

    _documents, layout, _data, _fonts = _render(data, render_record.options(fonts))
    model = model_side(layout)
    totals: dict = {}
    cases, problems = [], {}
    for index, (word_row, model_row) in enumerate(zip(word, model)):
        found, wrong = compare(word_row, model_row)
        cases.append(found)
        if wrong:
            problems[index] = wrong
        for key, (agree, compared) in found.items():
            totals.setdefault(key, [0, 0])
            totals[key][0] += agree
            totals[key][1] += compared
    return {"scores": totals, "cases": cases, "problems": problems}


def main(argv: list[str] | None = None) -> int:
    import make_chart_probe
    import oracle
    import render_record

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--case", type=int, action="append")
    options = parser.parse_args(argv)
    recorded = {"faces": {}, "documents": {}}
    for name, data in documents():
        pdf = oracle.export(data, name=name)
        word = word_side(pdf)
        fonts = render_record.RecordingFonts(data, recorded["faces"])
        result = score(data, fonts, word)
        print(name, " ".join(f"{k} {a}/{c}" for k, (a, c) in result["scores"].items()))
        for index, found in enumerate(result["cases"]):
            if options.case and index not in options.case:
                continue
            case = make_chart_probe.CASES[index]
            line = " ".join(f"{k} {a}/{c}" for k, (a, c) in found.items())
            print(f"  {index:2d} {case.family:7s} {case.note:40s} {line}")
            if options.verbose:
                for problem in result["problems"].get(index, []):
                    print("       ", *problem)
        recorded["documents"][name] = {"word": word, "scores": result["scores"]}
    if options.record:
        render_record.dump(OBSERVATIONS, recorded)
        print("recorded", OBSERVATIONS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
