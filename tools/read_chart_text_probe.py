#!/usr/bin/env python3
"""Score ``make_chart_text_probe.py``'s charts against Word: titles, axis titles and short
charts, as Word's PDF draws them and as docx2svg draws them.

The method is ``read_chart_probe.py``'s -- every span of the chart's text, every filled
shape and every stroked segment, read off Word's PDF by PyMuPDF and off the chart's SVG
fragment, paired within :data:`read_chart_probe.TOLERANCE` -- with the one thing that
probe never needed: **turned text**.  An axis title at the left reads upwards, and a text
is now compared by the midpoint of its baseline, in both directions, and by its angle:

* Word's: the span's origin (the start of its baseline) and direction, and its length
  along that direction from its characters' boxes;
* the model's: its anchor (``start``, ``middle`` or ``end`` of the advance, per
  ``text-anchor``) moved along its own direction by half of Word's length as needed.

A text agrees when its string, face, weight and device size are Word's, its angle is
Word's to a degree, and both coordinates of its midpoint are within the tolerance.

``--record`` writes Word's side and the scores to
``tests/fixtures/chart-text-observations.json``, which ``tests/test_chart_text.py``
holds the model to offline.

Usage::

    python tools/read_chart_text_probe.py [-v] [--record] [--case N]
"""

from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_chart_probe as base  # noqa: E402
import read_render  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "chart-text-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_chart_text_probe

    return [(f"chart-text-{s}", make_chart_text_probe.build(s)) for s in make_chart_text_probe.SETTINGS]


def word_side(pdf: Path) -> list[dict]:
    """What Word drew for each case: :func:`read_chart_probe.word_page`'s fills and
    strokes, and its texts as ``[text, font, size, mid_x, mid_y, angle]``."""
    import make_chart_text_probe
    import pymupdf

    with pymupdf.open(str(pdf)) as document:
        return [word_page(document[index]) for index in range(len(make_chart_text_probe.CASES))]


def word_page(page) -> dict:
    row = base.word_page(page)
    row["texts"] = []
    for block in page.get_text("rawdict")["blocks"]:
        for text_line in block.get("lines", []):
            dx, dy = text_line["dir"]
            for span in text_line["spans"]:
                if span["font"].startswith(base.BODY_FACE):
                    continue
                for chars in _pieces(span["chars"], dx, dy):
                    inked = [c for c in chars if c["c"].strip()]
                    if not inked:
                        continue
                    ox, oy = chars[0]["origin"]
                    # How far along the direction the last character's box reaches.
                    length = max(max((x - ox) * dx + (y - oy) * dy for x in (c["bbox"][0], c["bbox"][2])
                                     for y in (c["bbox"][1], c["bbox"][3])) for c in inked)
                    angle = math.degrees(math.atan2(-dy, dx))
                    row["texts"].append(["".join(c["c"] for c in chars).strip(), span["font"],
                                         base._r(span["size"]), base._r(ox + dx * length / 2),
                                         base._r(oy + dy * length / 2), base._r(angle), base._r(length)])
    return row


def _pieces(chars: list, dx: float, dy: float) -> list[list]:
    """A span's characters, split where one does not advance along the line past the one
    before -- two labels on one line (a plot of no height draws its zero and its top label
    there), which the PDF writes as one span.  Leading spaces are dropped."""
    pieces: list[list] = []
    for char in chars:
        if pieces:
            last = pieces[-1][-1]["origin"]
            if (char["origin"][0] - last[0]) * dx + (char["origin"][1] - last[1]) * dy > 0.01:
                pieces[-1].append(char)
                continue
        if char["c"].strip() or pieces:
            pieces.append([char])
    return pieces


def compare(word: dict, model: dict, tolerance: float = base.TOLERANCE) -> tuple[dict, list]:
    """``read_chart_probe.compare`` with texts by the midpoints of their baselines."""
    scores, problems = base.compare({"texts": [], "fills": word["fills"], "strokes": word["strokes"]},
                                    {"texts": [], "fills": model["fills"], "strokes": model["strokes"]}, tolerance)
    agree, used = 0, set()
    for text, font, size, mid_x, mid_y, angle, length in word["texts"]:
        family, bold = base.family_of(font)

        def midpoint(t):
            radians = math.radians(t[6])
            ux, uy = math.cos(radians), -math.sin(radians)
            shift = {"start": length / 2, "end": -length / 2}.get(t[3], 0.0)
            return t[4] + ux * shift, t[5] + uy * shift

        candidates = [(k, t) for k, t in enumerate(model["texts"]) if k not in used and t[0] == text]
        if not candidates:
            problems.append(("text", text, "not drawn"))
            continue
        k, best = min(candidates, key=lambda c: math.dist(midpoint(c[1]), (mid_x, mid_y)))
        used.add(k)
        mx, my = midpoint(best)
        ok_face = best[1].replace(" ", "").lower() == family.replace(" ", "").lower() and best[7] == bold
        ok_size = abs(base.device(best[2]) - size) < 0.06
        ok_angle = abs((best[6] - angle + 180) % 360 - 180) < 1
        dx, dy = mx - mid_x, my - mid_y
        if ok_face and ok_size and ok_angle and base._near(dx, 0, tolerance) and base._near(dy, 0, tolerance):
            agree += 1
        else:
            problems.append(("text", text, f"{family}{' bold' if bold else ''} {size} {angle:g} vs {best[1]}"
                             f"{' bold' if best[7] else ''} {base._r(base.device(best[2]))} {best[6]:g}, "
                             f"dx {dx:+.2f} dy {dy:+.2f}"))
    extra = [t for k, t in enumerate(model["texts"]) if k not in used]
    problems += [("text", t[0], "drawn, and not by Word") for t in extra]
    scores["text"] = [agree, len(word["texts"]) + len(extra)]
    return {key: scores[key] for key in ("text", "fills", "strokes")}, problems


def score(data: bytes, fonts, word: list[dict]) -> dict:
    """The model's charts against ``word``, per case and in total."""
    import render_record
    from docx2svg import _render

    _documents, layout, _data, _fonts = _render(data, render_record.options(fonts))
    model = base.model_side(layout)
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
    import make_chart_text_probe
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
            case = make_chart_text_probe.CASES[index]
            line = " ".join(f"{k} {a}/{c}" for k, (a, c) in found.items())
            print(f"  {index:2d} {case.family:6s} {case.note:52s} {line}")
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
