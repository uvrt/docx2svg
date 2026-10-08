#!/usr/bin/env python3
"""Read ``make_wrap_kern_probe.py``: where Word wrapped each word, and which kerning that
takes -- the face's legacy ``kern`` table, its ``GPOS`` ``kern`` feature, or none.

Every paragraph's lines are read off Word's PDF (``quartz_pdf``); a word that fits is one
line, and its width is the paragraph mark's pen x less the margin, exact to 1e-4 px.  Each
model -- plain advances (``U``), the legacy table's pairs (``L``), ``GPOS``'s (``G``) --
predicts one line when the word is no wider than the text width the paragraph's indent
leaves, and is scored against what Word did, kerning on and off.

With ``--record``, writes the cases, Word's lines and widths, and the font units each
model needs (advances and pairs, read from the installed faces with fontTools) to
``tests/fixtures/wrap-kern-observations.json``; ``tests/test_wrap_kern.py`` holds the
rule to them offline.

Usage::

    python tools/read_wrap_kern_probe.py [--record]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_wrap_kern_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "wrap-kern-observations.json"
PX_PER_PT = 300 / 72
LEFT_PX = probe.MARGIN / 20 * PX_PER_PT


def observe() -> list[dict]:
    """Every case with the lines Word drew for it and, on one line, its width (pt)."""
    import oracle
    import quartz_pdf

    pdf = oracle.export(probe.build(), name="wrap-kern-probe")
    lines = quartz_pdf.lines(quartz_pdf.read(pdf))
    out = []
    cursor = 0
    for case in probe.cases():
        drawn: list[str] = []
        width = None
        while "".join(drawn) != case.word:
            if cursor >= len(lines) or len("".join(drawn)) > len(case.word):
                raise SystemExit(f"lost the paragraphs at {case}: read {drawn}")
            line = lines[cursor]
            cursor += 1
            text = "".join(run.text for run in line.runs).strip()
            drawn.append(text)
            mark = line.runs[-1]
            if mark.text == " ":
                width = (mark.x - LEFT_PX) / PX_PER_PT
        out.append({"face": case.face, "size": case.size, "word": case.word, "kern": case.kern, "tag": case.tag,
                    "twips": case.twips, "lines": drawn, "width": round(width, 5) if len(drawn) == 1 else None})
    if cursor != len(lines):
        raise SystemExit(f"{len(lines) - cursor} lines left over")
    return out


def model_width(units: dict, size: float, model: str) -> float:
    pairs = {"U": 0, "L": sum(units["legacy"]), "G": sum(units["gpos"])}[model]
    return (sum(units["advances"]) + pairs) * size / units["units_per_em"]


def score(observed: list[dict], units: dict) -> None:
    for kern in probe.KERNING:
        rows = [o for o in observed if o["kern"] == kern and o["tag"] != "width"]
        label = "w:kern 2" if kern else "no w:kern"
        for model in ("U", "L", "G"):
            right = 0
            band = [float("-inf"), float("inf")]  # fits when available - width >= band edge
            for o in rows:
                width = model_width(units[f"{o['face']}|{o['word']}"], o["size"], model)
                slack = o["twips"] / 20 - width
                fits = len(o["lines"]) == 1
                right += fits == (slack >= 0)
                if fits:
                    band[1] = min(band[1], slack)
                else:
                    band[0] = max(band[0], slack)
            print(f"{label:10} model {model}: {right}/{len(rows)} verdicts right; wraps at slack <= {band[0]:+.4f} pt,"
                  f" fits at slack >= {band[1]:+.4f} pt")
        worst = {model: 0.0 for model in "ULG"}
        for o in rows:
            if o["width"] is None:
                continue
            for model in worst:
                error = abs(o["width"] - model_width(units[f"{o['face']}|{o['word']}"], o["size"], model))
                worst[model] = max(worst[model], error)
        print(f"{label:10} widths of the one-line words, worst error: "
              + ", ".join(f"{m} {e:.4f} pt" for m, e in worst.items()))
    for o in observed:
        if o["tag"] == "width":
            found = units[f"{o['face']}|{o['word']}"]
            print(f"{o['face']:14} {o['word']:7} width {o['width']:.4f} pt: "
                  + ", ".join(f"{m} {model_width(found, o['size'], m):.4f}" for m in "ULG"))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args(argv[1:])
    observed = observe()
    units = {f"{face}|{word}": probe.units(face, word)
             for face, word in sorted({(o["face"], o["word"]) for o in observed})}
    score(observed, units)
    if args.record:
        import render_record

        render_record.dump(OBSERVATIONS, {
            "_about": ("Where Word 16.106 wrapped each paragraph of tools/make_wrap_kern_probe.py -- a word in a "
                       "text width set to its plain, legacy-kerned and GPOS-kerned width give or take a fraction "
                       "of a point, with w:kern and without -- the lines it drew, the width of each word it kept "
                       "on one line (pt), and the advances and kern pairs of each word in the faces Word lays out "
                       "with (font units). Measurements only; regenerate with tools/read_wrap_kern_probe.py "
                       "--record."),
            "cases": observed,
            "units": units,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
