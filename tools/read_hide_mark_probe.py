#!/usr/bin/env python3
"""Score empty table rows against Word: ``make_hide_mark_probe.py``.

Word's PDF gives every text object (``read_render.word_objects``); the model lays each
document out and draws it.  Scores, per document:

* **cases** -- a case agrees when the baselines of its ``N.bot`` row (or line) and of
  ``After N`` stand as far below ``N.top`` (``Case N`` where there is no top row) as
  Word's, to the device pixel; listed per family;
* **lines** -- ``read_wrap_side_probe.line_score``: every glyph of every line of Word's
  matched on its baseline, each text object's start within half a pixel;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it.

``--record`` writes Word's side to ``tests/fixtures/hide-mark-observations.json``;
``tests/test_hide_mark.py`` holds the model to it offline.

Usage::

    python tools/read_hide_mark_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_hide_mark_probe as probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "hide-mark-observations.json"
LABEL = re.compile(r"^(?:Case (\d+)$|After (\d+)$|(\d+)\.(top|bot|mid) [ab]$)")


def documents() -> list[tuple[str, bytes]]:
    return [(f"hide-mark-{setting}", probe.build(setting)) for setting in probe.SETTINGS]


def _where(lines) -> dict:
    """``{case: {role: (page, baseline)}}`` from ``(page, baseline, text)`` lines."""
    out: dict = {}
    for page, y, text in lines:
        match = LABEL.match(text.strip())
        if not match:
            continue
        if match[1] is not None:
            number, role = int(match[1]), "case"
        elif match[2] is not None:
            number, role = int(match[2]), "after"
        else:
            number, role = int(match[3]), match[4]
        out.setdefault(number, {}).setdefault(role, (page, round(y)))
    return out


def outcomes(lines) -> dict:
    """``{case key: [bot - reference, after - reference]}``, px, the reference being the
    top row's baseline (``Case N``'s where there is none); ``None`` where one is missing or
    on another page."""
    found = _where(lines)
    out = {}
    for number, case in enumerate(probe.CASES):
        seen = found.get(number, {})
        reference = seen.get("top") or seen.get("case")
        row = []
        for role in ("bot", "after"):
            at = seen.get(role)
            row.append(None if reference is None or at is None or at[0] != reference[0] else at[1] - reference[1])
        out[case.key] = row
    return out


def word_lines(objects: list) -> list:
    """``(page, baseline, text)`` of every text object of Word's (a cell's text is one)."""
    return [(o[0], o[1], o[4]) for o in objects]


def model_lines(layout) -> list:
    out = []
    for page in layout.pages:
        for line in page.lines:
            out.append((page.number, float(line.baseline), "".join("".join(s.chars) for s in line.spans)))
    return out


def families(word: dict, model: dict) -> dict:
    """``{family: [cases agreeing, cases]}``."""
    out: dict = {}
    for case in probe.CASES:
        row = out.setdefault(case.family, [0, 0])
        row[0] += model[case.key] == word[case.key]
        row[1] += 1
    return out


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = read_wrap_side_probe.line_score(model, word)
    word_cases = outcomes(word_lines(recorded["objects"]))
    model_cases = outcomes(model_lines(layout))
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "families": families(word_cases, model_cases), "word": word_cases, "model": model_cases,
            "disagree": sorted(k for k in word_cases if word_cases[k] != model_cases[k]),
            "warnings": sorted({w.code for w in options.warnings}), "pages": len(layout.pages), "result": result}


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    try:
        for name, data in documents():
            pdf = oracle.export(data, name=name)
            word = {"objects": read_render.word_objects(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} cases {sum(a for a, _ in result['families'].values())} / {len(probe.CASES)} "
                  f"({', '.join(f'{k} {a}/{b}' for k, (a, b) in result['families'].items())}); lines "
                  f"{result['lines'][0]} / {result['lines'][1]}; pages {result['pages']}; warnings "
                  f"{result['warnings']}")
            if args.verbose:
                for key in probe.CASES:
                    k = key.key
                    flag = "  " if result["word"][k] == result["model"][k] else "XX"
                    print(f"    {flag} {k:48} Word {result['word'][k]}  model {result['model'][k]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "word": result["word"],
                              "disagree": result["disagree"], "warnings": result["warnings"],
                              "pages": result["pages"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_hide_mark_probe.py: every text object (as "
                       "render-observations.json records them), each case's row distances as Word drew them "
                       "(the bottom row's and the After line's baseline below the top row's, px), the "
                       "renderer's scores against them -- cases per family, lines within half a pixel, the "
                       "glyph row -- and every face number the renderer asked for. Measurements only; "
                       "regenerate with tools/read_hide_mark_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
