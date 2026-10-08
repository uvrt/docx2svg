#!/usr/bin/env python3
"""Score list numbering against Word: ``make_numbering_probe.py``.

Word's PDF gives every text object (``read_render.word_objects``); the model lays each
document out and draws it.  Scores, per document:

* **cases** -- a case agrees when every line of its pages is Word's: the same glyphs (the
  label's included), at the same baseline below ``Case N`` and the same left edge, to the
  device pixel (:func:`outcomes`).  A case's lines are every line on the pages from its
  label's to the next case's, so a header's line counts with the case on whose page it is;
  listed per family;
* **lines** -- ``read_wrap_side_probe.line_score``: every glyph of every line of Word's
  matched on its baseline, each text object's start within half a pixel;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it.

``--record`` writes Word's side to ``tests/fixtures/numbering-observations.json``;
``tests/test_numbering.py`` holds the model to it offline.

Usage::

    python tools/read_numbering_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_numbering_probe as probe  # noqa: E402
import read_render  # noqa: E402
import read_revision_probe  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "numbering-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(f"numbering-{setting}", probe.build(setting)) for setting in probe.SETTINGS]


def outcomes(glyph_list) -> dict:
    """``{case key: [[page - case page, baseline - case baseline (px), left x (px),
    characters], ...]}`` for every line on the pages from ``Case N``'s to ``Case N+1``'s
    (that page excluded); ``None`` for a case whose label is not drawn."""
    lines = read_revision_probe._lines(glyph_list)
    starts = {text: (page, y) for page, y, _, text in lines if text.startswith("Case")}
    out = {}
    for number, case in enumerate(probe.CASES):
        start = starts.get(f"Case{number}")
        if start is None:
            out[case.key] = None
            continue
        page0, y0 = start
        following = starts.get(f"Case{number + 1}")
        end = following[0] if following else 10 ** 6
        out[case.key] = [[page - page0, round(y - y0), round(x), text]
                         for page, y, x, text in lines if page0 <= page < end]
    return out


def labels(case_lines) -> list[str]:
    """The case's list lines as Word drew them, ``label+text`` (``7.B0``), in order."""
    return [text for _, _, _, text in case_lines or [] if not text.startswith("Case")]


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
    word_cases = outcomes(word)
    model_cases = outcomes(model)
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
                  f"{result['lines'][0]} / {result['lines'][1]}; pages {result['pages']}; "
                  f"warnings {result['warnings']}")
            for case in probe.CASES:
                k = case.key
                same = result["word"][k] == result["model"][k]
                if args.verbose or not same:
                    print(f"    {'  ' if same else 'XX'} {k:44} Word {' '.join(labels(result['word'][k]))}")
                    if not same:
                        print(f"       {'':44} model {' '.join(labels(result['model'][k]))}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "word": result["word"],
                              "disagree": result["disagree"], "warnings": result["warnings"],
                              "pages": result["pages"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_numbering_probe.py: every text object (as "
                       "render-observations.json records them), each case's lines as Word drew them (page and "
                       "baseline from the case's label, left edge, glyphs: the list labels among them), the "
                       "renderer's scores against them -- cases per family, lines within half a pixel, the glyph "
                       "row -- and every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_numbering_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
