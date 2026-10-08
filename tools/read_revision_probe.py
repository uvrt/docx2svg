#!/usr/bin/env python3
"""Score tracked changes, content controls and comments against Word's final view:
``make_revision_probe.py``.

Word exports each document twice: in its *No Markup* view (``oracle.export``'s default,
``final``), which is what the renderer is scored against, and with every revision
accepted first (``accept``), which says whether the final view is what accepting makes.
Word's PDF gives every text object (``read_render.word_objects``) and filled rectangle;
the model lays each document out and draws it.  Scores, per document:

* **cases** -- a case agrees when every line of its pages, from ``Case N`` to the line
  before ``Case N+1``, is Word's: the same glyphs, on the same page (counted from the
  case's first), at the same baseline below ``Case N`` and the same left edge, to the
  device pixel (:func:`outcomes`); listed per family;
* **lines** -- ``read_wrap_side_probe.line_score``: every glyph of every line of Word's
  matched on its baseline, each text object's start within half a pixel;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **rules** -- ``read_render.rect_score``'s totals: the device pixels of the model's
  rules (table borders, shading) against Word's fills.

``--record`` writes Word's side to ``tests/fixtures/revision-observations.json``;
``tests/test_revisions.py`` holds the model to it offline.

Usage::

    python tools/read_revision_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_revision_probe as probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "revision-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(f"revision-{setting}", probe.build(setting)) for setting in probe.SETTINGS]


def _lines(glyph_list) -> list[tuple[int, float, float, str]]:
    """``(page, baseline, left x, characters)`` of every line, in page and baseline order:
    the glyphs sharing a page and a baseline (to the device pixel), spaces not counted."""
    rows: dict = {}
    for g in glyph_list:
        rows.setdefault((g.page, round(float(g.y))), []).append(g)
    out = []
    for (page, _), found in sorted(rows.items()):
        found.sort(key=lambda g: float(g.x))
        out.append((page, float(found[0].y), float(found[0].x), "".join(g.char for g in found)))
    return out


def outcomes(glyph_list) -> dict:
    """``{case key: [[page - case page, baseline - case baseline (px), left x (px),
    characters], ...]}`` for every line from ``Case N`` to the line before ``Case N+1``;
    ``None`` for a case whose label is not drawn."""
    lines = _lines(glyph_list)
    starts = {text: k for k, (_, _, _, text) in enumerate(lines) if text.startswith("Case")}
    out = {}
    for number, case in enumerate(probe.CASES):
        start = starts.get(f"Case{number}")
        if start is None:
            out[case.key] = None
            continue
        end = starts.get(f"Case{number + 1}", len(lines))
        page0, y0 = lines[start][0], lines[start][1]
        out[case.key] = [[page - page0, round(y - y0), round(x), text] for page, y, x, text in lines[start:end]]
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
    word_cases = outcomes(word)
    model_cases = outcomes(model)
    rules = read_render.rect_score(layout, recorded["fills"])["all"]
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "families": families(word_cases, model_cases), "word": word_cases, "model": model_cases,
            "disagree": sorted(k for k in word_cases if word_cases[k] != model_cases[k]), "rules": rules,
            "warnings": sorted({w.code for w in options.warnings}), "pages": len(layout.pages), "result": result}


def accept_differs(final_objects: list, accept_objects: list) -> list[str]:
    """The cases whose lines Word draws otherwise once every revision is accepted."""
    final = outcomes(read_render.glyphs_of(final_objects))
    accepted = outcomes(read_render.glyphs_of(accept_objects))
    return sorted(k for k in final if final[k] != accepted[k])


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
            accepted = oracle.export(data, name=name, view="accept")
            word = {"objects": read_render.word_objects(pdf), "fills": read_render.word_fills(pdf)}
            accept_objects = read_render.word_objects(accepted)
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            differs = accept_differs(word["objects"], accept_objects)
            print(result["result"].row())
            print(f"{'':34} cases {sum(a for a, _ in result['families'].values())} / {len(probe.CASES)} "
                  f"({', '.join(f'{k} {a}/{b}' for k, (a, b) in result['families'].items())}); lines "
                  f"{result['lines'][0]} / {result['lines'][1]}; rules {result['rules']}; pages {result['pages']}; "
                  f"warnings {result['warnings']}; accepting differs in {differs}")
            if args.verbose:
                for case in probe.CASES:
                    k = case.key
                    if result["word"][k] == result["model"][k]:
                        continue
                    print(f"    XX {k}")
                    print(f"       Word  {result['word'][k]}")
                    print(f"       model {result['model'][k]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "word": result["word"],
                              "disagree": result["disagree"], "rules": result["rules"],
                              "warnings": result["warnings"], "pages": result["pages"],
                              "accept_differs": differs}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_revision_probe.py in its No Markup view: every text "
                       "object and filled rectangle (as render-observations.json records them), each case's "
                       "lines as Word drew them (page and baseline from the case's label, left edge, glyphs), "
                       "the cases whose lines differ once every revision is accepted, the renderer's scores "
                       "against them -- cases per family, lines within half a pixel, the glyph row, the rule "
                       "pixels -- and every face number the renderer asked for. Measurements only; regenerate "
                       "with tools/read_revision_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
