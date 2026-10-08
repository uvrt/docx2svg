#!/usr/bin/env python3
"""Score tab leaders against Word: ``make_tab_leader_probe.py``.

Word's PDF gives every text object (``read_render.word_objects``) -- a leader is drawn as
glyphs -- and every filled rectangle (``read_render.word_fills``); the model lays each
document out and draws it.  Scores, per document:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it is within
  half a device pixel of the model's (``read_wrap_side_probe.line_score``), per family
  (a family starts its page with ``Family <name>``);
* **leaders** -- Word's leader glyphs (:data:`LEADER_CHARS`) the model draws at their
  position, of all Word draws;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **rules** -- the model's rules against Word's fills (``model only`` / ``Word only``
  device pixels, ``read_endnote_probe.rect_score``);
* **pages** -- the model's page count and Word's.

``--record`` writes Word's side to ``tests/fixtures/tab-leader-observations.json``;
``tests/test_tab_leaders.py`` holds the model to it offline.

Usage::

    python tools/read_tab_leader_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_tab_leader_probe as probe  # noqa: E402
import read_endnote_probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "tab-leader-observations.json"
#: The glyphs a leader is drawn with.
LEADER_CHARS = ".-_·․‧∙"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name)) for name in probe.DOCUMENTS]


def page_families(word: list) -> dict[int, str]:
    """Page -> the family on it: the last ``Family <name>`` anchor at or before it."""
    out: dict[int, str] = {}
    current = "none"
    for page in sorted({g.page for g in word}):
        top = sorted((g for g in word if g.page == page), key=lambda g: (round(g.y), g.x))
        first_line = [g for g in top if round(g.y) == round(top[0].y)] if top else []
        text = "".join(g.char for g in first_line)
        if text.startswith("Family"):
            current = text[len("Family"):]
        out[page] = current
    return out


def leader_score(model: list, word: list, lines: dict) -> list[int]:
    """``[leader glyphs on lines that agree, leader glyphs Word draws]``."""
    theirs = [g for g in word if g.char in LEADER_CHARS]
    agreeing = sum(1 for g in theirs if not lines.get((g.page, round(g.line_y or g.y)), [None]))
    return [agreeing, len(theirs)]


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    """Every score of one document against Word's recorded side."""
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = read_wrap_side_probe.line_score(model, word)
    families: dict = {}
    of_page = page_families(word)
    for (page, _y), problems in lines.items():
        row = families.setdefault(of_page.get(page, "none"), [0, 0])
        row[0] += not problems
        row[1] += 1
    rects = read_endnote_probe.rect_score(layout, recorded["fills"])
    word_pages = 1 + max((o[0] for o in recorded["objects"]), default=0)
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "families": dict(sorted(families.items())), "leaders": leader_score(model, word, lines),
            "rules": list(rects["all"][2:]), "pages": [len(layout.pages), word_pages],
            "warnings": sorted({w.code for w in options.warnings}),
            "problems": {f"p{page + 1} y={y}": problems for (page, y), problems in lines.items() if problems},
            "rect_pages": {key: row[2:] for key, row in rects.items() if key != "all" and any(row[2:])},
            "result": result}


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
            word = {"objects": read_render.word_objects(pdf), "fills": read_render.word_fills(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]}; leader glyphs "
                  f"{result['leaders'][0]} / {result['leaders'][1]}; rule px model only {result['rules'][0]}, "
                  f"Word only {result['rules'][1]}; pages {result['pages'][0]} (Word {result['pages'][1]}); "
                  f"warnings {result['warnings']}")
            print(f"{'':34} {result['families']}")
            if args.verbose:
                for where, problems in list(result["problems"].items())[:args.limit]:
                    print(f"    {where}: {problems[:3]}")
                for where, row in list(result["rect_pages"].items())[:args.limit]:
                    print(f"    rules {where}: model only {row[0]}, Word only {row[1]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "leaders": result["leaders"],
                              "rules": result["rules"], "pages": result["pages"], "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_tab_leader_probe.py: every text object (as "
                       "render-observations.json records them) and every filled rectangle, the renderer's scores "
                       "against them -- lines within half a pixel, leader glyphs, the glyph row, rule pixels, "
                       "pages -- and every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_tab_leader_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
