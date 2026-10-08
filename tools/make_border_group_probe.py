#!/usr/bin/env python3
"""Consecutive paragraphs with borders: which of their borders take room.

Word draws a run of consecutive paragraphs whose borders are the same as one box: the
bottom border under the last only, the top border over the first only, and between them
the ``w:between`` border if there is one.  The pagination probe met it (two paragraphs
with the same bottom border drew the second a border's height higher than the stack put
it); nothing had measured it, nor top borders at all.  Each case is a page: a plain
anchor with ``w:pageBreakBefore``, the paragraphs under test, and a plain paragraph after
them.  Calibri 11 pt, one line each, no spacing unless said; scored through
``baselines.predict``.

Cases (``CASES``), each with two border sets -- (``w:sz`` 4, ``w:space`` 1) and (12, 4):

* ``same`` -- two paragraphs with the same bottom border; ``three`` -- three;
* ``width`` / ``space`` / ``colour`` / ``style`` -- two whose bottom borders differ in
  that only;
* ``indent`` -- the same bottom border, the second indented 720 twips;
* ``spacing`` -- the same border, the first with 120 twips after;
* ``between`` -- two with the same bottom and ``w:between`` borders;
* ``top`` -- one paragraph with a top border; ``top pair`` -- two with the same top
  border; ``box pair`` -- two with the same top and bottom borders;
* ``lone`` -- one with a bottom border (the line-box probe's case, as a control).

One document per compatibility setting: no ``settings.xml``, and modes 12, 14 and 15.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
BORDERS = ((4, 1), (12, 4))


def _side(name: str, sz: int, space: int, color: str = "000000", val: str = "single") -> str:
    return f'<w:{name} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


def _cases() -> list[tuple[str, list[dict]]]:
    """``(name, [paragraph properties])`` per case."""
    out = []
    for sz, sp in BORDERS:
        b = _side("bottom", sz, sp)
        tag = f"{sz}/{sp}"
        out += [
            (f"lone {tag}", [{"pBdr": b}]),
            (f"same {tag}", [{"pBdr": b}, {"pBdr": b}]),
            (f"three {tag}", [{"pBdr": b}, {"pBdr": b}, {"pBdr": b}]),
            (f"width {tag}", [{"pBdr": b}, {"pBdr": _side("bottom", sz + 4, sp)}]),
            (f"space {tag}", [{"pBdr": b}, {"pBdr": _side("bottom", sz, sp + 2)}]),
            (f"colour {tag}", [{"pBdr": b}, {"pBdr": _side("bottom", sz, sp, "FF0000")}]),
            (f"style {tag}", [{"pBdr": b}, {"pBdr": _side("bottom", sz, sp, val="double")}]),
            (f"indent {tag}", [{"pBdr": b}, {"pBdr": b, "ind": {"left": 720}}]),
            (f"spacing {tag}", [{"pBdr": b, "spacing": {"before": 0, "after": 120, "line": 240, "lineRule": "auto"}},
                                {"pBdr": b}]),
            (f"between {tag}", [{"pBdr": b + _side("between", sz, sp)}, {"pBdr": b + _side("between", sz, sp)}]),
            (f"top {tag}", [{"pBdr": _side("top", sz, sp)}]),
            (f"top pair {tag}", [{"pBdr": _side("top", sz, sp)}, {"pBdr": _side("top", sz, sp)}]),
            (f"box pair {tag}", [{"pBdr": _side("top", sz, sp) + b}, {"pBdr": _side("top", sz, sp) + b}]),
        ]
    return out


CASES = _cases()


def blocks() -> list[tuple[str, str]]:
    """``(kind, w:p)``: ``anchor``, ``test`` or ``after``."""
    out = []
    for number, (name, paragraphs) in enumerate(CASES):
        out.append(("anchor", _p(f"Case {number} anchor", pageBreakBefore=True)))
        for index, props in enumerate(paragraphs):
            out.append(("test", _p(f"Case {number} p{index}", **props)))
        out.append(("after", _p(f"Case {number} after")))
    return out


def kinds() -> list[str]:
    return [kind for kind, _ in blocks()]


def _p(text: str, **props) -> str:
    props.setdefault("spacing", {"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
    return wml.paragraph(wml.run(text), mark={}, **props)


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=extra)
