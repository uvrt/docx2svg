#!/usr/bin/env python3
"""The glyph-position check: every glyph the SVG draws, against where Word drew it.

Phase 5's claim is that emission is a transcription of verified positions.  This makes
the claim an assertion, in two steps, each an equality:

1. **The SVG says what the layout says** (:func:`svg_glyphs` against
   :func:`layout_glyphs`): every ``<text>``'s per-glyph ``x`` and its ``y`` are the
   layout's exact positions, formatted.  Nothing in the writer moves a glyph.
2. **The layout says what Word drew** (:func:`compare`): the pen position of every
   glyph in Word's PDF (``quartz_pdf``, which reads the content stream Word's export
   writes in 1/300-inch device pixels), matched glyph for glyph, page by page.

Word's pen positions come in two precisions.  The first glyph of each text object
(Quartz starts a new object wherever Word positions a glyph on its own: a new run, a
tab, a script, the paragraph mark) is printed to seven significant digits -- 1e-4 px
below x = 1000, 1e-3 px above -- and is held to half of that, which is less than half
of Word's layout unit (0.00102 px): a one-unit error fails.  Inside an object Quartz
encodes each glyph's *advance*, not its position: integer ``/Widths`` per 1000 em of
the size it draws at, one ``Tc`` for the object and an integer ``TJ`` after a glyph,
so each advance is within half of 1/1000 em of Word's, and the error of a position
accumulates along the object (measured: up to 0.86 px after 100 glyphs of Calibri
11 pt, the same per glyph every time it occurs).  So inside an object the check is on
**advances**: from each glyph to the next non-space glyph of the same object, Word's
step and the model's agree to half of 1/1000 em per advance between them
(:data:`STEP_EM`).  That catches a missing kern pair or a wrong advance of any glyph;
and the pen after each object is held exactly by the next object's start.  Baselines (with a script's or
``w:position``'s offset) are whole device pixels and must be equal; so must the face
(from ``/BaseFont``) and the drawn size (Quartz's ``Tm`` scale is the size rounded to
whole px).

Word also draws what the layout does not: footnotes, and everything past the point where
the layout stops (a floating drawing).  Those glyphs are counted as *not drawn by the
model*, separately; a glyph the model draws that Word does not is *extra* and is a
failure.  Headers and footers are drawn (ROADMAP.md, "Headers, footers and fields"), and
their glyphs are held to the same checks and also counted apart (``story_*``).

Usage::

    python tools/glyphs.py DOCX...           # export with Word (cached), compare, summarise
    python tools/glyphs.py -v DOCX           # ... and list every disagreement
"""

from __future__ import annotations

import argparse
import difflib
import sys
from dataclasses import dataclass, field
from fractions import Fraction
from pathlib import Path
from xml.etree import ElementTree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

SVG = "{http://www.w3.org/2000/svg}"
#: How far Quartz's encoding of one advance may be from Word's, in ems of the size it
#: draws at: half of the ``TJ`` grid (1/1000 em) and half of the precision ``Tc`` is
#: printed to (four decimals of an em).
STEP_EM = 0.0005 + 0.00005


def exact_tolerance(x: float) -> float:
    """What Quartz's printing of a text object's pen x leaves uncertain: the position is a
    single-precision float in points (two units in its last place, for a sum and a
    product), scaled to device px, then printed to seven significant digits (half the
    last one).  At x = 989 px that is 0.00018 px -- a sixth of Word's layout unit, so an
    error of one unit still fails."""
    import math

    quantum = 10 ** (math.floor(math.log10(max(abs(x), 1))) - 6)
    points = max(abs(x) * 72 / 300, 1)
    ulp = 2.0 ** (math.floor(math.log2(points)) - 23) * 300 / 72
    return quantum / 2 + 2 * ulp + 1e-6


@dataclass(frozen=True)
class Glyph:
    page: int
    char: str
    x: float
    y: float
    face: str
    size_px: float
    bold: bool = False
    italic: bool = False
    #: The line's baseline, device px (for ordering; a script is drawn off it).
    line_y: float = 0.0
    #: Word's side: the first glyph of a text object, whose x is exact.
    exact: bool = True
    #: Word's side: which text object, and the glyph's index in it (spaces counted).
    run: int = -1
    index: int = 0
    #: The model's side: ``header`` or ``footer`` for a glyph of a story, else ``None``.
    story: str | None = None


# -- the SVG ----------------------------------------------------------------------------


def svg_glyphs(documents: list[str]) -> list[Glyph]:
    """Every non-space glyph an SVG page draws, in line order."""
    out = []
    for page, document in enumerate(documents):
        root = ElementTree.fromstring(document)
        for line in root.iter(f"{SVG}g"):
            baseline = line.get("data-docx-baseline")
            if baseline is None:
                continue
            for text in line.findall(f"{SVG}text"):
                xs = [float(v) for v in text.get("x").split()]
                chars = text.text or ""
                assert len(xs) == len(chars), (text.get("data-docx-path"), len(xs), len(chars))
                for char, x in zip(chars, xs):
                    if char.isspace():
                        continue
                    out.append(Glyph(page, char, x, float(text.get("y")), text.get("font-family"),
                                     float(text.get("font-size")), text.get("font-weight") == "bold",
                                     text.get("font-style") == "italic", float(baseline)))
    return out


def layout_glyphs(layout, glyph_size: str = "device") -> list[Glyph]:
    """Every non-space glyph of a :class:`docx2svg.layout.Layout`, exactly."""
    from docx2svg.svg import font_size_px

    out = []
    for page_index, page in enumerate(layout.pages):
        for line in page.text_lines():
            for span in line.spans:
                for char, x in zip(span.chars, span.xs):
                    if char.isspace():
                        continue
                    out.append(Glyph(page_index, char, x, span.y, span.face,
                                     float(font_size_px(span.half_points, glyph_size)), span.bold, span.italic,
                                     line.baseline, story=getattr(line, "story", None)))
    return out


# -- Word's PDF -------------------------------------------------------------------------


def word_glyphs(pdf: Path) -> list[Glyph]:
    """Every non-space glyph Word drew, with its pen x.  Both sides are compared in the
    order of the baseline each glyph is *drawn* on, then x -- not folded into lines, which
    by size and distance (``baselines.merge_raised``) would take a small line just below a
    large one for a raised run."""
    import baselines
    import quartz_pdf

    out = []
    number = 0
    for line in quartz_pdf.lines(quartz_pdf.read(pdf)):
        for run in line.runs:
            family, bold, italic = baselines._drawn_face(run.font)
            number += 1
            for k, (char, x) in enumerate(zip(run.text, run.xs)):
                if char.isspace():
                    continue
                out.append(Glyph(line.page, char, x, run.y, family, run.size_px, bold, italic, line.y, k == 0,
                                 number, k))
    return out


# -- comparing ----------------------------------------------------------------------------


@dataclass
class Score:
    name: str
    word: int = 0
    model: int = 0
    matched: int = 0
    #: Matched glyphs whose x agrees: at a text object's start (exact), or inside one by
    #: its advance from the glyph before it (:data:`STEP_EM`).
    x_agree: int = 0
    exact_scored: int = 0
    exact_agree: int = 0
    step_scored: int = 0
    step_agree: int = 0
    y_agree: int = 0
    face_agree: int = 0
    size_agree: int = 0
    #: Word's glyphs the model does not draw (headers, footers, notes, past a stop), and
    #: the model's that Word does not draw.
    not_drawn: int = 0
    extra: int = 0
    worst_dx: float = 0.0
    #: The largest inner drift seen, in ems of the drawn size.
    worst_inner_em: float = 0.0
    problems: list[str] = field(default_factory=list)
    #: Of the model's glyphs, those of headers and footers: drawn, matched, and matched
    #: with every check agreeing.
    story_model: int = 0
    story_matched: int = 0
    story_exact: int = 0

    @property
    def all_agree(self) -> int:
        return self.matched if not self.problems else self.matched - len({p.split("|")[0] for p in self.problems})

    def row(self) -> str:
        return (f"{self.name:34} glyphs {self.matched:7} matched of {self.model:7} drawn "
                f"(Word {self.word}); x {self.x_agree}/{self.matched} (object starts {self.exact_agree}/"
                f"{self.exact_scored}, advances {self.step_agree}/{self.step_scored}), baseline {self.y_agree}, face {self.face_agree}, size {self.size_agree}; "
                f"extra {self.extra}, not drawn {self.not_drawn}; worst |dx| at object starts {self.worst_dx:.4f} px, "
                f"advances {self.worst_inner_em * 1000:.3f}/1000 em")


def _key(glyph: Glyph) -> str:
    """What two glyphs must share to be matched: the character, except that a symbol
    font's private-use code (U+F0xx, how a list label names SymbolMT's bullet) and the
    Unicode Word's ``ToUnicode`` map gives it (``•``) are the same glyph."""
    if "\uf000" <= glyph.char <= "\uf0ff" or _norm_face(glyph.face) in ("symbol", "wingdings"):
        return "\uf8ff"
    return glyph.char


def _norm_face(face: str) -> str:
    return face.replace(" ", "").lower()


def compare(name: str, model: list[Glyph], word: list[Glyph], *, stop_pages: set[int] | None = None) -> Score:
    """Match ``model`` against ``word`` page by page (their characters, in line order) and
    score every matched glyph."""
    score = Score(name, word=len(word), model=len(model), story_model=sum(1 for g in model if g.story))
    pages = sorted({g.page for g in model} | {g.page for g in word})
    for page in pages:
        ours = sorted((g for g in model if g.page == page), key=lambda g: (round(g.y), g.x))
        theirs = sorted((g for g in word if g.page == page), key=lambda g: (round(g.y), g.x))
        matcher = difflib.SequenceMatcher(None, [_key(g) for g in ours], [_key(g) for g in theirs], autojunk=False)
        matched_here = 0
        previous = None
        for block in matcher.get_matching_blocks():
            for k in range(block.size):
                a, b = ours[block.a + k], theirs[block.b + k]
                matched_here += 1
                score.matched += 1
                dx = abs(a.x - b.x)
                if b.exact:
                    x_ok = dx <= exact_tolerance(b.x)
                    score.exact_scored += 1
                    score.exact_agree += x_ok
                    score.worst_dx = max(score.worst_dx, dx)
                elif previous is not None and previous[1].run == b.run:
                    advances = b.index - previous[1].index
                    error = abs((a.x - previous[0].x) - (b.x - previous[1].x))
                    # 2e-3 px: a recording keeps the steps to 1e-3 px.
                    x_ok = error <= advances * STEP_EM * b.size_px + 2e-3
                    score.step_scored += 1
                    score.step_agree += x_ok
                    score.worst_inner_em = max(score.worst_inner_em, error / b.size_px / advances)
                else:
                    # Inside an object, but the glyph before it was not matched: its
                    # position is not checkable to better than the object's drift.
                    x_ok = True
                score.x_agree += x_ok
                previous = (a, b)
                y_ok = round(a.y) == round(b.y)
                face_ok = (_norm_face(a.face), a.bold, a.italic) == (_norm_face(b.face), b.bold, b.italic)
                size_ok = round(a.size_px) == round(b.size_px)
                score.y_agree += y_ok
                score.face_agree += face_ok
                score.size_agree += size_ok
                if a.story:
                    score.story_matched += 1
                    score.story_exact += bool(x_ok and y_ok and face_ok and size_ok)
                if not (x_ok and y_ok and face_ok and size_ok):
                    score.problems.append(
                        f"p{page + 1} {a.char!r}|: model x={a.x:.4f} y={a.y:g} {a.face} {a.size_px:g}"
                        f"{' b' if a.bold else ''}{' i' if a.italic else ''}; Word x={b.x:.4f} y={b.y:g} "
                        f"{b.face} {b.size_px:g}{' b' if b.bold else ''}{' i' if b.italic else ''}"
                        f"{'' if b.exact else ' (in TJ)'} dx={a.x - b.x:+.4f}")
        score.extra += len(ours) - matched_here
        score.not_drawn += len(theirs) - matched_here
    return score


def score_document(path: Path, *, glyph_size: str = "device", options=None, pdf: Path | None = None):
    """Render ``path`` with the library, check the SVG against the layout, export it with
    Word (cached) and compare.  Returns ``(Score, layout)``."""
    import oracle

    from docx2svg import ConvertOptions, _render

    options = options or ConvertOptions(glyph_size=glyph_size)
    documents, layout, data, _fonts = _render(path, options)
    ours = svg_glyphs(documents)
    exact = layout_glyphs(layout, options.glyph_size)
    assert [(g.page, g.char) for g in ours] == [(g.page, g.char) for g in exact], "the SVG lost a glyph"
    for a, b in zip(ours, exact):
        # The writer formats to four decimals, half up: within 5e-5 px of the layout.
        assert abs(a.x - float(b.x)) <= 5e-5 + 1e-9 and a.y == float(b.y), (a, b)
    pdf = pdf or oracle.export(data, name=path.stem)
    score = compare(path.name, exact_floats(exact), word_glyphs(pdf))
    return score, layout


def exact_floats(glyphs: list[Glyph]) -> list[Glyph]:
    return [Glyph(g.page, g.char, float(Fraction(g.x)), float(g.y), g.face, g.size_px, g.bold, g.italic,
                  float(g.line_y), story=g.story) for g in glyphs]


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx", nargs="+", type=Path)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--limit", type=int, default=40)
    args = parser.parse_args(argv[1:])
    for path in args.docx:
        score, layout = score_document(path)
        print(score.row())
        if args.verbose:
            for problem in score.problems[:args.limit]:
                print("   ", problem.replace("|", ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
