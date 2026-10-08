#!/usr/bin/env python3
"""The visibility check: every glyph the model draws must show as ink in the raster.

The glyph-position check (``glyphs.py``) proves each glyph is *where* Word drew it; it
cannot see whether anything is drawn *over* it.  A cell's shading emitted after its text
paints over the text in SVG's painter's order, and every glyph still sits at Word's
position (ROADMAP.md, "Tables -- measured": the banded rows of ``sample-with-table``).
This check looks at the pixels instead.  Each page is rasterised three times:

* **A**, the page as emitted;
* **B**, the same page with the model's text removed, so every fill, border and
  picture is exactly where it is in A;
* **C**, the model's text alone, black on white, which says where this rasteriser can
  put ink for a glyph at all (a host without a face falls back to another; a glyph no
  loaded face has draws nothing, and is *not drawable here*, not a failure).

A glyph is **visible** when, inside its box (pen x to the next pen x, a size-scaled band
about its drawn baseline), A differs from B by at least :data:`INK_DIFFERENCE` in some
channel at :data:`MIN_INK` pixels or more.  A glyph under an opaque fill leaves A equal to
B there -- a flat fill -- and so does a glyph drawn in the colour it sits on; both fail.
The check needs no font file and no Word: the committed documents are laid out from the
recorded face numbers (``render_record.RecordedFonts``, as the VRT is), and the
rasteriser draws with whatever faces the host has.

Dev-only: numpy, pillow and resvg-py.

Usage::

    python tools/visibility.py DOCX...        # measured with the installed faces
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass, field
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

#: A channel difference, 0-255, that counts as ink rather than anti-aliasing noise.
INK_DIFFERENCE = 48
#: Pixels of ink a glyph's box must hold (at :data:`SCALE`) to be visible or drawable.
MIN_INK = 2
#: Raster pixels per device pixel.  Half of Word's 300 dpi grid: an 11 pt glyph is 23 px
#: tall, ample to find ink in, at a quarter of the cost.
SCALE = Fraction(1, 2)
#: The glyph box about the drawn baseline, in ems of the drawn size.
ABOVE_EM = Fraction(3, 4)
BELOW_EM = Fraction(1, 4)

_MODEL_TEXT = re.compile(r"^<text data-docx-path=.*</text>$", re.M)


@dataclass(frozen=True)
class Glyph:
    page: int
    char: str
    path: str
    #: The box in device px: left, top, right, bottom.
    box: tuple[float, float, float, float]
    #: Whether a floating drawing painted after it (in front of the text) covers its box:
    #: Word paints that drawing over it too, so it may legitimately show no ink.
    covered: bool = False


@dataclass
class Result:
    name: str
    glyphs: int = 0
    visible: int = 0
    #: Glyphs this rasteriser cannot draw with the faces it has (C holds no ink there).
    undrawable: int = 0
    hidden: list[Glyph] = field(default_factory=list)
    #: Glyphs showing no ink under a floating drawing Word paints over them as well.
    covered: int = 0

    def row(self) -> str:
        return (f"{self.name:34} glyphs {self.glyphs:6}, visible {self.visible:6}, hidden {len(self.hidden):5}, "
                f"covered by a drawing in front {self.covered}, not drawable here {self.undrawable}")


def _opaque(placed) -> bool:
    """Whether a floating drawing paints anything that can hide text (a fill, a picture)."""
    return any(p.kind == "image" or (p.kind == "path" and p.fill) or p.kind == "placeholder"
               for p in placed.primitives)


def layout_glyphs(layout) -> list[Glyph]:
    """Every non-space glyph of the layout with its box, in the SVG's order, and whether a
    floating drawing in front of the text, painted after it, covers it."""
    from docx2svg.svg import font_size_px

    out = []
    for index, page in enumerate(layout.pages):
        steps = page.paint()
        # Each line with the opaque drawings painted after it.
        groups = []
        for k, step in enumerate(steps):
            if step[0] not in ("lines", "floats"):
                continue
            later = [f for later_step in steps[k + 1:] if later_step[0] == "floats" for f in later_step[1]]
            if step[0] == "lines":
                groups += [(line, later) for line in step[1]]
            else:
                for m, placed in enumerate(step[1]):
                    groups += [(line, step[1][m + 1:] + later) for line in placed.lines]
        for line, later in groups:
            covers = [(float(f.x), float(f.y), float(f.x + f.width), float(f.y + f.height)) for f in later
                      if _opaque(f)]
            for span in line.spans:
                if not span.chars:
                    continue
                size = font_size_px(span.half_points, "device")
                ends = list(span.xs[1:]) + [span.end]
                for char, x, end in zip(span.chars, span.xs, ends):
                    if char.isspace():
                        continue
                    box = (float(x), float(span.y - ABOVE_EM * size), float(max(end, x + 1)),
                           float(span.y + BELOW_EM * size))
                    covered = any(box[0] < c[2] and c[0] < box[2] and box[1] < c[3] and c[1] < box[3]
                                  for c in covers)
                    out.append(Glyph(index, char, span.path, box, covered))
    return out


def without_model_text(document: str) -> str:
    """B: the page with every glyph the model draws removed, and nothing else."""
    return _MODEL_TEXT.sub("", document)


def model_text_only(document: str) -> str:
    """C: the model's glyphs alone, black, on white."""
    header = document.split("\n", 1)[0]
    texts = [re.sub(r' fill="[^"]*"', ' fill="#000000"', text) for text in _MODEL_TEXT.findall(document)]
    return "\n".join([header, *texts, "</svg>"]) + "\n"


def _raster(document: str, width: int, height: int):
    import numpy as np
    from PIL import Image

    from docx2svg.png import svg_to_png

    # Sized in raster pixels: resvg does not take a size in points.
    document = re.sub(r' width="[^"]*" height="[^"]*"', f' width="{width}" height="{height}"', document, count=1)
    png = svg_to_png(document, backend="resvg", background="white")
    return np.asarray(Image.open(io.BytesIO(png)).convert("RGB")).astype(np.int16)


def check_pages(name: str, documents: list[str], layout) -> Result:
    """Rasterise each page as A, B and C, and score every glyph of ``layout``."""
    import numpy as np

    result = Result(name)
    glyphs = layout_glyphs(layout)
    for index, document in enumerate(documents):
        page = layout.pages[index]
        width, height = int(page.width_px * SCALE), int(page.height_px * SCALE)
        a = _raster(document, width, height)
        b = _raster(without_model_text(document), width, height)
        c = _raster(model_text_only(document), width, height)
        shown = np.abs(a - b).max(axis=2) >= INK_DIFFERENCE
        drawable = (255 - c).max(axis=2) >= INK_DIFFERENCE
        for glyph in (g for g in glyphs if g.page == index):
            x0, y0, x1, y1 = (v * float(SCALE) for v in glyph.box)
            rows = slice(max(0, int(y0)), max(0, int(y1 + 0.999)))
            cols = slice(max(0, int(x0)), max(0, int(x1 + 0.999)))
            result.glyphs += 1
            if int(drawable[rows, cols].sum()) < MIN_INK:
                result.undrawable += 1
            elif int(shown[rows, cols].sum()) >= MIN_INK:
                result.visible += 1
            elif glyph.covered:
                result.covered += 1
            else:
                result.hidden.append(glyph)
    return result


def check(name: str, source, options=None) -> Result:
    """Lay ``source`` out (with ``options``: the installed faces by default), render and
    check every page."""
    from docx2svg import ConvertOptions, _render

    documents, layout, _data, _fonts = _render(source, options or ConvertOptions())
    return check_pages(name, documents, layout)


def describe(hidden: list[Glyph], limit: int = 12) -> str:
    """The hidden glyphs, grouped by page and run: which text is not seen."""
    groups: dict[tuple[int, str], str] = {}
    for glyph in hidden:
        groups[(glyph.page, glyph.path)] = groups.get((glyph.page, glyph.path), "") + glyph.char
    lines = [f"p{page + 1} {path}: {text!r}" for (page, path), text in list(groups.items())[:limit]]
    if len(groups) > limit:
        lines.append(f"... and {len(groups) - limit} more run(s)")
    return "\n".join(lines)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx", nargs="+", type=Path)
    args = parser.parse_args(argv[1:])
    failed = 0
    for path in args.docx:
        result = check(path.name, path)
        print(result.row())
        if result.hidden:
            failed += 1
            print("   " + describe(result.hidden).replace("\n", "\n   "))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
