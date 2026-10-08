#!/usr/bin/env python3
"""Does Word kern, and how does tracking combine with it?  Reads the kerning probe.

Each line's advance is read exactly: the paragraph mark is its own Quartz text object, so
its pen x is the line's width to 1e-4 px.  Predictions come from the face's legacy
``kern`` table and ``hmtx`` via fontTools, read from the font files *on this machine*
(Word's bundle, or the macOS system copy for Times New Roman -- see ROADMAP.md for why
that one).  Nothing is copied or committed; without fontTools or the files the reader
prints the measured widths only.

Measured on Word 16.106 (ROADMAP.md, Phase 2):

* **No ``w:kern``, no kerning.**  Every width equals the plain sum of advances to 1e-4 px.
* **With ``w:kern``, the ``kern`` table's pairs, exactly** -- and across run boundaries:
  one run per glyph (alternating colours) kerns identically to one run.
* **Tracking is per glyph, including the last, in 1/4096 pt**: 7 twips comes back as
  1434 units (1.45874 px), not 1433.6.
* **Kerning and tracking add.**  Neither is rounded to the device grid, so "before or
  after the rounding" has no content horizontally: nothing horizontal rounds but the
  1/4096 pt layout unit.

Usage::

    python tools/read_kerning_probe.py
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_kerning_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg.vertical import twips_to_px  # noqa: E402

_WORD_FONTS = Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts")
FONT_FILES = {
    "Calibri": _WORD_FONTS / "Calibri.ttf",
    "Arial": _WORD_FONTS / "arial.ttf",
    "Times New Roman": Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf"),
}


def _tables(path: Path):
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return None
    if not path.exists():
        return None
    font = TTFont(str(path))
    kern = {}
    if "kern" in font:
        for table in font["kern"].kernTables:
            kern.update(table.kernTable)
    return font.getBestCmap(), font["hmtx"], font["head"].unitsPerEm, kern


def main() -> int:
    pdf = oracle.export(probe.build(), name="kerning-probe")
    lines = quartz_pdf.lines(quartz_pdf.read(pdf))
    items = probe.items()
    if len(lines) != len(items):
        raise SystemExit(f"{len(lines)} lines for {len(items)} probe paragraphs")
    tables = {face: _tables(path) for face, path in FONT_FILES.items()}

    worst = 0.0
    for item, line in zip(items, lines):
        width = line.runs[-1].x - 300
        t = tables[item.face]
        if t is None:
            print(f"{item.face:16} {item.half_points} {item.text} {item.variant:12} width={width:.4f}")
            continue
        cmap, hmtx, upm, kern = t
        size_px = item.half_points / 2 * 300 / 72
        plain = sum(hmtx[cmap[ord(c)]][0] for c in item.text) * size_px / upm
        pairs = sum(kern.get((cmap[ord(a)], cmap[ord(b)]), 0) for a, b in zip(item.text, item.text[1:]))
        predicted = plain
        if item.kern is not None:
            predicted += pairs * size_px / upm
        if item.track is not None:
            predicted += float(twips_to_px(item.track)) * len(item.text)
        worst = max(worst, abs(width - predicted))
        print(
            f"{item.face:16} {item.half_points} {item.text} {item.variant:12}"
            f" {'per-glyph runs' if item.per_glyph else 'one run':14} width={width:9.4f}"
            f" predicted={predicted:9.4f} residual={width - predicted:+.4f}"
        )
    print(f"worst residual: {worst:.4f} px over {len(items)} lines")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
