#!/usr/bin/env python3
"""Measure ``layout-sweep.docx`` against Word's own PDF export of it.

This is the loop, and at the time of writing it is the whole of the loop::

    tools/make_layout_sweep.py  ->  .docx
    tools/word_export_pdf.applescript  ->  .pdf      (Word decides the layout)
    src/docx2svg/parse  ->  what the file asked for
    this script  ->  the residual between the two

Everything measured here is a *difference* between two glyph ink boxes, for the reason
set out at length in ``make_layout_sweep.py``: a difference cancels the glyph's left side
bearing, the face's vertical metrics, and Word's rounding of the page box, none of which
this project may ship a font file to model.  What survives is a number authored in twips
in our own fixture that has to come back out of Word's PDF in points.

Needs the ``oracle`` extra (``pypdfium2``).  Word's PDF carries text as vector objects,
so every number below is read from a text object rather than from pixels -- see the
header of ``word_export_pdf.applescript`` for why that makes this a better oracle than
the sibling project's.

Usage::

    python tools/read_layout_sweep.py [--pdf PATH] [--docx PATH] [--export]

``--export`` runs Word first.  Without it the PDF is expected to be there already.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from docx2svg.parse import parse_package  # noqa: E402
from docx2svg.units import device_page_extent_pt, twips_to_pt  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from oracle import ORACLE_DIR, SCRIPT, run  # noqa: E402
FIXTURE = REPO / "tests" / "fixtures" / "layout-sweep.docx"

#: Characters whose ink box bottom is the baseline cluster together within this many
#: points.  A descender sits about 3 pt below the baseline at 11 pt and the line advance
#: is about 13.5 pt, so any threshold in between works; 6 pt is the midpoint and the
#: clustering is stable over the whole gap rather than tuned to it.
_LINE_TOLERANCE_PT = 6.0


@dataclass(frozen=True)
class Glyph:
    char: str
    left: float
    bottom: float
    right: float
    top: float


@dataclass(frozen=True)
class Line:
    page: int
    glyphs: tuple[Glyph, ...]

    @property
    def text(self) -> str:
        return "".join(glyph.char for glyph in self.glyphs)

    @property
    def left(self) -> float:
        return min(glyph.left for glyph in self.glyphs)

    @property
    def right(self) -> float:
        return max(glyph.right for glyph in self.glyphs)

    @property
    def baseline(self) -> float:
        """The modal ink-box bottom.

        Not the minimum: that is the deepest descender.  Not the mean: that is dragged
        down by how many descenders the line happens to have, which would make the
        line advance depend on the *text* rather than on the layout.  The bottom shared
        by the most glyphs is the baseline, because every non-descending glyph sits on it.
        """
        counts: dict[float, int] = {}
        for glyph in self.glyphs:
            key = round(glyph.bottom, 2)
            counts[key] = counts.get(key, 0) + 1
        return max(counts.items(), key=lambda item: (item[1], -item[0]))[0]


def read_pages(pdf_path: Path) -> tuple[list[tuple[float, float]], list[Line]]:
    try:
        import pypdfium2 as pdfium
    except ImportError:  # pragma: no cover - environment-dependent
        raise SystemExit(
            "pypdfium2 is needed to read the oracle's PDF: pip install -e '.[oracle]'"
        )

    pdf = pdfium.PdfDocument(str(pdf_path))
    page_sizes = [tuple(page.get_size()) for page in pdf]

    lines: list[Line] = []
    for page_index, page in enumerate(pdf):
        text_page = page.get_textpage()
        glyphs: list[Glyph] = []
        for index in range(text_page.count_chars()):
            char = text_page.get_text_range(index, 1)
            if char in ("\r", "\n"):
                continue
            left, bottom, right, top = text_page.get_charbox(index)
            if right <= left:  # a zero-width mark; carries no position information
                continue
            glyphs.append(Glyph(char, left, bottom, right, top))

        # Cluster into lines by ink-box bottom.  Sorting descending puts the top line
        # of the page first, which is document order.
        glyphs.sort(key=lambda glyph: (-glyph.bottom, glyph.left))
        current: list[Glyph] = []
        for glyph in glyphs:
            if current and abs(current[0].bottom - glyph.bottom) > _LINE_TOLERANCE_PT:
                lines.append(Line(page_index, tuple(sorted(current, key=lambda g: g.left))))
                current = []
            current.append(glyph)
        if current:
            lines.append(Line(page_index, tuple(sorted(current, key=lambda g: g.left))))

    return page_sizes, lines


def base_fonts(pdf_path: Path) -> list[str]:
    """The ``/BaseFont`` names in the PDF.

    A substitution Word performed is visible here and nowhere else.  ``hello.docx``,
    which names Calibri and carries no ``w:docDefaults``, came back with
    ``AAAAAC+Calibri``, ``AAAAAE+Aptos`` *and* ``AAAAAG+Calibri`` -- Word resolved the
    theme's minor font for the runs that named no face.  Any face other than the one the
    fixture asks for means the measurements below are partly a measurement of font
    resolution, which is the sibling project's most expensive lesson.
    """
    raw = pdf_path.read_bytes()
    names = {
        match.group(1).decode("latin-1")
        for match in re.finditer(rb"/BaseFont\s*/([A-Za-z0-9+#-]+)", raw)
    }
    return sorted(names)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", type=Path, default=FIXTURE)
    parser.add_argument("--pdf", type=Path, default=ORACLE_DIR / "layout-sweep.pdf")
    parser.add_argument("--export", action="store_true", help="run Word first")
    arguments = parser.parse_args(argv[1:])

    if arguments.export:
        # The approved directory, the lock-file sweep the AppleScript's header explains,
        # and the machine-wide Word lock (tools/oracle.py's ``run``).  All are part of
        # "running the oracle", not incidental.
        arguments.pdf.unlink(missing_ok=True)
        try:
            run(SCRIPT, arguments.docx.read_bytes(), ORACLE_DIR / arguments.docx.name, arguments.pdf)
        except RuntimeError as error:  # WordBusy too: the lock held too long, or Word in use
            raise SystemExit(str(error)) from error

    document = parse_package(arguments.docx)
    page_sizes, lines = read_pages(arguments.pdf)

    print(f"fixture : {arguments.docx}")
    print(f"oracle  : {arguments.pdf}")
    print(f"faces   : {', '.join(base_fonts(arguments.pdf))}")
    print(f"warnings: {document.warnings or 'none'}")
    print()

    failures = 0

    # -- the page box -------------------------------------------------------
    print("Page box -- authored w:pgSz against the exported /MediaBox")
    print(f"  {'page':>4}  {'authored':>18}  {'exported':>18}  {'predicted':>18}  residual")
    for index, section in enumerate(document.sections):
        # One page per section is enough for this check and is what the fixture
        # provides for section 2; section 1 spills, and every one of its pages has
        # the same box.
        pages = [i for i, _ in enumerate(page_sizes)]
        page = pages[0] if index == 0 else pages[-1]
        authored = (
            twips_to_pt(section.page_size.width_twips),
            twips_to_pt(section.page_size.height_twips),
        )
        predicted = (
            device_page_extent_pt(section.page_size.width_twips),
            device_page_extent_pt(section.page_size.height_twips),
        )
        exported = page_sizes[page]
        residual = max(abs(exported[0] - predicted[0]), abs(exported[1] - predicted[1]))
        status = "ok" if residual < 5e-4 else "FAIL"
        if status == "FAIL":
            failures += 1
        print(
            f"  {page + 1:>4}  {authored[0]:8.3f} x{authored[1]:8.3f}"
            f"  {exported[0]:8.3f} x{exported[1]:8.3f}"
            f"  {predicted[0]:8.3f} x{predicted[1]:8.3f}"
            f"  {residual:.6f}  {status}"
        )
    print()

    # -- block A, the indent ladder ----------------------------------------
    print("Indent ladder -- authored w:ind/@w:left against measured ink-box left")
    # Pair by text rather than by index.  An index correspondence between paragraphs
    # and lines holds only while no paragraph wraps, and the first thing this project
    # will do is make paragraphs wrap.
    ladder_lines = [line for line in lines if line.text.startswith("H indent")]
    ladder_paragraphs = [
        paragraph for paragraph in document.paragraphs if paragraph.text.startswith("H indent ")
    ]
    if len(ladder_lines) != len(ladder_paragraphs):
        print(
            f"  paragraphs {len(ladder_paragraphs)} but lines {len(ladder_lines)}"
            " -- a ladder entry wrapped; widen the column or shorten the label"
        )
        failures += 1
    else:
        base_line = ladder_lines[0]
        base_paragraph = ladder_paragraphs[0]
        base_twips = (
            base_paragraph.properties.indent.left
            if base_paragraph.properties and base_paragraph.properties.indent
            else 0
        ) or 0
        print(f"  {'twips':>6}  {'authored dx':>12}  {'measured dx':>12}  residual")
        worst = 0.0
        for line, paragraph in zip(ladder_lines, ladder_paragraphs):
            indent = paragraph.properties.indent if paragraph.properties else None
            twips = (indent.left if indent and indent.left else 0) or 0
            authored_dx = twips_to_pt(twips - base_twips)
            measured_dx = line.left - base_line.left
            residual = abs(measured_dx - authored_dx)
            worst = max(worst, residual)
            print(
                f"  {twips:>6}  {authored_dx:12.4f}  {measured_dx:12.4f}  {residual:.6f}"
            )
        print(f"  worst residual: {worst:.6f} pt")
        if worst > 1e-3:
            failures += 1
    print()

    # -- block B, the tab ladder -------------------------------------------
    print("Tab ladder -- authored w:tabs/@w:pos against measured ink-box left")
    tab_paragraph = next(
        (
            paragraph
            for paragraph in document.paragraphs
            if paragraph.properties and paragraph.properties.tab_stops
        ),
        None,
    )
    tab_line = next(
        (line for line in lines if line.text.count("H") >= 4 and len(line.text) <= 10), None
    )
    if tab_paragraph and tab_line:
        section = document.sections[0]
        column_left = twips_to_pt(section.margins.left)
        probes = [glyph for glyph in tab_line.glyphs if glyph.char == "H"]
        print(f"  {'stop tw':>8}  {'authored x':>11}  {'measured x':>11}  residual")
        worst = 0.0
        first = probes[0]
        for stop, glyph in zip(tab_paragraph.properties.tab_stops, probes[1:]):
            # The bearing cancels against the *first* probe glyph, which sits at the
            # column's left edge: measured_x - first.left is the distance from the
            # column origin to the stop, bearing-free.
            authored = twips_to_pt(stop.position_twips)
            measured = glyph.left - first.left
            residual = abs(measured - authored)
            worst = max(worst, residual)
            print(
                f"  {stop.position_twips:>8}  {authored:11.4f}  {measured:11.4f}  {residual:.6f}"
            )
        print(f"  worst residual: {worst:.6f} pt  (column left = {column_left:.3f} pt)")
    print()

    # -- block E, pagination -----------------------------------------------
    print("Pagination -- observed only; nothing here is predicted yet")
    numbered = [line for line in lines if re.fullmatch(r"L\d\d", line.text)]
    by_page: dict[int, list[str]] = {}
    for line in numbered:
        by_page.setdefault(line.page, []).append(line.text)
    for page in sorted(by_page):
        labels = by_page[page]
        print(f"  page {page + 1}: {labels[0]}..{labels[-1]}  ({len(labels)} lines)")
    # The line advance, measured from **one glyph**.
    #
    # Not from `Line.baseline`: the modal ink-box bottom flips between a line whose
    # digits are flat ("L11") and one whose digits overshoot ("L10"), which produced a
    # spurious four-valued advance the first time this was measured and read as a Word
    # behaviour.  The probe glyph is on every line, so its bearing and its overshoot
    # subtract out exactly -- the same principle the horizontal blocks are built on,
    # applied to the vertical axis.
    for page in sorted(by_page):
        on_page = [line for line in numbered if line.page == page]
        tops = []
        for line in on_page:
            probe = [glyph for glyph in line.glyphs if glyph.char == "L"]
            if probe:
                tops.append((page_sizes[page][1] - probe[0].bottom) * 300 / 72)
        if len(tops) < 2:
            continue
        rounded = [round(value) for value in tops]
        off_grid = max(abs(value - whole) for value, whole in zip(tops, rounded))
        gaps = [rounded[i + 1] - rounded[i] for i in range(len(rounded) - 1)]
        histogram = {gap: gaps.count(gap) for gap in sorted(set(gaps))}
        print(
            f"  page {page + 1} baselines: off the 1/300in grid by at most"
            f" {off_grid:.4f} px; gaps {histogram}"
        )
    print()

    print("FAILURES" if failures else "all checked quantities within tolerance")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
