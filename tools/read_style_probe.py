#!/usr/bin/env python3
"""Hold ``docx2svg.resolve`` to what Word drew for every style-probe case.

For each probe in ``make_style_probe.py``: export through ``tools/oracle.py``, read the
PDF with ``quartz_pdf`` (face, drawn size and text of every glyph), parse the ``.docx``
with ``docx2svg``, resolve every run and paragraph mark of every case through
``docx2svg.resolve`` -- the reader, not the generator, decides what each case asks -- and
compare glyph by glyph:

* **face, bold, italic** -- the ``/BaseFont`` of the text object the glyph is in;
* **size** -- the ``Tm`` scale, which is the size in device px rounded half up;
* **caps** -- uppercase text; **vanish** -- the glyph is absent;
* **indent** -- the pen x of the first glyph, for the numbering/indent cases.

A glyph Word drew in a *different family* from the one resolved is counted separately
when it is a known glyph fallback (Georgia has no Hebrew; Word draws East Asian text in
MS Mincho when the slot names a face that is not East Asian or names none) -- those are
font substitution, which is a later stage, not the cascade.  Everything else is a miss.

The report also prints what a literal reading of ECMA-376 17.7.3 predicts for the
toggle cases, so the deviations are visible next to the measurements.

Usage::

    python tools/read_style_probe.py [-v]
"""

from __future__ import annotations

import argparse
import json
import sys
import unicodedata
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_style_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg import parse_package  # noqa: E402
from docx2svg.model import Paragraph, Table  # noqa: E402
from docx2svg.resolve import (  # noqa: E402
    character_format, resolve_mark, resolve_paragraph, resolve_run, style_chain,
)
from docx2svg.vertical import twips_to_px  # noqa: E402

#: (resolved face, drawn family) pairs that are glyph fallback, not resolution.
FALLBACK = {
    ("Georgia", "TimesNewRoman"),  # Hebrew / Arabic in a face without them
    ("Courier New", "MS-Mincho"),  # an East Asian slot naming a non-East-Asian face
    (None, "TimesNewRoman"),  # complex-script theme reference without themeFontLang
    (None, "MS-Mincho"),  # East Asian theme reference with no face
}
MARGIN_PX = 300
OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "style-observations.json"


def drawn(base: str) -> tuple[str, bool, bool]:
    """``TimesNewRomanPS-BoldMT`` -> ``("TimesNewRoman", True, False)``."""
    family, _, style = base.partition("-")
    for suffix in ("PSMT", "MT", "PS"):
        if family.endswith(suffix) and family != suffix:
            family = family[: -len(suffix)]
            break
    style = style.removesuffix("MT")
    bold = "Bold" in style
    italic = "Italic" in style or "Oblique" in style
    rest = style.replace("Bold", "").replace("Italic", "").replace("Oblique", "").replace("Regular", "")
    if family == "MS":  # MS-Mincho
        return base, bold, italic
    return family + rest, bold, italic


def size_px(half_points: int) -> int:
    return int(Fraction(half_points, 2) * 300 / 72 + Fraction(1, 2))


def paragraphs(blocks) -> list[Paragraph]:
    out = []
    for block in blocks:
        if isinstance(block, Table):
            for row in block.rows:
                for cell in row:
                    out.extend(paragraphs(cell))
        else:
            out.append(block)
    return out


def expected(document, paragraph: Paragraph) -> list[tuple[str, object, bool, bool]]:
    """(char, format, caps, hidden) for every character of the paragraph, then the mark."""
    out = []
    for run in paragraph.runs:
        resolved = resolve_run(document, paragraph, run)
        for char in run.text:
            out.append((char, character_format(resolved, char, document),
                        bool(resolved.get("caps")), bool(resolved.get("vanish"))))
    mark = resolve_mark(document, paragraph)
    out.append((" ", character_format(mark, " ", document), False, bool(mark.get("vanish"))))
    return out


def observe(pdf: Path) -> list[list]:
    """Every glyph Word drew, in order: ``[char, font, size_px, x, line]``.

    NFKC undoes what shaping did to the text: Arabic comes out of the PDF as
    presentation forms, with the lam-alef ligature one glyph for two characters, and
    ``中文`` came back once as a Kangxi radical.  The order is already logical.  ``line``
    is an index; the first glyph of a line is where its paragraph starts (or its list
    number does).  This list is what ``--record`` writes: measurements, no font data.
    """
    out = []
    for index, line in enumerate(quartz_pdf.lines(quartz_pdf.read(pdf))):
        for run in line.runs:
            for i, glyph in enumerate(run.text):
                x = run.xs[i] if i < len(run.xs) else run.x
                for char in unicodedata.normalize("NFKC", glyph) or glyph:
                    out.append([char, run.font, run.size_px, round(x, 4), index])
    return out


def check_probe(p: probe.Probe, stream: list[list], verbose: bool = False
                ) -> tuple[int, int, int, list[str]]:
    """Compare every case of ``p`` with the drawn ``stream``; (glyphs, misses, fallbacks, report)."""
    document = parse_package(p.build())
    line_start = {}
    for glyph in stream:
        line_start.setdefault(glyph[4], glyph[3])
    cases = {case.label: case for case in p.cases}
    glyphs = misses = fallbacks = 0
    report = []
    for paragraph in paragraphs(document.body):
        label = paragraph.text[:3]
        if label not in cases:
            continue
        want = [e for e in expected(document, paragraph) if not e[3]]
        # Find the case's label in the drawn stream (case-insensitively: caps).
        start = next((i for i in range(len(stream) - 2)
                      if "".join(stream[i + k][0] for k in range(3)).lower() == label), None)
        if start is None:
            report.append(f"  {p.name} {label}: label not drawn")
            misses += 1
            continue
        got = stream[start:start + len(want)]
        problems = []
        for (char, fmt, caps, _), (g_char, g_font, g_size, _, _) in zip(want, got):
            glyphs += 1
            family, bold, italic = drawn(g_font)
            want_char = char.upper() if caps else char
            want_family = fmt.face.replace(" ", "") if fmt.face else None
            issue = []
            if g_char != want_char:
                issue.append(f"text {g_char!r} for {want_char!r}")
            fallback = family != want_family and (fmt.face, family) in FALLBACK
            if family != want_family and not fallback:
                issue.append(f"face {g_font} for {fmt.face}")
            if not family.startswith("MS-") and (bold, italic) != (fmt.bold, fmt.italic):
                issue.append(f"b/i {(bold, italic)} for {(fmt.bold, fmt.italic)}")
            if g_size != size_px(fmt.half_points):
                issue.append(f"size {g_size:g}px for {fmt.half_points} hp")
            if issue:
                misses += 1
                problems.append(f"{char!r}: " + ", ".join(issue))
            elif fallback:
                fallbacks += 1
        if len(got) < len(want):
            problems.append(f"only {len(got)} of {len(want)} glyphs drawn")
            misses += 1
        # Indent: pen x of the first text glyph against the resolved left indent.  Only
        # a paragraph that begins its own line has one to check: a hidden paragraph
        # mark joins the next paragraph onto the same line (t20-t22).
        pp = resolve_paragraph(document, paragraph)
        numbered = "numPr.numId" in pp
        first_on_line = start == 0 or stream[start - 1][4] != got[0][4] or numbered
        if first_on_line and (paragraph.table_style_id is None or pp.get("ind.left")):
            left = pp.get("ind.left", 0) + pp.get("ind.firstLine", 0)
            x_want = MARGIN_PX + twips_to_px(left)
            if abs(got[0][3] - float(x_want)) > 1e-3:
                problems.append(f"text x {got[0][3]:.3f} for {float(x_want):.3f} (ind {left})")
                misses += 1
            if numbered:
                x_num = MARGIN_PX + twips_to_px(pp.get("ind.left", 0) - pp.get("ind.hanging", 0))
                if abs(line_start[got[0][4]] - float(x_num)) > 1e-3:
                    problems.append(f"number x {line_start[got[0][4]]:.3f} for {float(x_num):.3f}")
                    misses += 1
        status = "ok" if not problems else "MISS"
        if problems or verbose:
            report.append(f"  {p.name} {label} {status}: {cases[label].note}")
            for problem in problems:
                report.append(f"      {problem}")
            if verbose:
                sample = paragraph.runs[-1]
                resolved = resolve_run(document, paragraph, sample)
                for text in resolved.explain(["b", "i", "caps", "vanish", "sz"]).splitlines():
                    report.append(f"      {text}")
    return glyphs, misses, fallbacks, report


def ecma_literal(p: probe.Probe) -> list[str]:
    """Bold of each case's sample under a literal reading of ECMA-376 17.7.3.

    Start from docDefaults; every style in the hierarchy -- each style of each
    ``w:basedOn`` chain, table then paragraph then character -- that declares ``w:b``
    true flips the value, and false changes nothing; direct formatting is absolute.  The
    default paragraph style applies when none is named, as the spec says it does.
    """
    document = parse_package(p.build())
    sheet = document.styles
    out = []
    for paragraph in paragraphs(document.body):
        label = paragraph.text[:3]
        if not paragraph.runs:
            continue
        run = paragraph.runs[-1]
        value = bool(sheet.run_defaults.get("b", False))
        chains = []
        if paragraph.table_style_id is not None:
            table = paragraph.table_style_id or (sheet.default("table") or Paragraph()).style_id
            chains.append(style_chain(sheet, table, "table"))
        pstyle = paragraph.properties.style_id if paragraph.properties else None
        if pstyle not in sheet.styles or sheet.styles[pstyle].kind != "paragraph":
            pstyle = sheet.default("paragraph").style_id
        chains.append(style_chain(sheet, pstyle, "paragraph"))
        rstyle = run.properties.style_id if run.properties else None
        if rstyle:
            chains.append(style_chain(sheet, rstyle, "character"))
        for chain in chains:
            # Literally, the default character style is in the chain like any other.
            root = chain[0] if chain else None
            while root is not None and root.based_on in sheet.styles:
                root = sheet.styles[root.based_on]
                chain = [root] + chain
            for style in chain:
                if style.run.get("b") is True:
                    value = not value
        direct = run.properties.declared if run.properties else {}
        if "b" in direct:
            value = direct["b"]
        word = bool(resolve_run(document, paragraph, run).get("b"))
        if value != word:
            out.append(f"  {p.name} {label}: ECMA-376 literal says bold={value}, Word drew bold={word}")
    return out


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--record", action="store_true",
                        help=f"write the drawn glyphs to {OBSERVATIONS.name}")
    args = parser.parse_args(argv[1:])
    total = [0, 0, 0]
    lines = []
    recorded = {}
    for p in probe.probes():
        stream = observe(oracle.export(p.build(), name=p.name))
        recorded[p.name] = stream
        glyphs, misses, fallbacks, report = check_probe(p, stream, args.verbose)
        total[0] += glyphs
        total[1] += misses
        total[2] += fallbacks
        print(f"{p.name:28} {len(p.cases):3} cases  {glyphs:4} glyphs  {misses:3} misses"
              f"  {fallbacks:3} glyph fallbacks")
        lines += report
    print("\n".join(lines))
    print(f"\nall probes: {total[0]} glyphs, {total[1]} misses, {total[2]} drawn in a fallback face")
    print("\nWhere a literal reading of ECMA-376 17.7.3 disagrees with what Word drew:")
    for p in probe.probes():
        if p.name in ("style-toggles", "style-defaults-bold", "style-normal-bold",
                      "style-charstyle-bold"):
            print("\n".join(ecma_literal(p)))
    if args.record:
        payload = {
            "_about": (
                "Every glyph Word 16.106 drew for each probe of tools/make_style_probe.py: "
                "[char, /BaseFont, drawn size in device px, pen x in device px, line index]. "
                "Measurements only; regenerate with tools/read_style_probe.py --record."
            ),
            "probes": recorded,
        }
        OBSERVATIONS.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 1 if total[1] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
