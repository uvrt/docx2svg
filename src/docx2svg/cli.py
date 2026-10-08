"""Command-line interface::

    docx2svg report.docx -o out/             # one SVG per page
    docx2svg report.docx -f png -o out/      # PNG (needs the png extra)
    docx2svg report.docx -p 1,3-4            # some pages

Warnings -- what the output does not show faithfully, each with a stable code -- go to
standard error unless ``-q``, after a one-line coverage summary (how much was laid out,
where it stopped, which faces were substituted: :mod:`docx2svg.coverage`); ``--strict``
makes any warning an exit status of 2, so a build can refuse a document the renderer
cannot draw faithfully -- a substituted face included.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import ConvertOptions, __version__, _lay_out, _rasterise, _svgs, available_backends


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="docx2svg", description="Render Word (.docx) documents to SVG or PNG.")
    parser.add_argument("input", type=Path, help="path to a .docx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("."),
                        help="output directory (default: current directory)")
    parser.add_argument("-f", "--format", choices=("svg", "png", "both"), default="svg",
                        help="output format (default: svg)")
    parser.add_argument("-p", "--pages", help="page numbers to render, e.g. '1', '1,3', '2-5' (default: all)")
    parser.add_argument("--width", type=int, help="output width in pixels")
    parser.add_argument("--height", type=int, help="output height in pixels")
    parser.add_argument("--glyph-size", choices=("device", "exact"), default=ConvertOptions.glyph_size,
                        help="scale glyph outlines to the size rounded to whole 1/300-inch pixels, as Word draws "
                             "its ink (device, the default), or to the exact size")
    parser.add_argument("--font-dir", action="append", dest="font_dirs", metavar="DIR",
                        help="another directory to find faces in (repeatable)")
    parser.add_argument("--no-substitute-fonts", action="store_false", dest="substitute_fonts",
                        help="do not lay out an absent Office face with its open metric compatible substitute "
                             "(Carlito, Liberation) or a symbol face from its recorded metrics")
    parser.add_argument("--backend", choices=("auto", "resvg", "cairosvg"), default="auto",
                        help="PNG rasterizer backend (default: auto)")
    parser.add_argument("-q", "--quiet", action="store_true", help="do not print warnings")
    parser.add_argument("--strict", action="store_true", help="exit with status 2 when there is any warning")
    parser.add_argument("--version", action="version", version=f"docx2svg {__version__}")
    return parser


def parse_page_selection(value: str | None) -> list[int] | None:
    """``"1,3,5-7"`` -> ``[1, 3, 5, 6, 7]``."""
    if not value:
        return None
    numbers: list[int] = []
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, _, end = part.partition("-")
            numbers.extend(range(int(start), int(end) + 1))
        else:
            numbers.append(int(part))
    return numbers or None


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(sys.argv[1:] if argv is None else list(argv))
    if not args.input.is_file():
        print(f"docx2svg: no such file: {args.input}", file=sys.stderr)
        return 1
    options = ConvertOptions(pages=parse_page_selection(args.pages), width=args.width, height=args.height,
                             glyph_size=args.glyph_size, font_dirs=args.font_dirs,
                             substitute_fonts=args.substitute_fonts)
    wants_svg = args.format in ("svg", "both")
    wants_png = args.format in ("png", "both")
    if wants_png and not available_backends():
        print("docx2svg: PNG output needs a rasterizer; run `pip install docx2svg[png]`", file=sys.stderr)
        return 1
    try:
        layout, data, fonts = _lay_out(args.input, options)
        documents = _svgs(layout, data, options, fonts=fonts)
        pngs = _rasterise(layout, data, fonts, options, backend=args.backend) if wants_png else []
    except Exception as error:  # a malformed package, unreadable XML...
        print(f"docx2svg: {error}", file=sys.stderr)
        return 1
    if not documents:
        print("docx2svg: no pages matched the selection", file=sys.stderr)
        return 1
    args.output.mkdir(parents=True, exist_ok=True)
    numbers = options.pages or range(1, len(documents) + 1)
    stem = args.input.stem
    for index, (number, document) in enumerate(zip(numbers, documents)):
        if wants_svg:
            path = args.output / f"{stem}-{number}.svg"
            path.write_bytes(document.encode("utf-8"))
            print(path)
        if wants_png:
            path = args.output / f"{stem}-{number}.png"
            path.write_bytes(pngs[index])
            print(path)
    if not args.quiet and options.coverage is not None and (options.warnings or not options.coverage.complete):
        print(f"\ncoverage: {options.coverage.summary()}", file=sys.stderr)
    if options.warnings and not args.quiet:
        print(f"\n{len(options.warnings)} warning(s):", file=sys.stderr)
        for warning in options.warnings:
            print(f"  {warning}", file=sys.stderr)
    return 2 if (args.strict and options.warnings) else 0


if __name__ == "__main__":
    raise SystemExit(main())
