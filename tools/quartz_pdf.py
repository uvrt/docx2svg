#!/usr/bin/env python3
"""Read Word's PDF export at the precision Word wrote it, not the precision pdfium reports.

Why this exists
---------------

``read_layout_sweep.py`` measures glyph *ink boxes* through pypdfium2, in float32 points.
That closed the Phase 0 loop at 8e-6 pt, and for differences of whole twips it is
plenty.  Phase 2 asks a sharper question -- *is a position rounded, and to what grid?* --
and for that the ink box is the wrong instrument twice over: it is float32, and it is a
function of the glyph outline rather than of the pen position Word chose.

Word 16.106's export goes through Quartz (``/Producer`` reads "Quartz PDFContext"), and
Quartz writes a very regular content stream::

    q 0.24 0 0 0.24 0 640.08 cm BT -0.0022 Tc 46 0 0 46 300 426 Tm /TT4 1 Tf
      [ ("!) -2 (#) -2 ($) ... ] TJ ET Q

Three facts in that line are the reason this module exists, and each was read off the
stream rather than inferred:

* ``0.24 0 0 0.24`` -- the whole page is drawn in **1/300-inch device pixels**
  (0.24 pt = 1/300 in).  Every coordinate below is therefore already in the unit the
  vertical grid was found in, and a position that is a whole device pixel is written as
  an integer.
* ``Tm`` carries the pen position of the *first* glyph of each text object to four
  decimals.  Quartz starts a new text object whenever Word positions a glyph on its own
  (a tab, the paragraph mark, a new run), so those positions are exact to 1e-4 px.
* Inside a ``TJ`` the per-glyph positions are *not* exact: ``/Widths`` are integers per
  1000 em, and Quartz makes up the difference with a ``Tc`` and integer ``TJ``
  adjustments, i.e. to within 0.5/1000 em (0.023 px at 11 pt).  That is fine for
  detecting a whole-pixel grid and it is noted wherever it matters.

Also recorded here because it cost a moment: the ``Tm`` scale is the font size
*rounded to whole device pixels* (46 for 11 pt, which is 45.83 px), while the pen
advances are the unrounded size's.  So the ink is drawn at 46 ppem and placed at 45.83.

Standard library only (zlib).  This reads the Quartz dialect Word emits; it is not a
general PDF parser and makes no claim to be.
"""

from __future__ import annotations

import re
import zlib
from dataclasses import dataclass, field
from pathlib import Path

_OBJ = re.compile(rb"(\d+) 0 obj(.*?)endobj", re.S)
_STREAM = re.compile(rb"stream\r?\n(.*)endstream", re.S)
#: A font resource entry.  Quartz wraps long dictionaries, so the name and its reference
#: can be on different lines (``/TT13\n18 0 R``) -- found when a probe used nine faces.
_REF = re.compile(rb"/(\w+)\s+(\d+) 0 R")


@dataclass
class Font:
    base: str
    first: int
    widths: list[int]
    to_unicode: dict[int, str]

    def width(self, code: int) -> int:
        return self.widths[code - self.first]


@dataclass
class TextRun:
    """One Quartz text object: a pen start, a size, and the glyphs shown from it."""

    page: int
    font: str
    size_px: float  # the Tm scale: the size Quartz *draws* at, in device px
    x: float  # pen x of the first glyph, device px from the page's left edge
    y: float  # baseline, device px from the page's *top* edge
    text: str
    #: Pen x of every glyph.  The first is exact; the rest carry Quartz's TJ
    #: quantisation of up to 0.5/1000 em.
    xs: list[float] = field(default_factory=list)
    #: Pen x after the last glyph (same quantisation as ``xs``).
    end: float = 0.0


@dataclass
class Page:
    width_pt: float
    height_pt: float
    runs: list[TextRun]


def _objects(data: bytes) -> dict[int, bytes]:
    return {int(m.group(1)): m.group(2) for m in _OBJ.finditer(data)}


def _stream(body: bytes) -> bytes:
    m = _STREAM.search(body)
    if not m:
        return b""
    raw = m.group(1)
    if b"/FlateDecode" in body.split(b"stream", 1)[0]:
        return zlib.decompress(raw)
    return raw


def _dict_ref(body: bytes, key: str) -> int | None:
    m = re.search(rb"/" + key.encode() + rb"\s+(\d+) 0 R", body)
    return int(m.group(1)) if m else None


def _parse_cmap(text: bytes) -> dict[int, str]:
    out: dict[int, str] = {}
    for block in re.findall(rb"beginbfchar(.*?)endbfchar", text, re.S):
        for src, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            out[int(src, 16)] = bytes.fromhex(dst.decode()).decode("utf-16-be")
    for block in re.findall(rb"beginbfrange(.*?)endbfrange", text, re.S):
        for lo, hi, dst in re.findall(
            rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block
        ):
            start = int(dst, 16)
            for offset, code in enumerate(range(int(lo, 16), int(hi, 16) + 1)):
                out[code] = chr(start + offset)
    return out


def _font(objects: dict[int, bytes], number: int) -> Font:
    body = objects[number]
    base = re.search(rb"/BaseFont /([^\s/]+)", body).group(1).decode()
    first = int(re.search(rb"/FirstChar (\d+)", body).group(1))
    widths_ref = re.search(rb"/Widths\s+(\d+) 0 R", body)
    widths_src = objects[int(widths_ref.group(1))] if widths_ref else body
    widths = [int(w) for w in re.search(rb"/Widths\s*\[([^\]]*)\]|\[([^\]]*)\]", widths_src).group(
        0
    ).split(b"[", 1)[1].rstrip(b"]").split()]
    tu = _dict_ref(body, "ToUnicode")
    to_unicode = _parse_cmap(_stream(objects[tu])) if tu else {}
    if not tu and b"/MacRomanEncoding" in body:
        # A simple font Quartz wrote with a standard encoding carries no ToUnicode map;
        # its codes *are* MacRoman.
        to_unicode = {code: bytes([code]).decode("mac_roman") for code in range(32, 256)}
    return Font(base.split("+", 1)[-1], first, widths, to_unicode)


_TOKEN = re.compile(
    rb"\((?:\\.|[^\\)])*\)|\[|\]|/[^\s/\[\]()<>]+|<[0-9A-Fa-f\s]*>|[-+]?\d*\.?\d+|[A-Za-z'\"*]+"
)


def _unescape(literal: bytes) -> bytes:
    out = bytearray()
    i = 1
    end = len(literal) - 1
    while i < end:
        c = literal[i]
        if c == 0x5C:  # backslash
            i += 1
            e = literal[i]
            mapping = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}
            if e in mapping:
                out.append(mapping[e])
            elif 0x30 <= e <= 0x37:
                digits = bytes([e])
                while len(digits) < 3 and 0x30 <= literal[i + 1] <= 0x37:
                    i += 1
                    digits += bytes([literal[i]])
                out.append(int(digits, 8))
            else:
                out.append(e)
        else:
            out.append(c)
        i += 1
    return bytes(out)


def read(pdf_path: Path | str) -> list[Page]:
    data = Path(pdf_path).read_bytes()
    objects = _objects(data)
    fonts: dict[int, Font] = {}

    page_numbers = [
        n for n, body in objects.items() if re.search(rb"/Type\s*/Page[^s]", body)
    ]
    # Document order is the /Kids order of the page tree, not object-number order.
    kids: list[int] = []
    for body in objects.values():
        if re.search(rb"/Type\s*/Pages", body):
            kids = [int(k) for k in re.findall(rb"(\d+) 0 R", re.search(rb"/Kids\s*\[([^\]]*)\]", body).group(1))]
    order = [k for k in kids if k in page_numbers] or sorted(page_numbers)

    pages: list[Page] = []
    for page_index, number in enumerate(order):
        body = objects[number]
        box = [float(v) for v in re.search(rb"/MediaBox\s*\[([^\]]*)\]", body).group(1).split()]
        width_pt, height_pt = box[2] - box[0], box[3] - box[1]
        resources = body
        res_ref = _dict_ref(body, "Resources")
        if res_ref is not None:
            resources = objects[res_ref]
        # A page with no text (an image alone on it) has no /Font resource at all.
        found = re.search(rb"/Font\s*(<<.*?>>|\d+ 0 R)", resources, re.S)
        font_dict = found.group(1) if found else b""
        if font_dict.endswith(b"R"):
            font_dict = objects[int(font_dict.split()[0])]
        page_fonts: dict[str, Font] = {}
        for name, ref in _REF.findall(font_dict):
            ref = int(ref)
            if ref not in fonts:
                fonts[ref] = _font(objects, ref)
            page_fonts[name.decode()] = fonts[ref]

        contents = _dict_ref(body, "Contents")
        stream = _stream(objects[contents])
        pages.append(Page(width_pt, height_pt, _interpret(stream, page_fonts, page_index, height_pt)))
    return pages


def _interpret(stream: bytes, fonts: dict[str, Font], page: int, height_pt: float) -> list[TextRun]:
    runs: list[TextRun] = []
    stack: list[float] = []
    operands: list = []
    ctm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    saved: list[list[float]] = []
    tc = 0.0
    font: Font | None = None
    tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    array: list | None = None
    for token in _TOKEN.findall(stream):
        if token == b"[":
            array = []
            continue
        if token == b"]":
            operands.append(array)
            array = None
            continue
        if token.startswith(b"("):
            value = _unescape(token)
        elif token.startswith(b"/"):
            value = token[1:].decode()
        elif re.fullmatch(rb"[-+]?\d*\.?\d+", token):
            value = float(token)
        elif token.startswith(b"<"):
            value = bytes.fromhex(token[1:-1].decode())
        else:
            op = token.decode()
            if array is not None:
                continue
            if op == "q":
                saved.append(list(ctm))
            elif op == "Q":
                ctm = saved.pop() if saved else [1, 0, 0, 1, 0, 0]
            elif op == "cm":
                a, b, c, d, e, f = operands[-6:]
                # Quartz only ever concatenates onto an identity-scaled page here.
                A, B, C, D, E, F = ctm
                ctm = [a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D,
                       e * A + f * C + E, e * B + f * D + F]
            elif op == "BT":
                tc = 0.0
                tm = [1, 0, 0, 1, 0, 0]
            elif op == "Tc":
                tc = operands[-1]
            elif op == "Tm":
                tm = list(operands[-6:])
            elif op == "Tf":
                font = fonts[operands[-2]]
            elif op in ("Tj", "TJ") and font is not None:
                items = operands[-1] if op == "TJ" else [operands[-1]]
                scale = tm[0]
                pen = tm[4]
                xs: list[float] = []
                chars: list[str] = []
                for item in items:
                    if isinstance(item, float):
                        pen -= item / 1000.0 * scale
                        continue
                    for code in item:
                        xs.append(pen)
                        chars.append(font.to_unicode.get(code, "?"))
                        pen += (font.width(code) / 1000.0 + tc) * scale
                # Convert Tm space (device px under a 0.24 pt cm) to page px.
                unit = ctm[0]
                to_px = unit / 0.24
                x0 = ctm[4] / 0.24
                y_pt = ctm[5] + unit * tm[5]
                runs.append(
                    TextRun(
                        page=page,
                        font=font.base,
                        size_px=scale,
                        x=x0 + tm[4] * to_px,
                        y=(height_pt - y_pt) / 0.24,
                        text="".join(chars),
                        xs=[x0 + v * to_px for v in xs],
                        end=x0 + pen * to_px,
                    )
                )
            operands = []
            continue
        if array is not None:
            array.append(value)
        else:
            operands.append(value)
    return runs


@dataclass(frozen=True)
class Fill:
    """An axis-aligned filled rectangle Word drew (a rule, an underline, a highlight, a
    border), in device px from the page's top-left, and its fill colour."""

    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    color: str  # RRGGBB


def fills(pdf_path: Path | str) -> list[Fill]:
    """Every filled axis-aligned rectangle of every page (``re`` or a four-corner path,
    then ``f``), in page order.  Word's export draws underlines, strikes, highlights,
    shading and paragraph borders this way, snapped to whole device pixels."""
    data = Path(pdf_path).read_bytes()
    objects = _objects(data)
    page_numbers = [n for n, body in objects.items() if re.search(rb"/Type\s*/Page[^s]", body)]
    kids: list[int] = []
    for body in objects.values():
        if re.search(rb"/Type\s*/Pages", body):
            kids = [int(k) for k in re.findall(rb"(\d+) 0 R", re.search(rb"/Kids\s*\[([^\]]*)\]", body).group(1))]
    order = [k for k in kids if k in page_numbers] or sorted(page_numbers)
    out: list[Fill] = []
    for index, number in enumerate(order):
        body = objects[number]
        box = [float(v) for v in re.search(rb"/MediaBox\s*\[([^\]]*)\]", body).group(1).split()]
        height = box[3] - box[1]
        stream = _stream(objects[_dict_ref(body, "Contents")])
        stream = re.sub(rb"BT.*?ET", b" ", stream, flags=re.S)
        ctm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
        saved: list = []
        color = "000000"
        colors: list = []
        operands: list[float] = []
        path: list[tuple[float, float]] = []
        rects: list[tuple[float, float, float, float]] = []

        def point(x, y):
            a, b, c, d, e, f = ctm
            return a * x + c * y + e, b * x + d * y + f

        for token in re.findall(rb"/[^\s/\[\]()<>]+|[-+]?\d*\.?\d+|[A-Za-z*'\"]+", stream):
            if re.fullmatch(rb"[-+]?\d*\.?\d+", token):
                operands.append(float(token))
                continue
            if token.startswith(b"/"):
                continue
            op = token.decode()
            if op == "q":
                saved.append((list(ctm), color))
            elif op == "Q":
                ctm, color = saved.pop() if saved else ([1, 0, 0, 1, 0, 0], "000000")
            elif op == "cm" and len(operands) >= 6:
                a, b, c, d, e, f = operands[-6:]
                A, B, C, D, E, F = ctm
                ctm = [a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E,
                       e * B + f * D + F]
            elif op in ("sc", "scn", "rg") and len(operands) >= 3:
                color = "".join(f"{round(v * 255):02X}" for v in operands[-3:])
            elif op in ("g",) and operands:
                color = f"{round(operands[-1] * 255):02X}" * 3
            elif op == "re" and len(operands) >= 4:
                x, y, w, h = operands[-4:]
                corners = [point(x, y), point(x + w, y + h)]
                rects.append((corners[0][0], corners[0][1], corners[1][0], corners[1][1]))
            elif op == "m" and len(operands) >= 2:
                path = [point(*operands[-2:])]
            elif op == "l" and len(operands) >= 2:
                path.append(point(*operands[-2:]))
            elif op == "h":
                if len(path) >= 4 and len({round(p[0], 3) for p in path}) == 2 and len({round(p[1], 3) for p in path}) == 2:
                    xs = [p[0] for p in path]
                    ys = [p[1] for p in path]
                    rects.append((min(xs), min(ys), max(xs), max(ys)))
                path = []
            elif op in ("f", "f*", "F", "B", "B*"):
                for x0, y0, x1, y1 in rects:
                    x0, x1 = sorted((x0, x1))
                    y0, y1 = sorted((y0, y1))
                    out.append(Fill(index, x0 / 0.24, (height - y1) / 0.24, x1 / 0.24, (height - y0) / 0.24, color))
                rects = []
                path = []
            elif op in ("n", "S", "s", "W", "W*"):
                if op in ("n", "S", "s"):
                    rects = []
                    path = []
            operands = []
        del colors
    return out


@dataclass
class Line:
    page: int
    y: float
    runs: list[TextRun]

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)


def lines(pages: list[Page]) -> list[Line]:
    """Runs grouped by baseline, in document order (page, then top to bottom).

    Baselines are compared after rounding to 1e-3 px: Quartz writes them as integers in
    device space, and the float noise is the ``cm`` product's, around 1e-12.
    """
    out: list[Line] = []
    for index, page in enumerate(pages):
        by_y: dict[float, list[TextRun]] = {}
        for run in page.runs:
            by_y.setdefault(round(run.y, 3), []).append(run)
        for y in sorted(by_y):
            out.append(Line(index, y, sorted(by_y[y], key=lambda r: r.x)))
    return out


if __name__ == "__main__":
    import sys

    for page in read(sys.argv[1]):
        print(f"page {page.width_pt} x {page.height_pt}")
        for run in page.runs:
            print(f"  {run.font:>16} {run.size_px:5g} x={run.x:10.4f} y={run.y:10.4f} {run.text!r}")
