"""Glyph outlines' boxes and OpenType pair kerning, read with the standard library.

What a VML text path that fits its shape needs of a face (:meth:`docx2svg.fonts.Face.ink`,
:func:`docx2svg.drawing.text_path_markup`), and nothing else:

* **The box round a glyph's outline**, exact -- the curves' extremes, not their control
  points -- from a TrueType ``glyf`` outline (quadratic, composite glyphs followed) or a
  CFF ``CharStrings`` Type 2 program (cubic, subroutines followed), optionally slanted
  (``x + slant * y``: the oblique Word draws a face that has no italic, measured in
  :data:`docx2svg.drawing.TEXT_PATH_OBLIQUE`).  A ``glyf`` header's box is the box of the
  upright outline already; the outline itself is read for a slant, and for CFF, whose
  charstrings state no box.
* **Pair kerning from ``GPOS``** -- the ``kern`` feature's pair adjustment lookups
  (``PairPos`` formats 1 and 2, through ``Extension`` lookups), each lookup's first
  subtable that covers the pair -- which Word applies to a watermark's letters (measured:
  Minion Pro, a CFF face with no legacy ``kern`` table, is kerned by its ``GPOS`` pairs).

What is not read returns ``None`` (an outline this reader cannot follow: a CFF2 or
CID-keyed table, a ``seac`` accent, Type 2 arithmetic, a composite placed by point
numbers) and the caller falls back as it would for a face without outlines.  The format
references are the OpenType specification's ``glyf``, ``GPOS`` and ``CFF`` chapters and
Adobe's Technical Notes 5176 (CFF) and 5177 (Type 2 charstrings).
"""

from __future__ import annotations

import struct

Box = tuple  # (x min, y min, x max, y max), font units, y up


# -- boxes -------------------------------------------------------------------------------------

class _Bounds:
    """The running box of segments, each in the slanted frame (``x + slant * y``)."""

    def __init__(self, slant: float) -> None:
        self.slant = slant
        self.box: list | None = None

    def point(self, x: float, y: float) -> None:
        self._slanted(x + self.slant * y, y)

    def _slanted(self, x: float, y: float) -> None:
        if self.box is None:
            self.box = [x, y, x, y]
        else:
            box = self.box
            box[0], box[1], box[2], box[3] = min(box[0], x), min(box[1], y), max(box[2], x), max(box[3], y)

    def quad(self, p0, p1, p2) -> None:
        """A quadratic segment's extremes (its end ``p2``; ``p0`` was added before)."""
        s = self.slant
        a, b, c = ((p[0] + s * p[1], p[1]) for p in (p0, p1, p2))
        self._slanted(*c)
        for axis in (0, 1):
            d = a[axis] - 2 * b[axis] + c[axis]
            if d:
                t = (a[axis] - b[axis]) / d
                if 0 < t < 1:
                    weights = ((1 - t) ** 2, 2 * t * (1 - t), t * t)
                    self._slanted(*(sum(w * p[k] for w, p in zip(weights, (a, b, c))) for k in (0, 1)))

    def cubic(self, p0, p1, p2, p3) -> None:
        """A cubic segment's extremes (its end ``p3``; ``p0`` was added before)."""
        s = self.slant
        a, b, c, d = ((p[0] + s * p[1], p[1]) for p in (p0, p1, p2, p3))
        self._slanted(*d)
        for axis in (0, 1):
            # Where the derivative, 3 (q t^2 + r t + u), is zero.
            q = -a[axis] + 3 * b[axis] - 3 * c[axis] + d[axis]
            r = 2 * (a[axis] - 2 * b[axis] + c[axis])
            u = b[axis] - a[axis]
            roots = []
            if abs(q) < 1e-12:
                if r:
                    roots.append(-u / r)
            else:
                disc = r * r - 4 * q * u
                if disc >= 0:
                    root = disc ** 0.5
                    roots += [(-r + root) / (2 * q), (-r - root) / (2 * q)]
            for t in roots:
                if 0 < t < 1:
                    m = 1 - t
                    weights = (m ** 3, 3 * t * m * m, 3 * t * t * m, t ** 3)
                    self._slanted(*(sum(w * p[k] for w, p in zip(weights, (a, b, c, d))) for k in (0, 1)))


# -- TrueType ----------------------------------------------------------------------------------

def glyf_box(glyf: bytes, loca: bytes, long_offsets: bool, glyph: int, slant: float = 0.0,
             depth: int = 0) -> Box | None:
    """The exact box of TrueType glyph ``glyph``'s outline, slanted by ``slant``; ``()`` for
    an empty glyph (a space), ``None`` for one this reader cannot follow."""
    bounds = _Bounds(slant)
    if not _glyf_outline(glyf, loca, long_offsets, glyph, (1.0, 0.0, 0.0, 1.0, 0.0, 0.0), bounds, depth):
        return None
    return () if bounds.box is None else tuple(bounds.box)


def glyph_range(loca: bytes, long_offsets: bool, glyph: int) -> tuple[int, int] | None:
    """Where glyph ``glyph``'s data lies in ``glyf`` (start, end), from ``loca``; ``None`` past its end."""
    size = 4 if long_offsets else 2
    if (glyph + 2) * size > len(loca):
        return None
    start, end = struct.unpack_from(">II" if long_offsets else ">HH", loca, glyph * size)
    return (start, end) if long_offsets else (2 * start, 2 * end)


def _glyf_outline(glyf, loca, long_offsets, glyph, matrix, bounds: _Bounds, depth: int) -> bool:
    found = glyph_range(loca, long_offsets, glyph)
    if found is None or depth > 8:
        return False
    start, end = found
    if end <= start:
        return True  # an empty glyph (a space): nothing drawn
    if end > len(glyf) or end - start < 10:
        return False
    contours = struct.unpack_from(">h", glyf, start)[0]
    at = start + 10
    if contours < 0:
        return _composite(glyf, loca, long_offsets, at, end, matrix, bounds, depth)
    ends = struct.unpack_from(f">{contours}H", glyf, at)
    at += 2 * contours
    count = ends[-1] + 1 if contours else 0
    at += 2 + struct.unpack_from(">H", glyf, at)[0]  # the instructions
    flags: list[int] = []
    while len(flags) < count:
        flag = glyf[at]
        at += 1
        repeat = 1
        if flag & 8:
            repeat += glyf[at]
            at += 1
        flags.extend([flag] * repeat)
    coordinates = []
    for short, same in ((2, 16), (4, 32)):
        value, values = 0, []
        for flag in flags[:count]:
            if flag & short:
                delta = glyf[at]
                at += 1
                value += delta if flag & same else -delta
            elif not flag & same:
                value += struct.unpack_from(">h", glyf, at)[0]
                at += 2
            values.append(value)
        coordinates.append(values)
    a, b, c, d, e, f = matrix
    points = [(a * x + c * y + e, b * x + d * y + f, bool(flag & 1))
              for x, y, flag in zip(coordinates[0], coordinates[1], flags)]
    first = 0
    for last in ends:
        _contour(points[first:last + 1], bounds)
        first = last + 1
    return True


def _contour(points: list, bounds: _Bounds) -> None:
    if not points:
        return
    # Start on an on-curve point (or the midpoint of the first two off-curve ones).
    on = next((i for i, p in enumerate(points) if p[2]), None)
    if on is None:
        start = ((points[0][0] + points[1 % len(points)][0]) / 2,
                 (points[0][1] + points[1 % len(points)][1]) / 2)
        rest = points
    else:
        start = points[on][:2]
        rest = points[on + 1:] + points[:on]
    bounds.point(*start)
    current, control = start, None
    for x, y, on_curve in rest + [(start[0], start[1], True)]:
        if on_curve:
            if control is None:
                bounds.point(x, y)
            else:
                bounds.quad(current, control, (x, y))
            current, control = (x, y), None
        elif control is None:
            control = (x, y)
        else:
            middle = ((control[0] + x) / 2, (control[1] + y) / 2)
            bounds.quad(current, control, middle)
            current, control = middle, (x, y)


def _composite(glyf, loca, long_offsets, at, end, matrix, bounds, depth) -> bool:
    while True:
        if at + 4 > end:
            return False
        flags, component = struct.unpack_from(">HH", glyf, at)
        at += 4
        if flags & 1:  # ARG_1_AND_2_ARE_WORDS
            dx, dy = struct.unpack_from(">hh", glyf, at)
            at += 4
        else:
            dx, dy = struct.unpack_from(">bb", glyf, at)
            at += 2
        if not flags & 2:  # ARGS_ARE_XY_VALUES: placed by point numbers instead
            return False
        a, b, c, d = 1.0, 0.0, 0.0, 1.0
        if flags & 8:  # WE_HAVE_A_SCALE
            a = d = struct.unpack_from(">h", glyf, at)[0] / 16384
            at += 2
        elif flags & 0x40:  # WE_HAVE_AN_X_AND_Y_SCALE
            a, d = (v / 16384 for v in struct.unpack_from(">hh", glyf, at))
            at += 4
        elif flags & 0x80:  # WE_HAVE_A_TWO_BY_TWO
            a, b, c, d = (v / 16384 for v in struct.unpack_from(">hhhh", glyf, at))
            at += 8
        pa, pb, pc, pd, pe, pf = matrix
        inner = (pa * a + pc * b, pb * a + pd * b, pa * c + pc * d, pb * c + pd * d,
                 pa * dx + pc * dy + pe, pb * dx + pd * dy + pf)
        if not _glyf_outline(glyf, loca, long_offsets, component, inner, bounds, depth + 1):
            return False
        if not flags & 0x20:  # MORE_COMPONENTS
            return True


# -- CFF ---------------------------------------------------------------------------------------

def _index(data: bytes, at: int) -> tuple[list[bytes], int]:
    """A CFF INDEX at ``at``: its items and the offset after it."""
    count = struct.unpack_from(">H", data, at)[0]
    if count == 0:
        return [], at + 2
    size = data[at + 2]
    offsets = [int.from_bytes(data[at + 3 + i * size:at + 3 + (i + 1) * size], "big") for i in range(count + 1)]
    base = at + 2 + (count + 1) * size
    return [data[base + offsets[i]:base + offsets[i + 1]] for i in range(count)], base + offsets[-1]


def _dict(data: bytes) -> dict[int, list]:
    """A CFF DICT: operator (``12 x`` as ``1200 + x``) to its operands."""
    out: dict[int, list] = {}
    operands: list = []
    at = 0
    while at < len(data):
        b0 = data[at]
        if b0 <= 21:
            if b0 == 12:
                out[1200 + data[at + 1]] = operands
                at += 2
            else:
                out[b0] = operands
                at += 1
            operands = []
        elif b0 == 28:
            operands.append(struct.unpack_from(">h", data, at + 1)[0])
            at += 3
        elif b0 == 29:
            operands.append(struct.unpack_from(">i", data, at + 1)[0])
            at += 5
        elif b0 == 30:  # a real number, nibble coded: not an offset, skipped
            at += 1
            while at < len(data):
                byte = data[at]
                at += 1
                if byte & 0x0F == 0x0F or byte >> 4 == 0x0F:
                    break
            operands.append(0)
        elif 32 <= b0 <= 246:
            operands.append(b0 - 139)
            at += 1
        elif 247 <= b0 <= 250:
            operands.append((b0 - 247) * 256 + data[at + 1] + 108)
            at += 2
        elif 251 <= b0 <= 254:
            operands.append(-(b0 - 251) * 256 - data[at + 1] - 108)
            at += 2
        else:
            raise ValueError("a reserved byte in a CFF DICT")
    return out


def _bias(count: int) -> int:
    return 107 if count < 1240 else 1131 if count < 33900 else 32768


class CFF:
    """A ``CFF `` table's charstrings, for their outlines' boxes (name-keyed fonts only)."""

    def __init__(self, table: bytes) -> None:
        at = table[2]  # hdrSize
        _, at = _index(table, at)  # Name INDEX
        tops, at = _index(table, at)
        _, at = _index(table, at)  # String INDEX
        self.global_subrs, _ = _index(table, at)
        top = _dict(tops[0])
        if 1230 in top or top.get(1206, [2])[0] != 2:  # ROS (CID-keyed), or not Type 2
            raise ValueError("a CID-keyed or non-Type 2 CFF")
        self.charstrings, _ = _index(table, top[17][0])
        self.subrs: list[bytes] = []
        if 18 in top:
            size, offset = top[18]
            private = _dict(table[offset:offset + size])
            if 19 in private:
                self.subrs, _ = _index(table, offset + private[19][0])

    def box(self, glyph: int, slant: float = 0.0) -> Box | None:
        """As :func:`glyf_box`, for charstring ``glyph``."""
        if glyph >= len(self.charstrings):
            return None
        bounds = _Bounds(slant)
        try:
            if not _Type2(self, bounds).run(self.charstrings[glyph]):
                return None
        except (IndexError, struct.error, ValueError, RecursionError):
            return None
        return () if bounds.box is None else tuple(bounds.box)


class _Type2:
    """A Type 2 charstring interpreter that only follows the pen (Technical Note 5177)."""

    def __init__(self, cff: CFF, bounds: _Bounds) -> None:
        self.cff, self.bounds = cff, bounds
        self.stack: list[float] = []
        self.stems = 0
        self.width_done = False
        self.x = self.y = 0.0
        self.open = False

    def _width(self, even: bool) -> None:
        # The first stack-clearing operator may carry the advance width before its operands.
        if not self.width_done:
            self.width_done = True
            if (len(self.stack) % 2 == 1) == even:
                self.stack.pop(0)

    def _move(self, dx: float, dy: float) -> None:
        self.x += dx
        self.y += dy
        self.open = True
        self.bounds.point(self.x, self.y)

    def _line(self, dx: float, dy: float) -> None:
        self.x += dx
        self.y += dy
        self.bounds.point(self.x, self.y)

    def _curve(self, a, b, c, d, e, f) -> None:
        p0 = (self.x, self.y)
        p1 = (p0[0] + a, p0[1] + b)
        p2 = (p1[0] + c, p1[1] + d)
        p3 = (p2[0] + e, p2[1] + f)
        self.bounds.cubic(p0, p1, p2, p3)
        self.x, self.y = p3

    def run(self, code: bytes, depth: int = 0) -> bool | None:
        """Follow ``code``; ``True`` at ``endchar``, ``False`` where it cannot be followed,
        ``None`` at a subroutine's ``return``."""
        if depth > 10:
            return False
        s = self.stack
        at = 0
        while at < len(code):
            b0 = code[at]
            at += 1
            if b0 >= 32 or b0 == 28:
                if b0 == 28:
                    s.append(struct.unpack_from(">h", code, at)[0])
                    at += 2
                elif b0 <= 246:
                    s.append(b0 - 139)
                elif b0 <= 250:
                    s.append((b0 - 247) * 256 + code[at] + 108)
                    at += 1
                elif b0 <= 254:
                    s.append(-(b0 - 251) * 256 - code[at] - 108)
                    at += 1
                else:
                    s.append(struct.unpack_from(">i", code, at)[0] / 65536)
                    at += 4
                continue
            if b0 in (1, 3, 18, 23):  # hstem vstem hstemhm vstemhm
                self._width(even=True)
                self.stems += len(s) // 2
                s.clear()
            elif b0 in (19, 20):  # hintmask cntrmask: implicit vstems first
                self._width(even=True)
                self.stems += len(s) // 2
                s.clear()
                at += (self.stems + 7) // 8
            elif b0 == 21:  # rmoveto
                self._width(even=True)
                self._move(s[-2], s[-1])
                s.clear()
            elif b0 in (22, 4):  # hmoveto vmoveto
                self._width(even=False)
                self._move(s[-1], 0) if b0 == 22 else self._move(0, s[-1])
                s.clear()
            elif b0 == 5:  # rlineto
                for i in range(0, len(s) - 1, 2):
                    self._line(s[i], s[i + 1])
                s.clear()
            elif b0 in (6, 7):  # hlineto vlineto: alternating
                horizontal = b0 == 6
                for value in s:
                    self._line(value, 0) if horizontal else self._line(0, value)
                    horizontal = not horizontal
                s.clear()
            elif b0 == 8:  # rrcurveto
                for i in range(0, len(s) - 5, 6):
                    self._curve(*s[i:i + 6])
                s.clear()
            elif b0 == 24:  # rcurveline
                count = (len(s) - 2) // 6
                for i in range(count):
                    self._curve(*s[6 * i:6 * i + 6])
                self._line(s[-2], s[-1])
                s.clear()
            elif b0 == 25:  # rlinecurve
                lines = len(s) - 6
                for i in range(0, lines, 2):
                    self._line(s[i], s[i + 1])
                self._curve(*s[lines:lines + 6])
                s.clear()
            elif b0 in (26, 27):  # vvcurveto hhcurveto
                args = list(s)
                first = args.pop(0) if len(args) % 4 == 1 else 0
                for i in range(0, len(args) - 3, 4):
                    a, b, c, d = args[i:i + 4]
                    if b0 == 26:
                        self._curve(first, a, b, c, 0, d)
                    else:
                        self._curve(a, first, b, c, d, 0)
                    first = 0
                s.clear()
            elif b0 in (30, 31):  # vhcurveto hvcurveto: alternating
                args = list(s)
                horizontal = b0 == 31
                while len(args) >= 4:
                    a, b, c, d = args[:4]
                    args = args[4:]
                    last = args.pop(0) if len(args) == 1 else 0
                    if horizontal:
                        self._curve(a, 0, b, c, last, d)
                    else:
                        self._curve(0, a, b, c, d, last)
                    horizontal = not horizontal
                s.clear()
            elif b0 in (10, 29):  # callsubr callgsubr
                subrs = self.cff.subrs if b0 == 10 else self.cff.global_subrs
                index = int(s.pop()) + _bias(len(subrs))
                done = self.run(subrs[index], depth + 1)
                if done is not None:
                    return done
            elif b0 == 11:  # return
                return None
            elif b0 == 14:  # endchar
                self._width(even=True)
                return not s  # four operands left: a seac accent, not followed
            elif b0 == 12:
                b1 = code[at]
                at += 1
                if b1 == 35:  # flex
                    self._curve(*s[0:6])
                    self._curve(*s[6:12])
                elif b1 == 34:  # hflex
                    dx1, dx2, dy2, dx3, dx4, dx5, dx6 = s[:7]
                    self._curve(dx1, 0, dx2, dy2, dx3, 0)
                    self._curve(dx4, 0, dx5, -dy2, dx6, 0)
                elif b1 == 36:  # hflex1
                    dx1, dy1, dx2, dy2, dx3, dx4, dx5, dy5, dx6 = s[:9]
                    self._curve(dx1, dy1, dx2, dy2, dx3, 0)
                    self._curve(dx4, 0, dx5, dy5, dx6, -(dy1 + dy2 + dy5))
                elif b1 == 37:  # flex1
                    f = s[:11]
                    dx = sum(f[0:10:2])
                    dy = sum(f[1:10:2])
                    self._curve(*f[0:6])
                    if abs(dx) > abs(dy):
                        self._curve(f[6], f[7], f[8], f[9], f[10], -dy)
                    else:
                        self._curve(f[6], f[7], f[8], f[9], -dx, f[10])
                else:
                    return False  # Type 2 arithmetic and storage: not followed
                s.clear()
            else:
                return False
        return None


# -- GPOS --------------------------------------------------------------------------------------

def _coverage(table: bytes, at: int) -> dict[int, int]:
    fmt = struct.unpack_from(">H", table, at)[0]
    count = struct.unpack_from(">H", table, at + 2)[0]
    if fmt == 1:
        return {glyph: i for i, glyph in enumerate(struct.unpack_from(f">{count}H", table, at + 4))}
    out = {}
    for i in range(count):
        start, end, index = struct.unpack_from(">HHH", table, at + 4 + 6 * i)
        for glyph in range(start, end + 1):
            out[glyph] = index + glyph - start
    return out


def _class_def(table: bytes, at: int) -> dict[int, int]:
    fmt = struct.unpack_from(">H", table, at)[0]
    if fmt == 1:
        first, count = struct.unpack_from(">HH", table, at + 2)
        return {first + i: c for i, c in enumerate(struct.unpack_from(f">{count}H", table, at + 6)) if c}
    count = struct.unpack_from(">H", table, at + 2)[0]
    out = {}
    for i in range(count):
        start, end, c = struct.unpack_from(">HHH", table, at + 4 + 6 * i)
        if c:
            for glyph in range(start, end + 1):
                out[glyph] = c
    return out


def _value_size(fmt: int) -> int:
    return 2 * bin(fmt & 0xFF).count("1")


def _x_advance(table: bytes, at: int, fmt: int) -> int:
    """A ValueRecord's XAdvance (0 where it states none)."""
    if not fmt & 4:
        return 0
    return struct.unpack_from(">h", table, at + 2 * bin(fmt & 3).count("1"))[0]


class PairKerning:
    """``GPOS``'s ``kern`` feature, pair adjustments only: the first glyph's advance change."""

    def __init__(self, table: bytes) -> None:
        self.table = table
        self.subtables: list[list] = []  # per lookup: its PairPos subtables' offsets
        self._cache: dict = {}
        features_at, lookups_at = struct.unpack_from(">HH", table, 6)
        count = struct.unpack_from(">H", table, features_at)[0]
        indices: list[int] = []
        for i in range(count):
            tag, offset = struct.unpack_from(">4sH", table, features_at + 2 + 6 * i)
            if tag == b"kern":
                at = features_at + offset
                n = struct.unpack_from(">H", table, at + 2)[0]
                for index in struct.unpack_from(f">{n}H", table, at + 4):
                    if index not in indices:
                        indices.append(index)
        lookup_count = struct.unpack_from(">H", table, lookups_at)[0]
        for index in sorted(indices):
            if index >= lookup_count:
                continue
            at = lookups_at + struct.unpack_from(">H", table, lookups_at + 2 + 2 * index)[0]
            kind, _, n = struct.unpack_from(">HHH", table, at)
            found = []
            for j in range(n):
                sub = at + struct.unpack_from(">H", table, at + 6 + 2 * j)[0]
                sub_kind = kind
                if kind == 9:  # Extension
                    sub_kind = struct.unpack_from(">H", table, sub + 2)[0]
                    sub += struct.unpack_from(">I", table, sub + 4)[0]
                if sub_kind == 2:
                    found.append(sub)
            if found:
                self.subtables.append(found)

    def _parsed(self, at: int):
        found = self._cache.get(at)
        if found is None:
            t = self.table
            fmt, coverage, vf1, vf2 = struct.unpack_from(">HHHH", t, at)
            if fmt == 1:
                found = (1, _coverage(t, at + coverage), vf1, vf2)
            else:
                cd1, cd2, c1, c2 = struct.unpack_from(">HHHH", t, at + 8)
                found = (2, _coverage(t, at + coverage), vf1, vf2, _class_def(t, at + cd1), _class_def(t, at + cd2),
                         c1, c2)
            self._cache[at] = found
        return found

    def pair(self, left: int, right: int) -> int:
        t = self.table
        total = 0
        for lookup in self.subtables:
            for at in lookup:
                parsed = self._parsed(at)
                index = parsed[1].get(left)
                if index is None:
                    continue
                vf1, vf2 = parsed[2], parsed[3]
                record = _value_size(vf1) + _value_size(vf2)
                if parsed[0] == 1:
                    sets = at + struct.unpack_from(">H", t, at + 10 + 2 * index)[0]
                    n = struct.unpack_from(">H", t, sets)[0]
                    lo, hi = 0, n - 1
                    hit = None
                    while lo <= hi:
                        mid = (lo + hi) // 2
                        second = struct.unpack_from(">H", t, sets + 2 + mid * (2 + record))[0]
                        if second == right:
                            hit = sets + 2 + mid * (2 + record) + 2
                            break
                        lo, hi = (mid + 1, hi) if second < right else (lo, mid - 1)
                    if hit is None:
                        continue  # format 1: the next subtable may hold the pair
                    total += _x_advance(t, hit, vf1)
                else:
                    c1 = parsed[4].get(left, 0)
                    c2 = parsed[5].get(right, 0)
                    if c1 < parsed[6] and c2 < parsed[7]:
                        total += _x_advance(t, at + 16 + (c1 * parsed[7] + c2) * record, vf1)
                break  # this lookup has applied
        return total
