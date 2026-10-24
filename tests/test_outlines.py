"""Glyph outline boxes and GPOS pair kerning (docx2svg.outlines), from made-up tables built
here byte by byte: the structures, not any face's numbers."""

import struct

import pytest

from docx2svg import outlines


# -- glyf --------------------------------------------------------------------------------------

def _simple(points: list[tuple[int, int, bool]], ends: list[int]) -> bytes:
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    out = struct.pack(">hhhhh", len(ends), min(xs), min(ys), max(xs), max(ys))
    out += struct.pack(f">{len(ends)}H", *ends) + struct.pack(">H", 0)  # no instructions
    out += bytes(1 if on else 0 for _, _, on in points)                 # long coordinates, no repeats
    x = y = 0
    for px, _, _ in points:
        out += struct.pack(">h", px - x)
        x = px
    for _, py, _ in points:
        out += struct.pack(">h", py - y)
        y = py
    return out + b"\0" * (-len(out) % 4)


def _glyf(glyphs: list[bytes]) -> tuple[bytes, bytes]:
    data, offsets = b"", [0]
    for glyph in glyphs:
        data += glyph
        offsets.append(len(data))
    return data, struct.pack(f">{len(offsets)}H", *(o // 2 for o in offsets))


#: One contour: a quadratic arch from (0, 0) over (100, 200) to (200, 0), closed by a line.
ARCH = _simple([(0, 0, True), (100, 200, False), (200, 0, True)], [2])


def test_a_quadratic_outline_s_box_is_its_curve_s_extremes_not_its_control_points():
    glyf, loca = _glyf([ARCH])
    assert outlines.glyf_box(glyf, loca, False, 0) == (0, 0, 200, 100)


def test_a_slanted_outline_s_box_is_the_slanted_curve_s():
    glyf, loca = _glyf([ARCH])
    x_min, y_min, x_max, y_max = outlines.glyf_box(glyf, loca, False, 0, slant=1.0)
    # x + y along the arch is 600 t - 400 t^2: past the end point's 200, at t = 0.75.
    assert (x_min, y_min, y_max) == (0, 0, 100) and x_max == pytest.approx(225)


def test_a_composite_glyph_s_box_is_its_components_placed_and_scaled():
    component = struct.pack(">hhhhh", -1, 0, 0, 0, 0) + struct.pack(">HHhhh", 0x0001 | 0x0002 | 0x0008, 0, 50, 10,
                                                                    8192)  # words, x/y, a scale of 0.5
    glyf, loca = _glyf([ARCH, component])
    assert outlines.glyf_box(glyf, loca, False, 1) == (50, 10, 150, 60)


def test_an_empty_glyph_has_an_empty_box_and_a_glyph_past_the_table_none():
    glyf, loca = _glyf([ARCH, b""])
    assert outlines.glyf_box(glyf, loca, False, 1) == ()
    assert outlines.glyf_box(glyf, loca, False, 5) is None


# -- CFF ---------------------------------------------------------------------------------------

def _index(items: list[bytes]) -> bytes:
    if not items:
        return struct.pack(">H", 0)
    offsets, at = [1], 1
    for item in items:
        at += len(item)
        offsets.append(at)
    return struct.pack(">HB", len(items), 4) + b"".join(struct.pack(">I", o) for o in offsets) + b"".join(items)


def _n(value: int) -> bytes:
    """A Type 2 charstring number (-107..107 in one byte, else ``28`` and a short)."""
    return bytes([value + 139]) if -107 <= value <= 107 else b"\x1c" + struct.pack(">h", value)


def _cff(charstrings: list[bytes], global_subrs: list[bytes] = ()) -> bytes:
    header = bytes([1, 0, 4, 4])
    names = _index([b"MadeUp"])
    strings = _index([])
    gsubrs = _index(list(global_subrs))

    def top(offset: int) -> bytes:
        return _index([b"\x1d" + struct.pack(">i", offset) + bytes([17])])

    before = len(header) + len(names) + len(top(0)) + len(strings) + len(gsubrs)
    return header + names + top(before) + strings + gsubrs + _index(charstrings)


#: A width (500, before the first move), a move to (10, 0), a line to (110, 0) and a cubic
#: back to (10, 0) through (110, 300) and (10, 300): its top is at 225, not 300.
CUBIC = (_n(500) + _n(10) + _n(0) + bytes([21]) + _n(100) + _n(0) + bytes([5]) +
         _n(0) + _n(300) + _n(-100) + _n(0) + _n(0) + _n(-300) + bytes([8, 14]))


def test_a_cff_charstring_s_box_is_its_curves_extremes():
    cff = outlines.CFF(_cff([bytes([14]), CUBIC]))
    assert cff.box(0) == ()
    assert cff.box(1) == pytest.approx((10, 0, 110, 225))


def test_a_cff_charstring_follows_its_global_subroutines():
    # The line and the curve in a global subroutine (bias 107: index 0 is called as -107).
    body = _n(100) + _n(0) + bytes([5]) + _n(0) + _n(300) + _n(-100) + _n(0) + _n(0) + _n(-300) + bytes([8, 11])
    call = _n(500) + _n(10) + _n(0) + bytes([21]) + _n(-107) + bytes([29, 14])
    assert outlines.CFF(_cff([bytes([14]), call], [body])).box(1) == pytest.approx((10, 0, 110, 225))


def test_a_cff_charstring_it_cannot_follow_has_no_box():
    seac = _n(0) + _n(0) + _n(0) + _n(65) + _n(66) + bytes([14])  # endchar as an accented character
    arithmetic = _n(1) + _n(2) + bytes([12, 10, 14])               # add
    cff = outlines.CFF(_cff([bytes([14]), seac, arithmetic]))
    assert cff.box(1) is None and cff.box(2) is None and cff.box(9) is None


# -- GPOS --------------------------------------------------------------------------------------

def _coverage(glyphs: list[int]) -> bytes:
    return struct.pack(f">HH{len(glyphs)}H", 1, len(glyphs), *glyphs)


def _pair_format1(left: int, right: int, advance: int) -> bytes:
    # PairPos 1: coverage, ValueFormat1 = XAdvance, ValueFormat2 = none, one PairSet.
    head = 12
    pair_set = struct.pack(">HHh", 1, right, advance)
    coverage = _coverage([left])
    return (struct.pack(">HHHHHH", 1, head + len(pair_set), 4, 0, 1, head) + pair_set + coverage)


def _pair_format2(lefts: list[int], rights: list[int], advance: int) -> bytes:
    # PairPos 2: two classes on each side; class 1 x class 1 kerns by ``advance``.
    head = 16
    records = struct.pack(">hhhh", 0, 0, 0, advance)
    coverage = _coverage(lefts)
    def class_def(glyphs):
        span = range(min(glyphs), max(glyphs) + 1)
        return struct.pack(f">HHH{len(span)}H", 1, span[0], len(span), *(int(g in glyphs) for g in span))

    class1, class2 = class_def(lefts), class_def(rights)
    at_cov = head + len(records)
    return (struct.pack(">HHHHHHHH", 2, at_cov, 4, 0, at_cov + len(coverage),
                        at_cov + len(coverage) + len(class1), 2, 2) + records + coverage + class1 + class2)


def _lookup(subtables: list[bytes], kind: int = 2) -> bytes:
    head = 6 + 2 * len(subtables)
    offsets, at = [], head
    for sub in subtables:
        offsets.append(at)
        at += len(sub)
    return struct.pack(f">HHH{len(subtables)}H", kind, 0, len(subtables), *offsets) + b"".join(subtables)


def _extension(sub: bytes) -> bytes:
    return struct.pack(">HHI", 1, 2, 8) + sub


def _gpos(features: list[tuple[bytes, list[int]]], lookups: list[bytes]) -> bytes:
    feature_tables = [struct.pack(f">HH{len(ix)}H", 0, len(ix), *ix) for _, ix in features]
    head = 2 + 6 * len(features)
    records, at = b"", head
    for (tag, _), table in zip(features, feature_tables):
        records += struct.pack(">4sH", tag, at)
        at += len(table)
    feature_list = struct.pack(">H", len(features)) + records + b"".join(feature_tables)
    lhead = 2 + 2 * len(lookups)
    offsets, at = [], lhead
    for lookup in lookups:
        offsets.append(at)
        at += len(lookup)
    lookup_list = struct.pack(f">H{len(lookups)}H", len(lookups), *offsets) + b"".join(lookups)
    script_list = struct.pack(">H", 0)
    start = 10
    return (struct.pack(">HHHHH", 1, 0, start, start + len(script_list), start + len(script_list) + len(feature_list))
            + script_list + feature_list + lookup_list)


def test_gpos_pair_kerning_reads_the_kern_feature_s_pair_lookups():
    kerning = outlines.PairKerning(_gpos(
        [(b"kern", [0, 1]), (b"mark", [2])],
        [_lookup([_pair_format1(5, 6, -50)]), _lookup([_extension(_pair_format2([5, 7], [6, 8], -30))], kind=9),
         _lookup([_pair_format1(5, 6, -999)])]))
    assert kerning.pair(5, 6) == -80      # both kern lookups add; the mark feature's is not kern
    assert kerning.pair(7, 8) == -30      # the class pair, through an extension lookup
    assert kerning.pair(6, 5) == 0 and kerning.pair(5, 9) == 0


def test_gpos_pair_kerning_stops_at_a_lookup_s_first_subtable_that_covers_the_pair():
    kerning = outlines.PairKerning(_gpos([(b"kern", [0])], [_lookup([
        _pair_format1(5, 9, -10),                 # covers 5 but not the pair: the next one is tried
        _pair_format2([5], [6], -40),             # covers it: applies
        _pair_format1(5, 6, -999)])]))            # never reached
    assert kerning.pair(5, 6) == -40 and kerning.pair(5, 9) == -10
