"""An EMF or WMF picture is drawn from its GDI records (``ooxml_common.metafile``), not as a
placeholder.  The metafiles are built here from MS-EMF's record layouts."""

from __future__ import annotations

import base64
import struct

from docx2svg.svg import _picture


def _record(kind: int, payload: bytes = b"") -> bytes:
    payload += b"\0" * ((-len(payload)) % 4)
    return struct.pack("<II", kind, 8 + len(payload)) + payload


def _emf(records: list[bytes]) -> bytes:
    eof = struct.pack("<5I", 14, 20, 0, 16, 20)
    body = b"".join(records) + eof
    header = struct.pack("<II4i4iIIIIHHIIIiiii", 1, 88, 0, 0, 99, 99, 0, 0, 2646, 2646, 0x464D4520, 0x10000,
                         88 + len(body), len(records) + 2, 2, 0, 0, 0, 0, 1920, 1080, 508, 286)
    return header + body


RED_SQUARE = [
    _record(39, struct.pack("<4I", 1, 0, 0x0000FF, 0)),
    _record(37, struct.pack("<I", 1)),
    _record(37, struct.pack("<I", 0x80000008)),
    _record(43, struct.pack("<4i", 10, 10, 90, 90)),
]


def _draw(data: bytes):
    warnings = []
    markup = _picture("p", "rId1", 0, 0, 100, 100, None, lambda rel: ("image/x-emf", data),
                      lambda code, message: warnings.append((code, message)))
    return markup, warnings


def test_an_emf_picture_is_drawn_from_its_records():
    markup, warnings = _draw(_emf(RED_SQUARE))
    assert "data:image/svg+xml;base64," in markup and "data-docx-unsupported" not in markup
    svg = base64.b64decode(markup.split("base64,", 1)[1].split('"', 1)[0])
    assert b'fill="#FF0000"' in svg
    assert warnings == []


def test_records_left_out_are_reported():
    markup, warnings = _draw(_emf(RED_SQUARE + [_record(53, bytes(16))]))
    assert "data:image/svg+xml" in markup
    assert warnings and warnings[0][0] == "metafile-records" and "EMR_EXTFLOODFILL" in warnings[0][1]


def test_a_metafile_that_draws_nothing_keeps_the_placeholder():
    markup, warnings = _draw(_emf([]))
    assert "data-docx-unsupported" in markup
    assert warnings[0][0] == "picture-not-drawn"
