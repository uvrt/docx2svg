"""A folder scan skips a font file it cannot read and indexes the rest.

macOS ships a file with no ``head`` table (Supplemental/NISC18030.ttf); since
ooxml-common's ``Face`` raises its own ``FontError`` for that, the scan has to skip that
error as it skips a truncated table.  Both faces here are made up in the test.
"""

from __future__ import annotations

import struct

from docx2svg import fonts


def _name_table(family: str) -> bytes:
    """A format-0 ``name`` table with the family (name ID 1) and subfamily (ID 2)."""
    strings = [(1, family.encode("utf-16-be")), (2, "Regular".encode("utf-16-be"))]
    records, storage = b"", b""
    for name_id, text in strings:
        records += struct.pack(">6H", 3, 1, 0x409, name_id, len(text), len(storage))
        storage += text
    return struct.pack(">3H", 0, len(strings), 6 + len(records)) + records + storage


def _font(tables: dict[bytes, bytes]) -> bytes:
    """An sfnt holding ``tables`` (no checksums: nothing here reads them)."""
    count = len(tables)
    directory, body = b"", b""
    start = 12 + 16 * count
    for tag in sorted(tables):
        data = tables[tag]
        directory += struct.pack(">4sIII", tag, 0, start + len(body), len(data))
        body += data + b"\0" * (-len(data) % 4)
    return struct.pack(">IHHHH", 0x00010000, count, 0, 0, 0) + directory + body


def test_a_font_without_head_is_skipped(tmp_path):
    head = struct.pack(">IIIIHH", 0x00010000, 0, 0, 0x5F0F3CF5, 0, 1000) + b"\0" * 34
    (tmp_path / "a-broken.ttf").write_bytes(_font({b"name": _name_table("Scan Probe Broken")}))
    (tmp_path / "b-good.ttf").write_bytes(_font({b"name": _name_table("Scan Probe Good"), b"head": head}))
    index = fonts.installed_index((tmp_path,))
    assert index == {("scan probe good", False, False): (str(tmp_path / "b-good.ttf"), 0)}
