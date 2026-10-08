#!/usr/bin/env python3
"""The advance widths the Phase 3 probes are built from, recorded once.

The line-breaking probes place a decisive character a known distance from the right
edge, which needs its width at generation time; and a probe must regenerate byte for
byte without a font file, so the widths are a committed recording,
``tests/fixtures/probe-advances.json``: for every face and style in :data:`FACES`, the
``hmtx`` advance of every character in :data:`CHARS`; for the faces in
:data:`KERN_FACES`, their legacy ``kern`` pairs among printable ASCII (what Word applies,
Phase 2) and, beside them, the OpenType ``kern`` feature's pairs (``gpos``), so a probe
can tell the two apart.  Read from the faces Word lays out with (``face_advances``).

Measurements only -- numbers, no font file.  ``python tools/probe_advances.py`` rewrites
the recording; ``--check`` compares it with the installed faces.
"""

from __future__ import annotations

import functools
import json
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
PATH = HERE.parent / "tests" / "fixtures" / "probe-advances.json"

FACES = (
    "Calibri", "Arial", "Times New Roman", "Cambria", "Georgia", "Verdana", "Aptos",
    "Courier New", "Helvetica Neue", "Century Gothic", "Trebuchet MS", "Tahoma", "Garamond",
    "Constantia",
)
STYLES = ((False, False), (True, False), (False, True), (True, True))
CHARS = "".join(
    [chr(c) for c in range(0x20, 0x7F)] + [chr(c) for c in range(0xA0, 0x100)]
    + [chr(c) for c in range(0x2002, 0x200C)] + [chr(c) for c in range(0x2010, 0x2028)]
    + ["\u202f", "\u2060", "−", "€", "™"]
)
KERN_FACES = ("Calibri", "Times New Roman", "Arial", "Aptos", "Georgia", "Cambria")
KERN_CHARS = "".join(chr(c) for c in range(0x20, 0x7F))


def key(face: str, bold: bool, italic: bool) -> str:
    return f"{face}|{int(bold)}|{int(italic)}"


def record() -> dict:
    sys.path.insert(0, str(HERE))
    sys.path.insert(0, str(HERE.parent / "src"))
    import face_advances
    from fontTools.ttLib import TTFont
    import face_metrics
    from ooxml_common.text.measure import _font_kern_units

    out = {}
    for face in FACES:
        for bold, italic in STYLES:
            found = face_advances.advances(face, bold, italic)
            if found is None:
                continue
            entry = {"upm": found.units_per_em,
                     "advances": {c: found.advance(c) for c in CHARS if found.advance(c) is not None}}
            if face in KERN_FACES and (bold, italic) == (False, False):
                entry["kern"] = {a + b: found.kern(a, b) for a in KERN_CHARS for b in KERN_CHARS
                                 if found.kern(a, b)}
                index = face_metrics._index()[(face.lower(), False, False)]
                font = TTFont(index[0], fontNumber=index[1], lazy=True)
                entry["gpos"] = {a + b: v for a in KERN_CHARS for b in KERN_CHARS
                                 if (v := _font_kern_units(font, a, b))}
            out[key(face, bold, italic)] = entry
    return out


@functools.lru_cache(maxsize=1)
def load() -> dict:
    return json.loads(PATH.read_text(encoding="utf-8"))["faces"]


class Advances:
    """:class:`docx2svg.measure.Advances` over the recording (legacy ``kern`` pairs, or
    the OpenType feature's with ``gpos=True``)."""

    def __init__(self, gpos: bool = False) -> None:
        self.faces = load()
        self.table = "gpos" if gpos else "kern"

    def advance(self, face, bold, italic, char):
        entry = self.faces.get(key(face, bold, italic))
        if entry is None or char not in entry["advances"]:
            return None
        return entry["advances"][char], entry["upm"]

    def kern(self, face, bold, italic, left, right):
        entry = self.faces.get(key(face, bold, italic))
        if entry is None or self.table not in entry:
            return None
        return entry[self.table].get(left + right, 0)


def units(face: str, text: str, half_points: int, *, bold: bool = False, italic: bool = False,
          kern: str | None = None) -> Fraction:
    """``text``'s width in 1/4096 pt (``kern``: ``"kern"`` or ``"gpos"`` pairs)."""
    entry = load()[key(face, bold, italic)]
    total = sum(Fraction(entry["advances"][c] * half_points * 2048, entry["upm"]) for c in text)
    if kern:
        total += sum(Fraction(entry[kern].get(a + b, 0) * half_points * 2048, entry["upm"])
                     for a, b in zip(text, text[1:]))
    return total


def main(argv: list[str]) -> int:
    found = record()
    if "--check" in argv[1:]:
        recorded = load()
        bad = [k for k in set(found) | set(recorded) if found.get(k) != recorded.get(k)]
        print("recording matches the installed faces" if not bad else f"differs: {sorted(bad)}")
        return 1 if bad else 0
    PATH.write_text(json.dumps({
        "_about": ("hmtx advances (font units, upm per face) of the characters the Phase 3 probes use, "
                   "and legacy kern / OpenType kern-feature pairs among printable ASCII for a few "
                   "faces, read from the faces Word 16.106 lays out with on macOS. Measurements "
                   "only; regenerate with tools/probe_advances.py."),
        "faces": found,
    }, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
    print(f"wrote {PATH} ({PATH.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
