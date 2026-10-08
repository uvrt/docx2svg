"""What one character costs a line: advance widths and kern pairs, by face.

The line breaker (:mod:`docx2svg.linebreak`) asks two questions of a face -- the advance
of a character and the kern adjustment between two -- through the :class:`Advances`
protocol, in the face's own units with its ``unitsPerEm``.  Word turns them into its
layout unit, 1/4096 pt, without rounding (ROADMAP.md, "Phase 2 -- measured", 2.1): for a
2048-upm face at a half-point size an advance is exactly ``advance x half points`` units.

:class:`TableAdvances` answers from ``ooxml-common``'s generated tables -- measured for
PowerPoint, and trusted here only where they were checked against the copy of the face
Word lays out with (ROADMAP.md, "Phase 3 -- measured": which faces agree).  Everything
else is ``None``, which the breaker reports rather than guesses: a width from the wrong
face breaks a long line in the wrong place.  ``tools/face_advances.py`` answers from the
installed font files instead, and the tests answer from recordings of those.
"""

from __future__ import annotations

from typing import Protocol


class Advances(Protocol):
    def advance(self, face: str, bold: bool, italic: bool, char: str) -> tuple[int, int] | None:
        """``(advance, unitsPerEm)`` of ``char`` in the face Word lays ``face`` out with,
        or ``None`` when unknown."""

    def kern(self, face: str, bold: bool, italic: bool, left: str, right: str) -> int | None:
        """The ``kern`` adjustment between two characters, in the face's units; ``0``
        for a pair that does not kern, ``None`` when unknown."""


#: Word's face name -> the ``ooxml-common`` table whose advances are that face's, for
#: every character the table holds.  Checked against the installed faces Word lays out
#: with (``tools/check_advance_tables.py``): Calibri (Carlito's table), Times New Roman
#: (Tinos'), Courier New (Cousine's), Cambria and Aptos, upright and bold, all agree;
#: Arial (Arimo's) in every character but the two in :data:`_OVERRIDES`.  Calibri
#: Light has no table of its own (``ooxml-common`` stands Carlito in for it, 1.3%
#: wider), and no table holds an italic cut.
TABLES = {
    "calibri": "Carlito",
    "arial": "Arimo",
    "times new roman": "Tinos",
    "courier new": "Cousine",
    "cambria": "Cambria",
    "aptos": "Aptos",
}

#: ``(face, bold, char) -> advance`` where the installed face (Word's bundle and the
#: macOS copy agree) differs from the table: Arimo is a metric clone of Arial, not a copy.
_OVERRIDES = {
    ("arial", False, "ˆ"): 682,
    ("arial", True, "µ"): 1180,
}


class TableAdvances:
    """Advances from ``ooxml-common``'s tables, for the faces in :data:`TABLES`.

    Kerning is not answered (``None``): the tables' pairs are the OpenType ``kern``
    feature of the face -- or of its clone -- that PowerPoint applies, and Word applies
    the legacy ``kern`` table, which differs for Calibri, Cambria and Aptos.
    """

    def __init__(self) -> None:
        from ooxml_common.text.metrics import METRICS

        self._metrics = METRICS

    def advance(self, face: str, bold: bool, italic: bool, char: str) -> tuple[int, int] | None:
        key = face.lower()
        name = TABLES.get(key)
        if name is None or italic:
            return None
        override = _OVERRIDES.get((key, bold, char))
        table = self._metrics[name]
        if override is not None:
            return override, table.units_per_em
        widths = table.bold_widths if bold else table.widths
        width = widths.get(char)
        return None if width is None else (width, table.units_per_em)

    def kern(self, face: str, bold: bool, italic: bool, left: str, right: str) -> int | None:
        return None
