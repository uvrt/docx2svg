"""Record what the renderer asks of its faces, and answer from the recording.

The renderer measures and draws with ``docx2svg.fonts.InstalledFonts``: advances, legacy
kern pairs, the four vertical metrics and the script sizes, and the decoration numbers
(``OS/2`` script offsets and strikeout, ``post`` underline).  A test, a VRT snapshot or a
CI machine without Word's faces must get the same answers, so every answer a render
needs is recorded -- numbers only, never a font file -- and replayed.

A recording is ``{"face|bold|italic": {"upm", "advances": {char: width}, "kern": {pair:
value}, "metrics": [...], "decorations": [...], "drawing": [family, weight, stretch,
style]}}`` -- ``drawing`` only for a face the SVG names by a fallback list
(``InstalledFonts.drawing_name``).  An advance not recorded is unknown
(``None``, which the layout reports as unmeasurable); a kern pair not recorded is zero.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))

from docx2svg.fonts import Decorations, DrawingName, InstalledFonts  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402


def _key(face: str, bold: bool, italic: bool) -> str:
    return f"{face}|{int(bool(bold))}|{int(bool(italic))}"


class RecordingFonts:
    """``InstalledFonts`` (the document's embedded faces included), remembering every
    answer in ``faces``."""

    def __init__(self, package: bytes | None, faces: dict | None = None) -> None:
        self.inner = InstalledFonts(package)
        self.faces: dict = faces if faces is not None else {}

    def _entry(self, face, bold, italic) -> dict:
        return self.faces.setdefault(_key(face, bold, italic), {"upm": None, "advances": {}, "kern": {}})

    def face(self, family, bold=False, italic=False):
        return self.inner.face(family, bold, italic)

    def advance(self, face, bold, italic, char):
        found = self.inner.advance(face, bold, italic, char)
        entry = self._entry(face, bold, italic)
        if found is not None:
            entry["upm"] = found[1]
            entry["advances"][char] = found[0]
        return found

    def kern(self, face, bold, italic, left, right):
        value = self.inner.kern(face, bold, italic, left, right)
        if value:
            self._entry(face, bold, italic)["kern"][left + right] = value
        return value

    def metrics(self, face, bold=False, italic=False):
        found = self.inner.metrics(face, bold, italic)
        entry = self._entry(face, bold, italic)
        if found is not None:
            entry["metrics"] = [found.units_per_em, found.ascent, found.descent, found.line_gap,
                                found.superscript_size, found.subscript_size, found.superscript_offset,
                                found.subscript_offset]
        else:
            entry["metrics"] = None
        return found

    def decorations(self, face, bold=False, italic=False):
        found = self.inner.decorations(face, bold, italic)
        entry = self._entry(face, bold, italic)
        entry["decorations"] = None if found is None else found.integers()
        return found

    def drawing_name(self, face, bold=False, italic=False):
        found = self.inner.drawing_name(face, bold, italic)
        if found is not None:
            self._entry(face, bold, italic)["drawing"] = [found.family, found.weight, found.stretch, found.style]
        return found


class RecordedFonts:
    """Answers from a recording, exactly as :class:`RecordingFonts` answered."""

    def __init__(self, faces: dict) -> None:
        self.faces = faces

    def _entry(self, face, bold, italic):
        return self.faces.get(_key(face, bold, italic))

    def face(self, family, bold=False, italic=False):
        return None

    def advance(self, face, bold, italic, char):
        entry = self._entry(face, bold, italic)
        if entry is None or char not in entry["advances"]:
            return None
        return entry["advances"][char], entry["upm"]

    def kern(self, face, bold, italic, left, right):
        entry = self._entry(face, bold, italic)
        if entry is None:
            return None
        return entry["kern"].get(left + right, 0)

    def metrics(self, face, bold=False, italic=False):
        entry = self._entry(face, bold, italic)
        if entry is None or not entry.get("metrics"):
            return None
        return FaceMetrics(*entry["metrics"])

    def decorations(self, face, bold=False, italic=False):
        entry = self._entry(face, bold, italic)
        if entry is None or not entry.get("decorations"):
            return None
        values = entry["decorations"]
        return Decorations(values[0], *values[1:7], hhea_ascent=values[7], hhea_descent=values[8])

    def drawing_name(self, face, bold=False, italic=False):
        entry = self._entry(face, bold, italic)
        if entry is None or not entry.get("drawing"):
            return None
        return DrawingName(*entry["drawing"])


def options(fonts, **kwargs):
    """``ConvertOptions`` measuring and drawing with ``fonts``."""
    from docx2svg import ConvertOptions

    return ConvertOptions(advances=fonts, metrics=fonts.metrics, decorations=fonts.decorations,
                          names=fonts.drawing_name, **kwargs)


def dump(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n",
                    encoding="utf-8")
    print(f"wrote {path} ({path.stat().st_size} bytes)")
