"""Probes scored as whole documents, through the same path as the real ones.

``make_line_box_probe.py`` and its siblings rebuild each page's stack in their reader.
The probes that decide rules of the *page stack* -- what a paragraph at the top of a
page keeps, what the cascade makes of autospacing, what a list label adds to its line
-- cannot do that without restating the rule they test.  So these readers hand the
generated ``.docx`` and Word's drawn lines to ``baselines.predict``, which reads the
document with ``docx2svg`` and applies the model: the probe tests the model's code, not
a copy of it.

A recording holds, per document, the drawn lines as ``[page, baseline, text]`` (the
text is the probe's own) and the four metric integers of every face the model asked
for.  The documents themselves are regenerated from their generator, byte for byte, so
a test needs neither Word nor a font file.
"""

from __future__ import annotations

import functools
import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import baselines  # noqa: E402

from docx2svg import parse_package  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

FIXTURES = HERE.parent / "tests" / "fixtures"


def measure(name: str, data: bytes) -> list[baselines.DrawnLine]:
    """Word's lines of ``data`` (exported through ``oracle.py``, cached by content)."""
    import oracle
    import quartz_pdf

    return baselines.drawn_lines(quartz_pdf.read(oracle.export(data, name=name)))


class RecordingMetrics:
    """``face_metrics.metrics``, remembering the integers of every face asked for."""

    def __init__(self):
        self.faces: dict[str, list[int] | None] = {}

    def __call__(self, face, bold=False, italic=False):
        import face_metrics

        found = face_metrics.metrics(face, bold, italic)
        self.faces[f"{face}|{int(bold)}|{int(italic)}"] = (
            None if found is None else baselines.face_integers(found))
        return found


def recorded_metrics(faces: dict):
    def metrics(face, bold=False, italic=False):
        found = faces.get(f"{face}|{int(bold)}|{int(italic)}")
        return None if found is None else FaceMetrics(*found)
    return metrics


@functools.lru_cache(maxsize=8)
def _parsed(data: bytes):
    """A generated document, parsed once: a test that scores one probe under several
    rules would otherwise parse 400 pages each time (``predict`` changes nothing in it)."""
    return parse_package(data)


def score(data: bytes, drawn, metrics) -> list[baselines.LineResult]:
    return baselines.predict(_parsed(data), drawn, metrics, package=data)


def compact(drawn) -> list:
    return [[line.page, int(line.y) if line.y == int(line.y) else line.y, line.text] for line in drawn]


def expand(lines) -> list[baselines.DrawnLine]:
    return [baselines.DrawnLine(page, y, (("", 0, text),)) for page, y, text in lines]


def write(path: Path, about: str, faces: dict, documents: dict) -> None:
    payload = {"_about": about, "faces": dict(sorted(faces.items())), "documents": documents}
    path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
    print(f"wrote {path} ({path.stat().st_size} bytes)")


def tally(results, kinds: list[str | None], key) -> dict:
    """``{key(kind of the page's first paragraph): [exact, scored]}`` over every line.

    ``kinds`` is the generator's tag for every block, in ``baselines.blocks`` order.  A
    page is classed by the paragraph its first line belongs to, and every line of the
    page counts toward that class: a page-top error moves the whole page.
    """
    out: dict = defaultdict(lambda: [0, 0])
    page_kind: dict[int, str | None] = {}
    for r in results:
        if r.page not in page_kind and r.block >= 0:
            page_kind[r.page] = kinds[r.block]
        if r.status not in ("exact", "miss"):
            continue
        cell = out[key(page_kind.get(r.page))]
        cell[0] += r.status == "exact"
        cell[1] += 1
    return out
