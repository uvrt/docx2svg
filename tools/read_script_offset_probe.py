#!/usr/bin/env python3
"""Measure ``make_script_offset_probe.py``: how far Word moved every superscript and
subscript, and at what size it drew them; and score ``resolve.script_raise_half_points``
on it and on the real faces of ``make_script_probe.py``'s size sweep.

Every line of the probe is ``Hx``, a superscript, `` Hx ``, a subscript, `` Hx``; each
is a run of its own, and Quartz draws each on a baseline of its own, so the offsets are
the distances between the text objects' baselines (whole device px) and the drawn size
is the script object's ``Tm`` scale.  The half points a script is moved by are recovered
from its offset (``round(k x 25/12)`` px is one-to-one on the values that occur).

``--record`` writes ``tests/fixtures/script-offset-observations.json``: per line the
case, the two offsets and the two drawn sizes; and the ``OS/2`` and ``hhea`` numbers of
the size sweep's real faces, read from the installed faces (numbers only).

Usage::

    python tools/read_script_offset_probe.py [--record]
"""

from __future__ import annotations

import json
import math
import re
import sys
from collections import Counter
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_script_offset_probe as probe  # noqa: E402

from docx2svg.resolve import script_half_points, script_raise_half_points  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

REPO = HERE.parent
OBSERVATIONS = REPO / "tests" / "fixtures" / "script-offset-observations.json"
SCRIPT_OBSERVATIONS = REPO / "tests" / "fixtures" / "script-observations.json"


def half_points(px: int) -> int:
    """The whole half points a drawn offset of ``px`` device px is."""
    return math.floor(Fraction(px * 12, 25) + Fraction(1, 2))


def measure() -> list:
    """``[kind, face, w:sz, superscript raise px, subscript drop px, sizes drawn px]``."""
    import oracle
    import quartz_pdf

    pages = quartz_pdf.read(oracle.export(probe.build(), name="script-offset"))
    runs = [run for page in pages for run in page.runs]
    groups: list[list] = []
    for run in runs:
        if abs(run.x - 300) < 1e-6 and run.text.startswith("Hx"):
            groups.append([])
        groups[-1].append(run)
    cases = probe.cases()
    assert len(groups) == len(cases), (len(groups), len(cases))
    out = []
    for (kind, face, hp), group in zip(cases, groups):
        host = group[0].y
        sup, sub = [run for run in group if run.text == probe.SCRIPT][:2]
        out.append([kind, face, hp, round(host - sup.y), round(sub.y - host), sup.size_px, sub.size_px])
    return out


def face_metrics(face: probe.Face) -> FaceMetrics:
    return FaceMetrics(probe.UPM, face.ascent, face.descent, 0, face.script_size, face.script_size,
                       face.sup_offset, face.sub_offset)


def score_probe(rows: list) -> dict:
    """``{"size", "superscript", "subscript": [exact, within a half point, of]}`` and the
    same for the refuted rule (the face's offset alone)."""
    faces = {face.name: face for face in probe.FACES}
    out = {"size": [0, 0], "superscript": [0, 0, 0], "subscript": [0, 0, 0],
           "offset alone": [0, 0, 0]}
    for kind, name, hp, sup_px, sub_px, sup_size, _sub_size in rows:
        metrics = face_metrics(faces[name])
        drawn = script_half_points(hp, "superscript", metrics)
        out["size"][0] += math.floor(Fraction(drawn * 25, 12) + Fraction(1, 2)) == round(sup_size)
        out["size"][1] += 1
        for key, align, px in (("superscript", "superscript", sup_px), ("subscript", "subscript", sub_px)):
            error = script_raise_half_points(hp, align, metrics) - half_points(px)
            out[key][0] += error == 0
            out[key][1] += abs(error) <= 1
            out[key][2] += 1
        alone = math.floor(Fraction(hp * faces[name].sup_offset, probe.UPM) + Fraction(1, 2))
        out["offset alone"][0] += alone == half_points(sup_px)
        out["offset alone"][2] += 1
    return out


def real_faces() -> dict:
    """The size sweep's faces' numbers, from the installed faces."""
    from docx2svg.fonts import InstalledFonts

    fonts = InstalledFonts()
    out = {}
    for name in ("Calibri", "Aptos", "Times New Roman", "Arial", "Cambria", "Courier New", "Georgia",
                 "Helvetica Neue", "Baskerville Old Face", "Impact", "Galvji", "Charter"):
        m = fonts.metrics(name)
        out[name] = [m.units_per_em, m.ascent, m.descent, m.line_gap, m.superscript_size, m.subscript_size,
                     m.superscript_offset, m.subscript_offset]
    return out


def score_real(faces: dict) -> dict:
    """The rule on ``make_script_probe.py``'s size sweep (``script-observations.json``)."""
    data = json.loads(SCRIPT_OBSERVATIONS.read_text(encoding="utf-8"))["scripts"]["script-sizes"]
    names = {name.replace(" ", ""): name for name in faces}
    out = {"superscript": [0, 0, 0], "subscript": [0, 0, 0]}
    per_face: Counter = Counter()
    for key, runs in data.items():
        match = re.match(r"size/(.*?)(\d+)$", key)
        name, hp = names[match.group(1)], int(match.group(2))
        metrics = FaceMetrics(*faces[name])
        sup = [dy for dy, _ in runs if dy < 0]
        sub = [dy for dy, _ in runs if dy > 0]
        for label, align, observed in (("superscript", "superscript", half_points(-sup[0]) if sup else 0),
                                       ("subscript", "subscript", half_points(sub[0]) if sub else 0)):
            error = script_raise_half_points(hp, align, metrics) - observed
            out[label][0] += error == 0
            out[label][1] += abs(error) <= 1
            out[label][2] += 1
            if label == "superscript" and error == 0:
                per_face[name] += 1
    out["superscript exact by face"] = dict(sorted(per_face.items()))
    return out


def main(argv: list[str]) -> int:
    rows = measure()
    faces = real_faces()
    print("probe:", json.dumps(score_probe([r for r in rows if r[0] == "size"])))
    print("beside a taller run, and a script larger than its text:")
    for row in rows:
        if row[0] != "size":
            print("   ", row)
    print("real faces:", json.dumps(score_real(faces)))
    if "--record" in argv[1:]:
        payload = {
            "_about": ("Every line of tools/make_script_offset_probe.py as Word 16.106 drew it: [case, face, "
                       "w:sz, superscript raised by px, subscript lowered by px, superscript and subscript "
                       "drawn size px]; and the OS/2 and hhea numbers of make_script_probe.py's size-sweep "
                       "faces ([upm, ascent, descent, lineGap, superscript and subscript size, superscript "
                       "and subscript offset]). Measurements only; regenerate with "
                       "tools/read_script_offset_probe.py --record."),
            "lines": rows,
            "real_faces": faces,
        }
        OBSERVATIONS.write_text(json.dumps(payload, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
