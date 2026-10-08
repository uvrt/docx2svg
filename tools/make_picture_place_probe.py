#!/usr/bin/env python3
"""Where Word draws an inline picture in its line, and at what size.

``make_picture_probe.py`` measured how tall a picture's line is (Phase 4.7), read off the
baseline of the line after it; where the picture is *drawn* inside that line was never
read.  Phase 5.13's instrument showed ``samplelib/sample-long``'s full-width picture half a
pixel lower than Word's on every page it starts.  This probe reads Word's image boxes.

Every case is a page.  Pictures are a 1 x 1 PNG stretched to extents that are not whole
twips, so what Word does with the fraction shows.  Families (``CASES``):

* ``top`` -- the picture's paragraph starts the page (``w:pageBreakBefore``), as
  ``sample-long``'s do: extents stepped by 37,003 EMU down and 3,701 across, some with an
  effect extent (``wp:effectExtent`` top, left, bottom);
* ``exact`` -- after an anchor line, a picture under ``exact`` 200, 400, 600, 1000 and 3000
  twips, pictures shorter and taller than the line;
* ``atLeast`` -- ``atLeast`` 3000, taller than every picture;
* ``before`` -- space before of 120 and 137 twips;
* ``mark`` -- the paragraph mark at 20 pt, taller than the picture;
* ``pair`` -- two pictures of different heights on one line;
* ``text`` -- a word of text before the picture on its line.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_picture_place_probe.py``.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

import make_picture_probe as base
import probe_docx
import wml

FACE = base.FACE
SETTINGS = {"none": None, "15": 15}


@dataclass(frozen=True)
class Case:
    family: str
    heights: tuple[int, ...]  # EMU, one per picture on the line
    width: int = 1500000
    line: int = 240
    rule: str = "auto"
    effect: tuple[int, int, int, int] = (0, 0, 0, 0)  # l, t, r, b
    before: int = 0
    mark: int = 22
    top: bool = False
    text: bool = False


def _cases() -> tuple[Case, ...]:
    out = []
    effects = ((0, 0, 0, 0), (0, 0, 0, 6350), (12700, 6350, 0, 0), (0, 6350, 12700, 6350))
    for k in range(12):
        out.append(Case("top", (300000 + k * 37003,), width=1500000 + k * 3701, effect=effects[k % 4], top=True))
    for line in (200, 400, 600, 1000, 3000):
        for k in range(6):
            out.append(Case("exact", (300000 + k * 111007,), line=line, rule="exact", effect=effects[k % 2]))
    for k in range(6):
        out.append(Case("atLeast", (300000 + k * 111007,), line=3000, rule="atLeast", effect=effects[k % 2]))
    for before in (120, 137):
        for k in range(3):
            out.append(Case("before", (300000 + k * 111007,), before=before))
    for k in range(6):
        out.append(Case("mark", (150000 + k * 29011,), mark=40, effect=effects[k % 2]))
    for k in range(4):
        out.append(Case("pair", (300000 + k * 111007, 200000 + k * 50003), effect=effects[k % 2]))
    for k in range(3):
        out.append(Case("text", (300000 + k * 111007,), text=True))
    return tuple(out)


CASES = _cases()


def picture(number: int, cx: int, cy: int, effect: tuple[int, int, int, int]) -> str:
    l, t, r, b = effect
    return base.picture(number, cx, cy).replace(
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>', f'<wp:effectExtent l="{l}" t="{t}" r="{r}" b="{b}"/>')


def _p(text: str, **props) -> str:
    props.setdefault("spacing", {"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
    return wml.paragraph(wml.run(text), mark={}, **props)


def body() -> str:
    out = ""
    serial = 1
    for number, case in enumerate(CASES):
        if not case.top:
            out += _p(f"Case {number} anchor", pageBreakBefore=True)
        runs = wml.run(f"Case{number} ") if case.text else ""
        for height in case.heights:
            runs += picture(serial, case.width, height, case.effect)
            serial += 1
        mark = {"rFonts": FACE, "sz": case.mark, "szCs": case.mark}
        props = {"spacing": {"before": case.before, "after": 0, "line": case.line, "lineRule": case.rule}}
        if case.top:
            props["pageBreakBefore"] = True
        out += wml.paragraph(runs, mark=mark, **props)
        out += _p(f"Case {number} after")
    return out


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode = SETTINGS[setting]
    extra = [("word/media/image1.png", "image/png", base.IMAGE_REL, "")]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    data = probe_docx.package(body(), styles=styles, extra_parts=tuple(extra))
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in source.infolist():
            content = base.png() if info.filename == "word/media/image1.png" else source.read(info.filename)
            archive.writestr(info, content)
    return buffer.getvalue()


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"picture-place-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
