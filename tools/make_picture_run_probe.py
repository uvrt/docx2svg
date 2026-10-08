#!/usr/bin/env python3
"""Whether the run holding an inline picture counts its own text height in the line.

``make_picture_place_probe.py`` case 54 (mode 15, a 49 px picture under a 20 pt mark)
drew the picture, and the line after it, 6.78 px lower than the model: the line was 56 px,
Calibri 11 pt's natural height -- the size of the picture's run -- where the model had
the picture's 49 px (below mode 15 the mark's natural height floors the line, Phase 4.7).
Each case here is a page: an anchor line, a paragraph holding one picture shorter than
text (100,000, 150,000 and 200,000 EMU: 32.8, 49.2 and 65.6 px) in a run of ``w:sz`` 16,
22, 40 or 60, under a mark of ``w:sz`` 22 or 40, and a line after it, whose baseline says
how tall the picture's line was.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_picture_run_probe.py``.
"""

from __future__ import annotations

import io
import zipfile

import make_picture_probe as base
import probe_docx
import wml

FACE = base.FACE
SETTINGS = {"none": None, "15": 15}
CASES = tuple((height, run, mark) for height in (100000, 150000, 200000) for run in (16, 22, 40, 60)
              for mark in (22, 40))


def _p(text: str, **props) -> str:
    return wml.paragraph(wml.run(text), mark={}, **props)


def body() -> str:
    out = ""
    for number, (height, run, mark) in enumerate(CASES):
        out += _p(f"Case {number} anchor", pageBreakBefore=True)
        drawing = base.picture(number + 1, 1500000, height).replace(
            "<w:r>", f'<w:r><w:rPr><w:sz w:val="{run}"/><w:szCs w:val="{run}"/></w:rPr>', 1)
        mark_rpr = {"rFonts": FACE, "sz": mark, "szCs": mark}
        out += wml.paragraph(drawing, mark=mark_rpr)
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
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"picture-run-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
