#!/usr/bin/env python3
"""Where Word draws a paragraph's bottom border, to the device pixel.

Phase 5.2 drew the bottom border ``w:space`` below the last line's text row and found 125
of 150 bottom borders over two probes exact, the rest a pixel off, with two other rules
tried each failing different cases.  Those probes stepped few sizes.  This one sweeps the
fraction of a pixel where a line's text ends: each case is a pair -- a one-line paragraph
with a bottom border (red), then a plain one -- in Calibri at every half point from 10 to
22 pt, under five borders (``w:sz`` / ``w:space``: 4/0, 6/2, 8/1, 12/4, 24/10) and with
no space after and 120 twips after; and, at 11-15 pt, under ``auto`` 360, ``exact`` 400
and ``atLeast`` 480.  The plain paragraph's baseline says where the border's room ended.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_para_border_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "15": 15}
BORDERS = ((4, 0), (6, 2), (8, 1), (12, 4), (24, 10))
RED = "FF0000"
PER_PAGE = 18


def _cases() -> list[dict]:
    out = []
    for hp in range(20, 45):
        for sz, space in BORDERS:
            for after in (0, 120):
                out.append({"hp": hp, "sz": sz, "space": space, "after": after, "rule": "auto", "line": 240})
    for hp in range(22, 31):
        for sz, space in BORDERS:
            for rule, line in (("auto", 360), ("exact", 400), ("atLeast", 480)):
                out.append({"hp": hp, "sz": sz, "space": space, "after": 0, "rule": rule, "line": line})
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        rpr = {"rFonts": FACE, "sz": case["hp"], "szCs": case["hp"]}
        border = f'<w:bottom w:val="single" w:sz="{case["sz"]}" w:space="{case["space"]}" w:color="{RED}"/>'
        props = {"pBdr": border, "spacing": {"before": 0, "after": case["after"], "line": case["line"],
                                             "lineRule": case["rule"]}}
        if number and number % PER_PAGE == 0:
            props["pageBreakBefore"] = True
        out += wml.paragraph(wml.run(f"Bordered {number} Hxg", **rpr), mark=rpr, **props)
        out += wml.paragraph(wml.run(f"Plain {number} Hxg"))
    return out


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"para-border-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
