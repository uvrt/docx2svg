#!/usr/bin/env python3
"""What Word *draws* besides glyphs, and where: Phase 5's probe of decorations and
alignment.

Phases 2-4 measured where every line and glyph goes.  Drawing adds geometry nothing has
measured yet: where an underline or a strike sits and how thick it is, how far a
highlight or shading reaches, where a paragraph border is drawn and how far it runs,
and -- because the documents so far had none -- where the glyphs of a justified,
centred or right-aligned line, or of text after a right, centre or decimal tab, go.
One family per page (each led by a ``w:pageBreakBefore`` anchor), one variable at a
time:

* ``decor`` -- a run ``Hxample decorated`` between plain text, with ``w:u`` (single,
  double, thick, words, dotted, dash), ``w:strike``, ``w:dstrike``, ``w:highlight``
  (yellow) or run ``w:shd`` (a fill), in eight faces at five sizes;
* ``fill`` -- highlight and shading under ``auto`` 240 and 360, ``exact`` 400,
  ``atLeast`` 480, and with space before and after, at 11 and 20 pt;
* ``box`` -- paragraph borders: all four sides at ``w:sz`` 4, 12, 24 and ``w:space`` 0,
  4, 10, with and without indents; one side at a time; paragraph shading alone and
  with a box; two boxed paragraphs with a ``w:between`` border;
* ``align`` -- the wrapping probe's prose centred, right-aligned, justified (``both``)
  and distributed, in three faces, at three right indents; and single centred lines
  whose slack is odd and even in layout units;
* ``tabs`` -- right, centre and decimal stops, and a dot leader.

The text is this project's own.  Faces: Word's and macOS's installed ones, named, never
embedded.  One document per setting: no ``settings.xml`` and mode 15 (justification is
known to differ in 15: ROADMAP.md 4.6).
"""

from __future__ import annotations

import probe_docx
import wml

SETTINGS = {"none": None, "15": 15}
FACES = ("Calibri", "Cambria", "Times New Roman", "Arial", "Aptos", "Courier New", "Georgia", "Verdana")
SIZES = (16, 22, 29, 40, 72)

DECOR = {
    "u-single": {"u": "single"},
    "u-double": {"u": "double"},
    "u-thick": {"u": "thick"},
    "u-words": {"u": "words"},
    "u-dotted": {"u": "dotted"},
    "u-dash": {"u": "dash"},
    "strike": {"strike": True},
    "dstrike": {"dstrike": True},
    "highlight": {"highlight": "yellow"},
    "shading": {"shd": {"val": "clear", "color": "auto", "fill": "BDD7EE"}},
}

PROSE = (
    "Line breaking is the first thing a word processor does that a slide renderer never has to "
    "do, and it is the reason this project exists as a sibling rather than a feature. A shape "
    "positions its text box once; a document flows its text through every column and page, so "
    "a single width that is wrong by a hair moves every break after it."
)


def fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def _spacing(before: int = 0, after: int = 0, line: int = 240, rule: str = "auto") -> dict:
    return {"before": before, "after": after, "line": line, "lineRule": rule}


def _side(name: str, sz: int, space: int, color: str) -> str:
    return f'<w:{name} w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


def _p(runs: str, **props) -> str:
    props.setdefault("spacing", _spacing())
    return wml.paragraph(runs, mark={}, **props)


def _anchor(name: str) -> tuple[str, str]:
    return ("anchor", _p(wml.run(f"Family {name}"), pageBreakBefore=True))


def blocks() -> list[tuple[str, str]]:
    """``(case name, w:p)`` in document order."""
    out: list[tuple[str, str]] = []
    # -- decorations: every face and size --------------------------------------------------
    for kind, props in DECOR.items():
        if kind in ("u-dotted", "u-dash"):
            faces, sizes = ("Calibri", "Times New Roman"), (22, 40)
        else:
            faces, sizes = FACES, SIZES
        out.append(_anchor(kind))
        for face in faces:
            for size in sizes:
                rpr = {"rFonts": fonts(face), "sz": size, "szCs": size}
                runs = (wml.run("Hx ", **rpr) + wml.run("Hxample decorated", **rpr, **props)
                        + wml.run(" tail", **rpr))
                out.append((f"decor/{kind}/{face}/{size}", _p(runs)))
    # -- highlight and shading against the line ---------------------------------------------
    out.append(_anchor("fill"))
    for size in (22, 40):
        for rule, line, before, after in (("auto", 240, 0, 0), ("auto", 360, 0, 0), ("exact", 400, 0, 0),
                                          ("atLeast", 480, 0, 0), ("auto", 240, 120, 120)):
            for kind in ("highlight", "shading"):
                rpr = {"rFonts": fonts("Calibri"), "sz": size, "szCs": size}
                runs = (wml.run("Hx ", **rpr) + wml.run("Hxample filled", **rpr, **DECOR[kind])
                        + wml.run(" tail", **rpr))
                out.append((f"fill/{kind}/{size}/{rule}{line}/{before}", _p(
                    runs, spacing=_spacing(before, after, line, rule))))
    # -- paragraph borders and shading --------------------------------------------------------
    out.append(_anchor("box"))
    colors = ("C00000", "0070C0", "00B050", "7030A0")
    for sz in (4, 12, 24):
        for space in (0, 4, 10):
            for indent in (0, 720):
                sides = "".join(_side(name, sz, space, color) for name, color in
                                zip(("top", "left", "bottom", "right"), colors))
                props = {"pBdr": sides}
                if indent:
                    props["ind"] = {"left": indent, "right": indent}
                out.append((f"box/all/{sz}/{space}/{indent}", _p(wml.run(f"Boxed {sz} {space} {indent}"),
                                                                  **props)))
                out.append((f"box/gap/{sz}/{space}/{indent}", _p(wml.run("Between boxes"))))
    for name, color in zip(("top", "left", "bottom", "right"), colors):
        for sz, space in ((8, 4), (24, 1)):
            out.append((f"box/{name}/{sz}/{space}/0", _p(wml.run(f"Only {name}"),
                                                         pBdr=_side(name, sz, space, color))))
            out.append((f"box/gap/{name}", _p(wml.run("Between boxes"))))
    shd = {"val": "clear", "color": "auto", "fill": "FFF2CC"}
    out.append(("box/shd/0/0/0", _p(wml.run("Shaded paragraph"), shd=shd)))
    out.append(("box/gap/shd", _p(wml.run("Between boxes"))))
    out.append(("box/shd-indent/0/0/720", _p(wml.run("Shaded and indented"), shd=shd,
                                              ind={"left": 720, "right": 1440})))
    out.append(("box/gap/shd2", _p(wml.run("Between boxes"))))
    sides = "".join(_side(name, 8, 4, "000000") for name in ("top", "left", "bottom", "right"))
    out.append(("box/shd-box/8/4/0", _p(wml.run("Shaded in a box"), shd=shd, pBdr=sides)))
    out.append(("box/gap/shd3", _p(wml.run("Between boxes"))))
    group = sides + _side("between", 8, 4, "000000")
    out.append(("box/group-1/8/4/0", _p(wml.run("First of a group"), pBdr=group)))
    out.append(("box/group-2/8/4/0", _p(wml.run("Second of a group"), pBdr=group)))
    out.append(("box/gap/group", _p(wml.run("After the group"))))
    # -- alignment ---------------------------------------------------------------------------
    for jc in ("center", "right", "both", "distribute"):
        out.append(_anchor(f"align-{jc}"))
        for face, size in (("Calibri", 22), ("Times New Roman", 22), ("Arial", 29)):
            for right in (0, 2000, 4000):
                rpr = {"rFonts": fonts(face), "sz": size, "szCs": size}
                out.append((f"align/{jc}/{face}/{size}/{right}", _p(
                    wml.run(PROSE, **rpr), jc=jc, ind={"left": 0, "right": right})))
    out.append(_anchor("centre-lines"))
    for k in range(24):
        text = "H" + "i" * k + "x"
        out.append((f"centre/{k}", _p(wml.run(text, rFonts=fonts("Calibri"), sz=23, szCs=23), jc="center")))
    # -- tabs -------------------------------------------------------------------------------
    out.append(_anchor("tabs"))
    stops = {
        "right": '<w:tab w:val="right" w:pos="5000"/>',
        "center": '<w:tab w:val="center" w:pos="4000"/>',
        "decimal": '<w:tab w:val="decimal" w:pos="4500"/>',
        "leader": '<w:tab w:val="left" w:leader="dot" w:pos="6000"/>',
    }
    for kind, stop in stops.items():
        for text in ("Hx", "Hxample text", "12.345", "1234.5", "7"):
            runs = wml.run("Lead") + wml.run("\t") + wml.run(text) + wml.run("\t") + wml.run("end")
            out.append((f"tabs/{kind}/{text}", _p(runs, tabs=f"{stop}<w:tab w:val=\"left\" w:pos=\"8000\"/>")))
    return out


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": _spacing()},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=extra)


def names() -> list[str]:
    return [name for name, _ in blocks()]


if __name__ == "__main__":
    import sys
    from pathlib import Path

    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
    for setting in SETTINGS:
        path = out / f"render-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
