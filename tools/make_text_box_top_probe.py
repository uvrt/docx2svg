#!/usr/bin/env python3
"""Where a text box's first baseline and first glyph land, swept by fractions of a pixel.

``make_text_box_probe.py`` left a pixel of baseline in some boxes and a left edge on
exactly half a pixel that Word rounded either way (ROADMAP.md, F.3 and F.13).  A baseline
rounds to a device pixel (1/300 in) from where its line's box starts, so one box says only
on which side of a rounding edge its text starts; a sweep that moves one quantity by a
fraction of a pixel at a time says where that edge is, and so where Word starts the text.
Every family is a sweep of boxes, each a text box whose first line is labelled ``T<n>``
(a line after a break ``V<n>``, a second paragraph ``U<n>``), moving one quantity by
``step`` EMU per box on both axes at once -- the first glyph's pen x reads the left edge,
its baseline the top:

* ``size`` -- the box's offset against the page by whole twips, six to a pixel, at 8 to
  24 pt; ``anchor`` ``t``, ``ctr`` and ``b``; one line, two lines (a break), and one line
  over an 11 pt paragraph mark;
* ``emu`` -- the offset by single EMU about a twip's edge; ``sub`` -- by 127 EMU (a fifth
  of a twip);
* ``ins`` -- ``lIns`` and ``tIns`` (``bIns`` at ``anchor="b"``) by 127 EMU;
* ``line`` -- the outline's width by 254 EMU (its half by 127), ``a:noFill``;
* ``height`` -- the box's height, and its bottom inset, by 127 EMU below one line;
* ``para`` / ``margin`` -- the offset against the anchor's paragraph, which starts below
  a heading line, and against margins off the pixel grid (a section of its own, last);
* ``two`` -- two paragraphs, the first with space before;
* ``edge`` -- the inset swept under an outline half a fraction of a twip wide;
* ``spacing`` / ``face`` -- ``auto`` 276 and 200, ``exact`` and ``atLeast`` 400; Times New
  Roman and Arial at 11, 14 and 16 pt; anchors ``t`` and ``b``;
* ``turn`` -- a box flipped either way and turned 180 and 10 degrees (recorded only).

Boxes sit on a grid of whole pixels plus their sweep's offset.  Three documents: no
``settings.xml``, mode 14 and mode 15.  Reader: ``read_text_box_top_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_drawing_probe as drawing_probe
import probe_docx
import wml
from make_anchor_probe import FACE, Anchor
from make_drawing_probe import NO_LINE, graphic, preset, xfrm

SETTINGS = {"none": None, "14": 14, "15": 15}
EMU_PX = 3048
TWIP = 635
COLUMNS, ROWS = 5, 9
PITCH_X, PITCH_Y = 400, 320
#: Box size: one line of the label and a word at 24 pt, two lines, and the insets.
CX, CY = 380 * EMU_PX, 300 * EMU_PX
#: The section: A4, margins on the pixel grid but for the ``margin`` family's pages.
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440, "header": 708, "footer": 708, "gutter": 0}
OFF_GRID = {**MARGINS, "top": 1501, "left": 1443}


@dataclass(frozen=True)
class Sweep:
    family: str
    size: int = 22
    bold: bool = False
    anchor: str = "t"
    #: What moves: ``pos`` (the offset), ``ins`` (lIns and tIns, or bIns at ``b``), ``line``
    #: (the outline's width), ``height`` (the box's), ``bins`` (bIns alone).
    moves: str = "pos"
    step: int = TWIP
    steps: int = 6
    #: Where the sweep starts, EMU added to the base offset, inset or outline width.
    start: int = 0
    relative: str = "page"
    #: ``one`` line, ``break`` (two lines, a ``w:br`` between), ``two`` paragraphs.
    lines: str = "one"
    #: ``run``: the paragraph mark as the text; ``default``: the document's (11 pt).
    mark: str = "run"
    before: int = 0
    #: The paragraph's line rule: ``(line, lineRule)``.
    spacing: tuple = (240, "auto")
    face: str | None = None
    #: The outline's width (EMU, ``a:noFill``) under an ``ins`` sweep; ``a:xfrm`` flips
    #: and rotation (recorded only: the model does not turn a text box's text).
    outline: int = 0
    xfrm: str = ""

    @property
    def recorded_only(self) -> bool:
        return self.family == "turn"

    @property
    def note(self) -> str:
        parts = [self.family, f"{self.size / 2:g}pt" + (" bold" if self.bold else ""), self.anchor, self.lines]
        if self.moves != "pos":
            parts.append(self.moves)
        if self.relative != "page":
            parts.append(self.relative)
        if self.mark != "run":
            parts.append("mark " + self.mark)
        if self.spacing != (240, "auto"):
            parts.append(f"{self.spacing[1]} {self.spacing[0]}")
        if self.face:
            parts.append(self.face)
        if self.outline:
            parts.append(f"outline {self.outline}")
        if self.xfrm:
            parts.append(self.xfrm)
        return " ".join(parts)


def _sweeps() -> tuple[Sweep, ...]:
    out = []
    for size in (16, 18, 20, 21, 22, 24, 26, 28, 30, 32, 36, 40, 48):
        for anchor in ("t", "ctr", "b"):
            out.append(Sweep("size", size=size, anchor=anchor))
            out.append(Sweep("size", size=size, anchor=anchor, lines="break"))
            if size != 22:
                out.append(Sweep("size", size=size, anchor=anchor, mark="default"))
    for anchor in ("t", "b"):
        out.append(Sweep("size", size=32, bold=True, anchor=anchor))
        out.append(Sweep("size", size=32, bold=True, anchor=anchor, mark="default"))
    for size in (22, 28):
        out.append(Sweep("emu", size=size, step=1, steps=24, start=TWIP * 7 - 12))
        out.append(Sweep("sub", size=size, step=127, steps=24))
        out.append(Sweep("ins", size=size, moves="ins", step=127, steps=24))
        out.append(Sweep("ins", size=size, anchor="b", moves="ins", step=127, steps=24))
        out.append(Sweep("line", size=size, moves="line", step=254, steps=24))
        out.append(Sweep("two", size=size, lines="two", before=137, steps=12))
    # The box's height and its bottom inset below a line at the top: what the text
    # box's last line keeps below it.
    for size in (16, 32, 40, 48):
        out.append(Sweep("height", size=size, moves="height", step=127, steps=24))
    out.append(Sweep("height", size=32, moves="bins", step=127, steps=24))
    for anchor in ("ctr", "b"):
        out.append(Sweep("height", size=32, anchor=anchor, moves="height", step=127, steps=24))
    # An inset and half an outline of a fraction of a twip each: rounded apart or summed.
    for width in (635, 889):
        out.append(Sweep("edge", size=22, moves="ins", step=127, steps=24, outline=width))
    # Other line rules, and other faces, on the box's last line.
    for line, rule in ((276, "auto"), (400, "exact"), (400, "atLeast"), (200, "auto")):
        for anchor in ("t", "b"):
            for lines in ("one", "break"):
                out.append(Sweep("spacing", size=22, anchor=anchor, lines=lines, spacing=(line, rule)))
    for face in ("Times New Roman", "Arial"):
        for size in (22, 28, 32):
            for anchor in ("t", "b"):
                out.append(Sweep("face", size=size, anchor=anchor, face=face))
    # Turned boxes (recorded).
    for turn in ("flipH", "flipV", "rot 10800000", "rot 600000"):
        out.append(Sweep("turn", size=22, steps=2, xfrm=turn))
    out += [Sweep("para", size=size, relative="paragraph", steps=12) for size in (22, 28)]
    out += [Sweep("margin", size=size, relative="margin", steps=12) for size in (22, 28)]
    return tuple(out)


SWEEPS = _sweeps()
PER_PAGE = COLUMNS * ROWS


@dataclass(frozen=True)
class Box:
    number: int
    sweep: Sweep
    k: int

    @property
    def label(self) -> str:
        return f"T{self.number}"

    @property
    def keys(self) -> tuple[str, ...]:
        """The labels of its lines: ``T<n>``, and ``V<n>`` or ``U<n>``."""
        more = {"one": (), "break": (f"V{self.number}",), "two": (f"U{self.number}",)}[self.sweep.lines]
        return (self.label, *more)


def boxes() -> list[Box]:
    out = []
    for sweep in SWEEPS:
        for k in range(sweep.steps):
            out.append(Box(len(out), sweep, k))
    return out


BOXES = boxes()


def _paragraph(sweep: Sweep, runs: str, before: int = 0) -> str:
    mark = _run_props(sweep) if sweep.mark == "run" else {}
    line, rule = sweep.spacing
    return wml.paragraph(runs, mark=mark, spacing={"before": before, "after": 0, "line": line, "lineRule": rule})


def _run_props(sweep: Sweep) -> dict:
    run = {"sz": sweep.size, "szCs": sweep.size}
    if sweep.face:
        run = {"rFonts": {k: sweep.face for k in ("ascii", "hAnsi", "eastAsia", "cs")}, **run}
    if sweep.bold:
        run["b"] = True
    return run


def shape(box: Box) -> str:
    sweep, k = box.sweep, box.k
    props = _run_props(sweep)
    # A word after the label where it fits on the line at every size.
    word = " Hx" if sweep.size < 36 else ""
    runs = wml.run(f"{box.label}{word}", **props)
    if sweep.lines == "break":
        runs += f"<w:r>{wml.rpr(**props)}<w:br/></w:r>" + wml.run(f"V{box.number}{word}", **props)
    text = _paragraph(sweep, runs, sweep.before)
    if sweep.lines == "two":
        text += _paragraph(sweep, wml.run(f"U{box.number}{word}", **props))
    l, t, r, b = 91440, 45720, 91440, 45720
    line = f'<a:ln w="{sweep.outline}"><a:noFill/></a:ln>' if sweep.outline else NO_LINE
    moved = sweep.start + k * sweep.step
    if sweep.moves == "ins":
        l += moved
        if sweep.anchor == "b":
            b += moved
        else:
            t += moved
    elif sweep.moves == "line":
        line = f'<a:ln w="{moved}"><a:noFill/></a:ln>'
    elif sweep.moves == "bins":
        b += moved
    body = (f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}"'
            f' anchor="{sweep.anchor}" anchorCtr="0"><a:noAutofit/></wps:bodyPr>')
    turn = {"flipH": {"flip_h": True}, "flipV": {"flip_v": True}}.get(sweep.xfrm) or (
        {"rot": int(sweep.xfrm.split()[1])} if sweep.xfrm.startswith("rot") else {})
    return (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>{xfrm(0, 0, *extent(box), **turn)}{preset("rect")}<a:noFill/>'
            f"{line}</wps:spPr><wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx>{body}</wps:wsp>")


def extent(box: Box) -> tuple[int, int]:
    """The box's extent (EMU)."""
    if box.sweep.moves == "height":
        return CX, CY + box.sweep.start + box.k * box.sweep.step
    return CX, CY


def offset(box: Box, slot: int) -> tuple[int, int]:
    """The box's offset (EMU) on each axis from its frames, in slot ``slot`` of its page."""
    column, row = slot % COLUMNS, slot // COLUMNS
    x, y = column * PITCH_X * EMU_PX, row * PITCH_Y * EMU_PX
    if box.sweep.relative == "page":
        x, y = x + 150 * EMU_PX, y + 200 * EMU_PX
    elif box.sweep.relative == "paragraph":
        y -= 60 * EMU_PX  # the anchor's paragraph starts below a heading line
    if box.sweep.moves == "pos":
        moved = box.sweep.start + box.k * box.sweep.step
        x, y = x + moved, y + moved
    return x, y


def pages() -> list[list[Box]]:
    """Boxes to a page; a sweep against the margin or the paragraph starts a page of its
    own kind (``margin`` in the off-grid section, ``para`` below a heading line)."""
    out: list[list[Box]] = []
    kind = None
    for box in BOXES:
        here = box.sweep.relative
        if not out or len(out[-1]) == PER_PAGE or here != kind:
            out.append([])
            kind = here
        out[-1].append(box)
    return out


def body() -> str:
    """Each page: a heading paragraph holding the page's anchors (so their paragraph starts
    at the margin), or for the ``para`` family a heading and then the anchors' paragraph."""
    out = ""
    serial = 1
    spacing = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
    every = pages()
    for number, page in enumerate(every):
        relative = page[0].sweep.relative
        drawings = ""
        for slot, box in enumerate(page):
            x, y = offset(box, slot)
            h_frame = "page" if relative == "page" else "margin"
            drawings += Anchor((h_frame, "offset", x), (relative, "offset", y), *extent(box),
                               graphic=graphic(shape(box), "wps")).xml(serial)
            serial += 1
        heading = wml.run(f"Page {number} of text-box sweeps")
        if relative == "margin" and every[number - 1][0].sweep.relative != "margin":
            # The pages before end the first section; the margin family's section follows.
            out += wml.paragraph(wml.run(f"Page {number} ends a section."), mark={}, spacing=spacing,
                                 sect=section(MARGINS))
        if relative == "paragraph":
            out += wml.paragraph(heading, mark={}, spacing=spacing, pageBreakBefore=True)
            out += wml.paragraph(wml.run("Anchors.") + drawings, mark={}, spacing=spacing)
        else:
            out += wml.paragraph(heading + drawings, mark={}, spacing=spacing, pageBreakBefore=True)
    return out


def section(margins: dict) -> str:
    m = margins
    return (f'<w:sectPr><w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    parts = [drawing_probe.THEME]
    mode = SETTINGS[setting]
    if mode is not None:
        parts.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body(), styles=styles, extra_parts=tuple(parts), final_section=section(OFF_GRID))


if __name__ == "__main__":
    print(len(BOXES), "boxes on", len(pages()), "pages")
