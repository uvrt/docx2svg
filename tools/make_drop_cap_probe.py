#!/usr/bin/env python3
"""Drop caps: where Word puts a dropped letter, how large, and how the text goes beside it.

A drop cap is a paragraph in a frame (``w:framePr`` with ``w:dropCap`` ``drop`` or
``margin``, ``w:lines``, ``w:hSpace``) holding the letter, before the paragraph it drops
into; Word writes the letter's size, a line height and a lowering (``w:sz``, ``w:spacing``
exact, ``w:position``) on it when it makes one.  The layout stopped at a frame.

Every case is a page: ``Case N``, the drop cap's paragraph, the paragraph it drops into
(its text a label ``N.a`` and words, six lines or more), then ``After N``.  A4, the left
margin off the pixel grid (1,442 twips); Calibri 11 single unless the case says
otherwise.  Families (``CASES``):

* ``lines`` -- ``drop`` over 2, 3, 4 and 5 lines, with a size, an exact line and a
  lowering of the kind Word writes, the letters ``D`` and ``W``;
* ``bare`` -- ``w:lines`` 3 with nothing else stated (the letter at the text's size), and
  with only the size (does Word size the letter from ``w:lines``, or take what is
  written?);
* ``height`` -- ``w:lines`` 3 over a frame line far shorter (a small letter) and far
  taller than three lines;
* ``hspace`` -- ``w:hSpace`` 144 and 432;
* ``margin`` -- ``w:dropCap`` ``margin``, and with ``w:hSpace``;
* ``text`` -- text at 12 pt under ``auto`` 276 (a filesamples document's), a first-line
  indent of 432 on the text (and 0 on the paragraph, as that document has it), a right-
  aligned and a justified paragraph, two letters in the frame;
* ``short`` -- a paragraph of one line and of two, shorter than the drop, then another;
* round two: ``margin`` anchored to the page and to the margin, as well as to the text;
  ``before`` -- space before the text and before the drop cap; ``width`` -- the letters
  ``I M A O j T`` at 30 and 70 pt, in the first section and in a second whose left margin
  is on the pixel grid (1,440 twips): where the text beside a letter starts;
* round three: ``indent`` -- the drop cap's own paragraph indented (a first line of 432,
  as a filesamples document's Normal style indents it; a left indent of 360).

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_drop_cap_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
SETTINGS = {"none": None, "14": 14, "15": 15}
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}
WORDS = ("amber", "birch", "cedar", "delta", "ember", "fjord", "grove", "heron", "inlet", "juniper", "kestrel",
         "larch", "maple", "nectar", "osprey", "plover", "quartz", "raven", "sable", "thistle", "umber",
         "vireo", "willow", "yarrow", "zephyr")


def words(start: int, count: int) -> str:
    return " ".join(WORDS[(start + k) % len(WORDS)] for k in range(count))


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    letter: str = "D"
    #: ``w:framePr``'s attributes beyond ``w:wrap``/``w:vAnchor``/``w:hAnchor``.
    frame: dict = field(default_factory=lambda: {"dropCap": "drop", "lines": 3})
    #: The drop cap's ``w:sz``, exact line (twips) and ``w:position``; ``None``: unstated.
    size: int | None = 100
    line: int | None = 806
    position: int | None = -8
    #: The text's size and line rule, and its paragraph's ``w:ind`` and ``w:jc``.
    text_size: int = 22
    text_spacing: dict = field(default_factory=lambda: dict(SPACING))
    ind: dict | None = None
    jc: str | None = None
    #: Words in the paragraph it drops into, and a second paragraph's.
    count: int = 70
    second: int = 0
    #: ``w:hAnchor``; the space before the drop cap's paragraph and the text's.
    anchor: str = "text"
    cap_before: int = 0
    text_before: int = 0
    #: In the second section, whose left margin is on the pixel grid (1,440 twips).
    grid: bool = False
    #: The drop cap's paragraph's own ``w:ind``.
    cap_ind: dict | None = None


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []
    add = out.append
    for lines, size, line, position in ((2, 64, 537, -6), (3, 100, 806, -8), (4, 136, 1074, -10),
                                        (5, 172, 1343, -12)):
        add(Case("lines", f"drop over {lines} lines", frame={"dropCap": "drop", "lines": lines}, size=size,
                 line=line, position=position))
    add(Case("lines", "W over 3 lines", letter="W"))
    add(Case("bare", "lines 3, nothing else stated", size=None, line=None, position=None))
    add(Case("bare", "lines 3, only the size", line=None, position=None))
    add(Case("height", "lines 3 over a 20 pt letter on its own line", size=40, line=None, position=None))
    add(Case("height", "lines 3 over an exact line of 400", size=40, line=400, position=None))
    add(Case("height", "lines 3 over an exact line of 1600", size=100, line=1600, position=-8))
    add(Case("hspace", "hSpace 144", frame={"dropCap": "drop", "lines": 3, "hSpace": 144}))
    add(Case("hspace", "hSpace 432", frame={"dropCap": "drop", "lines": 3, "hSpace": 432}))
    add(Case("margin", "margin over 3 lines", frame={"dropCap": "margin", "lines": 3}))
    add(Case("margin", "margin, hSpace 144", frame={"dropCap": "margin", "lines": 3, "hSpace": 144}))
    add(Case("text", "12 pt auto 276", size=117, line=951, position=-10, text_size=24,
             text_spacing={"before": 0, "after": 0, "line": 276, "lineRule": "auto"}))
    add(Case("text", "first line 432", ind={"firstLine": 432}))
    add(Case("text", "right-aligned", jc="right"))
    add(Case("text", "justified", jc="both"))
    add(Case("text", "two letters", letter="DR"))
    add(Case("short", "one line, then another paragraph", count=6, second=60))
    add(Case("short", "two lines, then another paragraph", count=18, second=60))
    # Round two: the text's start in mode 15 is not on a whole pixel; the margin drop
    # cap as Word anchors it; space before.
    add(Case("margin", "margin, hAnchor page", frame={"dropCap": "margin", "lines": 3}, anchor="page"))
    add(Case("margin", "margin, hAnchor margin", frame={"dropCap": "margin", "lines": 3}, anchor="margin"))
    add(Case("before", "240 before the text", text_before=240))
    add(Case("before", "240 before the drop cap", cap_before=240))
    # Round three: a filesamples document's drop cap is indented by its Normal style.
    add(Case("indent", "the drop cap's first line indented 432", cap_ind={"firstLine": 432}))
    add(Case("indent", "the drop cap indented 360 left", cap_ind={"left": 360}))
    add(Case("indent", "the drop cap's first line indented 432, 12 pt auto 276", cap_ind={"firstLine": 432},
             size=117, line=951, position=-10, text_size=24,
             text_spacing={"before": 0, "after": 0, "line": 276, "lineRule": "auto"}, ind={"firstLine": 0}))
    for grid in (False, True):
        for letter in ("I", "M", "A", "O", "j", "T"):
            for size in (60, 140):
                add(Case("width", f"{letter} at {size // 2} pt" + (", margin on the grid" if grid else ""),
                         letter=letter, size=size, line=806, position=-8, grid=grid))
    return tuple(out)


CASES = _cases()


def _run(text: str, props: dict) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{probe_docx.escape(text)}</w:t></w:r>'


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        case_start = wml.paragraph(_run(f"Case {number}", {}), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += case_start
        cap_rpr = {}
        if case.position is not None:
            cap_rpr["position"] = case.position
        if case.size is not None:
            cap_rpr.update(sz=case.size, szCs=case.size)
        attrs = "".join(f' w:{k}="{v}"' for k, v in case.frame.items())
        frame = f'<w:framePr{attrs} w:wrap="around" w:vAnchor="text" w:hAnchor="{case.anchor}"/>'
        if case.grid and not any(c.grid for c in CASES[:number]):
            # The first case in the second section: the first section ends before it.
            out = out[:-len(case_start)] + wml.paragraph("", mark={}, spacing=SPACING, sect=section()) + case_start
        spacing = {"before": case.cap_before, "after": 0, "line": case.line, "lineRule": "exact"} \
            if case.line is not None else dict(SPACING, before=case.cap_before)
        cap_props = {"spacing": spacing}
        if case.cap_ind:
            cap_props["ind"] = case.cap_ind
        ppr = (f"<w:pPr><w:keepNext/>{frame}{wml._ordered(wml.PPR_ORDER, cap_props)}"
               f'<w:textAlignment w:val="baseline"/>{wml.rpr(**cap_rpr)}</w:pPr>')
        out += f"<w:p>{ppr}{_run(case.letter, cap_rpr)}</w:p>"
        text_rpr = {"sz": case.text_size, "szCs": case.text_size}
        props = {"spacing": dict(case.text_spacing, before=case.text_before)}
        if case.ind:
            props["ind"] = case.ind
        if case.jc:
            props["jc"] = case.jc
        out += wml.paragraph(_run(f"{number}.a " + words(number, case.count), text_rpr), mark=text_rpr, **props)
        if case.second:
            out += wml.paragraph(_run(f"{number}.b " + words(number + 3, case.second), text_rpr), mark=text_rpr,
                                 **props)
        out += wml.paragraph(_run(f"After {number}", {}), mark={}, spacing=SPACING)
    return out


def section(left: int | None = None) -> str:
    """The section: A4, the margins off the pixel grid, or the left one ``left``."""
    m = dict(MARGINS, left=left if left is not None else MARGINS["left"])
    return (f'<w:sectPr><w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    grid = any(case.grid for case in CASES)
    return probe_docx.package(body(), styles=styles, extra_parts=extra,
                              final_section=section(1440 if grid else None))


DOCUMENTS = tuple(f"drop-cap-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"drop-cap-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
