#!/usr/bin/env python3
"""An empty table row: how tall Word makes it, and what ``w:hideMark`` changes.

Found on local documents: their empty spacer rows (a small stated height ``atLeast``, one
empty paragraph in each cell) stood shorter in Word than in the model -- the row's stated
height, where the model made it as tall as the empty paragraph's line.  Their cells carry ``w:hideMark``, which ECMA-376 (17.4.20)
says makes the end-of-cell mark "ignored when determining the height of the row" when
the cell is empty.  This probe measures what the element does and does not do.

Every case is a page: ``Case N``, a three-row, two-column table (``N.top a`` / ``N.top b``,
the **case row**, ``N.bot a`` / ``N.bot b``) with single borders everywhere, and ``After
N``.  The case row's height is the distance from the top row's baseline to the bottom
row's.  Calibri 11 pt, no paragraph spacing, unless the case says otherwise.  Families:

* ``control`` -- no ``w:hideMark``: an empty cell pair, with no stated height and with
  ``atLeast`` 170; ``w:hideMark`` on cells with text;
* ``height`` -- empty cells under ``w:hideMark``: no stated height, ``atLeast`` 170, 400
  and 1000, ``exact`` 170 and 400, and a ``w:trHeight`` with no ``w:hRule``;
* ``content`` -- what "empty" is: two empty paragraphs, a paragraph of one space, of a no
  -break space, of an empty run, of a tab, a 48 pt mark, a 6 pt mark, space before or after
  on the empty paragraph (120 twips), a paragraph with text beside an empty cell, a
  ``\\xa0`` paragraph and an empty one, the same with 240 twips after
  the first or before the second, and a paragraph of text before an empty one;
* ``cells`` -- ``w:hideMark`` on one cell only (the other empty, or with text), cell top
  and bottom margins (100 / 50) with and without ``atLeast`` 170, ``w:vAlign`` centre,
  thick inside borders, the case row first, last and alone in its table, and
  10 pt Verdana marks under ``vAlign`` centre and ``atLeast`` 170, with and without the element;
* ``body`` -- controls outside a cell: an empty 12 pt paragraph between two tables, after
  a table, and the same with ``w:hideMark`` on the tables' cells.

One document per compatibility setting (none, 12, 14, 15).  Reader:
``read_hide_mark_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import probe_docx
import wml

SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
VERDANA = {"ascii": "Verdana", "hAnsi": "Verdana", "eastAsia": "Verdana",
          "cs": "Verdana"}
WIDTH = 3000


@dataclass(frozen=True)
class Cell:
    """One cell of the case row: its content (a kind, see :func:`content`), and whether
    it carries ``w:hideMark``."""

    kind: str = "empty"
    hide: bool = True


@dataclass(frozen=True)
class Case:
    family: str
    name: str
    cells: tuple = (Cell(), Cell())
    #: ``(twips, hRule or None)``, or ``None`` for no ``w:trHeight``.
    height: tuple | None = None
    margins: tuple = (0, 0)
    align: str | None = None
    inside: int = 4
    #: Where the case row is: ``middle`` (between two text rows), ``first``, ``last``,
    #: ``alone``; or a ``body`` layout (``between``, ``after``).
    place: str = "middle"
    #: Verdana at this size (half-points) for the case row's marks, or ``None``: Calibri 11.
    verdana: int | None = None
    extra: dict = field(default_factory=dict)

    @property
    def key(self) -> str:
        return f"{self.family} {self.name}"


def _cases() -> tuple[Case, ...]:
    out = [
        Case("control", "empty", cells=(Cell(hide=False), Cell(hide=False))),
        Case("control", "empty atLeast 170", cells=(Cell(hide=False), Cell(hide=False)), height=(170, "atLeast")),
        Case("control", "text hideMark", cells=(Cell("text"), Cell("text"))),
        Case("control", "text hideMark atLeast 170", cells=(Cell("text"), Cell("text")), height=(170, "atLeast")),
    ]
    for height, label in ((None, "no height"), ((170, "atLeast"), "atLeast 170"), ((400, "atLeast"), "atLeast 400"),
                          ((1000, "atLeast"), "atLeast 1000"), ((170, "exact"), "exact 170"),
                          ((400, "exact"), "exact 400"), ((170, None), "170 no hRule")):
        out.append(Case("height", label, height=height))
    for kind in ("empty2", "space", "nbsp", "emptyrun", "tab", "mark48", "mark6", "before", "after", "nbsp+empty"):
        out.append(Case("content", kind, cells=(Cell(kind), Cell(kind)), height=(170, "atLeast")))
        if kind in ("empty2", "mark48", "before", "after"):
            out.append(Case("content", f"{kind} no height", cells=(Cell(kind), Cell(kind))))
    out.append(Case("content", "text beside empty", cells=(Cell("text"), Cell()), height=(170, "atLeast")))
    for kind in ("nbsp after+empty", "nbsp+empty before", "text+empty"):
        out.append(Case("content", f"{kind} no height", cells=(Cell(kind), Cell(kind))))
    out += [
        Case("cells", "one hidden, other empty", cells=(Cell(), Cell(hide=False)), height=(170, "atLeast")),
        Case("cells", "one hidden, other empty, no height", cells=(Cell(), Cell(hide=False))),
        Case("cells", "margins 100/50", margins=(100, 50)),
        Case("cells", "margins 100/50 atLeast 170", margins=(100, 50), height=(170, "atLeast")),
        Case("cells", "margins 100/50 atLeast 400", margins=(100, 50), height=(400, "atLeast")),
        Case("cells", "vAlign center atLeast 170", align="center", height=(170, "atLeast")),
        Case("cells", "vAlign bottom atLeast 400", align="bottom", height=(400, "atLeast")),
        Case("cells", "insideH 24 atLeast 170", inside=24, height=(170, "atLeast")),
        Case("cells", "insideH 24 no height", inside=24),
        Case("cells", "first row atLeast 170", place="first", height=(170, "atLeast")),
        Case("cells", "last row atLeast 170", place="last", height=(170, "atLeast")),
        Case("cells", "alone atLeast 170", place="alone", height=(170, "atLeast")),
        Case("cells", "alone no height", place="alone"),
        Case("cells", "verdana 10", align="center", height=(170, "atLeast"), verdana=20),
        Case("cells", "verdana 10 no hideMark", cells=(Cell(hide=False), Cell(hide=False)), align="center",
             height=(170, "atLeast"), verdana=20),
        Case("body", "between tables", place="between"),
        Case("body", "after a table", place="after"),
        Case("body", "between hidden tables", place="between", extra={"hide": True}),
        Case("body", "after a hidden table", place="after", extra={"hide": True}),
    ]
    return tuple(out)


CASES = _cases()

RUN = {"rFonts": FACE, "sz": 22, "szCs": 22}


def _side(name: str, sz: int) -> str:
    return f'<w:{name} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'


def _tbl_pr(case: Case) -> str:
    top, bottom = case.margins
    borders = (_side("top", 4) + _side("left", 4) + _side("bottom", 4) + _side("right", 4)
               + _side("insideH", case.inside) + _side("insideV", 4))
    return (f'<w:tblPr><w:tblW w:w="{2 * WIDTH}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
            '<w:tblLayout w:type="fixed"/>'
            f'<w:tblCellMar><w:top w:w="{top}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
            ' w:noHBand="1" w:noVBand="1"/></w:tblPr>')


def content(kind: str, label: str, mark: dict) -> str:
    """A cell's paragraphs for ``kind``; ``label`` is the text of a ``text`` cell."""
    empty = wml.paragraph("", mark=mark)
    if kind == "empty":
        return empty
    if kind == "text":
        return wml.paragraph(wml.run(label, **mark), mark=mark)
    if kind == "empty2":
        return empty + empty
    if kind == "space":
        return wml.paragraph(wml.run(" ", **mark), mark=mark)
    if kind == "nbsp":
        return wml.paragraph(wml.run(" ", **mark), mark=mark)
    if kind == "nbsp+empty":
        return wml.paragraph(wml.run(" ", **mark), mark=mark) + empty
    if kind == "nbsp after+empty":
        spaced = {"before": 0, "after": 240, "line": 240, "lineRule": "auto"}
        return wml.paragraph(wml.run("\u00a0", **mark), mark=mark, spacing=spaced) + empty
    if kind == "nbsp+empty before":
        spaced = {"before": 240, "after": 0, "line": 240, "lineRule": "auto"}
        return wml.paragraph(wml.run("\u00a0", **mark), mark=mark) + wml.paragraph("", mark=mark, spacing=spaced)
    if kind == "text+empty":
        return wml.paragraph(wml.run(label, **mark), mark=mark) + empty
    if kind == "emptyrun":
        return wml.paragraph(f"<w:r>{wml.rpr(**mark)}</w:r>", mark=mark)
    if kind == "tab":
        return wml.paragraph(f"<w:r>{wml.rpr(**mark)}<w:tab/></w:r>", mark=mark)
    if kind in ("mark48", "mark6"):
        size = {"mark48": 96, "mark6": 12}[kind]
        big = dict(mark, sz=size, szCs=size)
        return wml.paragraph("", mark=big)
    if kind == "before":
        return wml.paragraph("", mark=mark, spacing={"before": 120, "after": 0, "line": 240, "lineRule": "auto"})
    if kind == "after":
        return wml.paragraph("", mark=mark, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"})
    raise ValueError(kind)


def _tc(inner: str, *, hide: bool, align: str | None) -> str:
    pr = f'<w:tcW w:w="{WIDTH}" w:type="dxa"/>'
    if align:
        pr += f'<w:vAlign w:val="{align}"/>'
    if hide:
        pr += "<w:hideMark/>"
    return f"<w:tc><w:tcPr>{pr}</w:tcPr>{inner}</w:tc>"


def _text_row(label: str, hide: bool = False) -> str:
    return "<w:tr>" + "".join(_tc(wml.paragraph(wml.run(f"{label} {side}", **RUN)), hide=hide, align=None)
                              for side in "ab") + "</w:tr>"


def _case_row(number: int, case: Case) -> str:
    tr = ""
    if case.height is not None:
        value, rule = case.height
        tr = f'<w:trPr><w:trHeight w:val="{value}"' + (f' w:hRule="{rule}"' if rule else "") + "/></w:trPr>"
    mark = {"rFonts": VERDANA, "sz": case.verdana, "szCs": case.verdana} if case.verdana else RUN
    cells = "".join(_tc(content(cell.kind, f"{number}.mid {side}", mark), hide=cell.hide, align=case.align)
                    for cell, side in zip(case.cells, "ab"))
    return f"<w:tr>{tr}{cells}</w:tr>"


def _table(rows: str, case: Case) -> str:
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * 2
    return f"<w:tbl>{_tbl_pr(case)}<w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def case_body(number: int, case: Case) -> str:
    props = {"pageBreakBefore": True} if number else {}
    out = wml.paragraph(wml.run(f"Case {number}", **RUN), **props)
    if case.place in ("between", "after"):
        hide = bool(case.extra.get("hide"))
        mark = {"rFonts": VERDANA, "sz": 24, "szCs": 24, "b": True}
        out += _table(_text_row(f"{number}.top", hide), case)
        out += wml.paragraph("", mark=mark)
        if case.place == "between":
            out += _table(_text_row(f"{number}.bot", hide), case)
        else:
            out += wml.paragraph(wml.run(f"{number}.bot", **RUN))
    else:
        rows = {
            "middle": _text_row(f"{number}.top") + _case_row(number, case) + _text_row(f"{number}.bot"),
            "first": _case_row(number, case) + _text_row(f"{number}.bot"),
            "last": _text_row(f"{number}.top") + _case_row(number, case),
            "alone": _case_row(number, case),
        }[case.place]
        out += _table(rows, case)
    return out + wml.paragraph(wml.run(f"After {number}", **RUN))


def body() -> str:
    return "".join(case_body(number, case) for number, case in enumerate(CASES))


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


DOCUMENTS = tuple(f"hide-mark-{setting}" for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"hide-mark-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
    print(len(CASES), "cases")
