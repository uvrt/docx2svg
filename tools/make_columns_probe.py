#!/usr/bin/env python3
"""Text columns: where Word puts a multi-column section's columns and how it fills them.

A section's ``w:cols`` splits its text area into columns: ``w:num`` of them, ``w:space``
twips apart, of equal width -- or, under ``w:equalWidth="0"``, each ``w:col`` its own
``w:w`` and the ``w:space`` after it -- and ``w:sep`` draws a line between them.  Text fills
the first column, then the next, then the next page; ``w:br w:type="column"`` ends a
column.  A section that a ``continuous`` section break follows shares its page with the
next, and Word balances its columns.  Nothing of this was laid out: a section of several
columns stopped the layout.

Every document is Calibri 11 pt on A4, single, no spacing unless stated, the left margin
off the pixel grid (1442 twips), in three settings -- no ``settings.xml``, mode 14 and
mode 15.  Paragraphs of named lines (``Name l0``, ``Name l1``... each ended by ``w:br``)
are what a measured line count needs; paragraphs of running words are what line breaking
in a column needs.  A column of this page holds 51 single lines.  Documents
(``KINDS``):

* ``geometry`` -- nine sections (``nextPage``), each of columns that put their edges on
  every fraction of a pixel: two columns 720 and 355 apart, three 401 apart, four 113
  apart, three touching, five 250 apart, and unequal ones (``w:equalWidth="0"``) whose
  widths sum to the text width and whose widths do not, one in a section with a gutter;
  every column holds a left-aligned and a right-aligned line, then a column break;
* ``flow2`` -- two columns, 720 apart: running paragraphs (space after 120, headings with
  ``keepNext`` and space before 240) over two and a half pages;
* ``flow3`` -- three columns, 401 apart, the same over a page and a half;
* ``unequal`` -- three columns of 2200, 4000 and the rest, 400 and 300 apart: running
  paragraphs that go on from a narrow column into a wide one and back;
* ``keeps`` -- two columns; sections (``nextPage``) whose one-line fillers leave a
  paragraph at the first column's foot: an orphan, a widow, a ``keepNext`` heading, a
  ``keepLines`` paragraph, widow control off, and an orphan at the second column's foot;
* ``breaks`` -- column breaks: in a line, at a paragraph's start, alone in a paragraph,
  two in a row, in the last column, before a paragraph with space before; a page break
  and ``pageBreakBefore`` in a column; a column break in three columns;
* ``sep`` -- ``w:sep``: two columns over a page and a half, three columns on part of a
  page, unequal columns;
* ``balance`` -- sections of two and three columns between one-column sections, all
  ``continuous``: one paragraph of 7 lines, paragraphs of 3, 4 and 2 lines with space
  after, 10 one-line paragraphs in three columns, 5 lines, a ``keepLines`` paragraph, a
  16 pt line, one line, two lines, four lines in three columns, a column break, ``w:sep``,
  running words, a three-line paragraph;
* ``balancelong`` -- a balanced section of 140 lines, over a page and on to the next;
* ``table`` -- tables in a column: fixed, autofit in percent, and one whose rows go on
  into the next column;
* ``picture`` -- inline pictures in a column, one at the column's foot;
* ``header`` -- a header and a footer on a two-column section's pages;
* ``footnote`` -- footnotes referenced from both columns;
* ``notesright`` -- notes referenced only from the second column;
* ``notes3`` -- three columns, notes from each;
* ``notesunequal`` -- unequal columns of 2200, 4000 and the rest, a note from each;
* ``notesover`` -- a note of some 25 lines near the first column's foot, and the page
  after;
* ``notesregion`` -- a one-column region, a balanced two-column region and a one-column
  region on one page, a note from each;
* ``endnotes`` -- endnotes referenced from both columns, set after the text;
* ``dropcap`` -- drop caps in columns: near the first column's top, at the second's top
  after a column break, and one in the margin;
* ``floating`` -- floating pictures in columns: text above and below one at the first
  column's left, around one aligned right in the second, behind one 360 twips into it;
* ``regions`` -- the space around a continuous break between one column and two: space
  after and before on the paragraphs either side, an empty paragraph holding the break
  as Word writes one (spaced and not), a balanced region whose columns end in space
  after, a second column that starts a paragraph with space before, and a column's top
  on a new page.

Reader: ``read_columns_probe.py``.
"""

from __future__ import annotations

import io
import zipfile

import make_anchor_probe as anchor_probe
import make_endnote_probe as endnote_probe
import make_footnote_draw_probe as footnote_probe
import make_picture_probe as picture_probe
import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
SETTINGS = {"none": None, "14": 14, "15": 15}
KINDS = ("geometry", "flow2", "flow3", "unequal", "keeps", "breaks", "sep", "balance", "balancelong", "table",
         "picture", "header", "footnote", "regions", "notesright", "notes3", "notesunequal", "notesover",
         "notesregion", "endnotes", "dropcap", "floating")
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}
#: The text width: 11906 - 1442 - 1300.
TEXT = 9164
#: Single lines a column of this page holds.
COLUMN_LINES = 51
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HEADER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

WORDS = footnote_probe.WORDS
words = footnote_probe.words


def text(value: str, **props) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{probe_docx.escape(value)}</w:t></w:r>'


def column_break() -> str:
    return '<w:r><w:br w:type="column"/></w:r>'


def page_break() -> str:
    return '<w:r><w:br w:type="page"/></w:r>'


def p(runs: str, **props) -> str:
    props.setdefault("spacing", SPACING)
    return wml.paragraph(runs, mark={}, **props)


def named(name: str, count: int, **props) -> str:
    """One paragraph of ``count`` lines, ``name l0`` to ``name l<count-1>``, ended by ``w:br``."""
    rpr = wml.rpr(**{k: v for k, v in props.items() if k in ("sz", "szCs")})
    runs = f"<w:r>{rpr}" + "<w:br/>".join(
        f'<w:t xml:space="preserve">{name} l{k}</w:t>' for k in range(count)) + "</w:r>"
    return p(runs, **{k: v for k, v in props.items() if k not in ("sz", "szCs")})


def fillers(name: str, count: int) -> str:
    return "".join(p(text(f"{name} f{k}")) for k in range(count))


def running(name: str, k: int, size: int, **props) -> str:
    return p(text(f"{name} " + words(k, size) + "."), **props)


def cols(num: int = 1, space: int = 720, sep: bool = False, widths: tuple = ()) -> str:
    sep_xml = ' w:sep="1"' if sep else ""
    if widths:
        inner = "".join(f'<w:col w:w="{w}" w:space="{s}"/>' for w, s in widths)
        return f'<w:cols w:num="{len(widths)}" w:space="{space}"{sep_xml} w:equalWidth="0">{inner}</w:cols>'
    return f'<w:cols w:num="{num}" w:space="{space}"{sep_xml}/>'


def section(columns: str = "", kind: str | None = None, gutter: int = 0, refs: str = "") -> str:
    m = dict(MARGINS, gutter=gutter)
    kind_xml = f'<w:type w:val="{kind}"/>' if kind else ""
    return (f'<w:sectPr xmlns:r="{R_NS}">{refs}{kind_xml}<w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/>'
            f"{columns or cols()}</w:sectPr>")


#: ``geometry``'s sections: ``(num, space, widths, gutter)``.
GEOMETRY = (
    (2, 720, (), 0), (2, 355, (), 0), (3, 401, (), 0), (4, 113, (), 0), (3, 0, (), 0), (5, 250, (), 0),
    (3, 0, ((2000, 501), (3500, 301), (TEXT - 2000 - 501 - 3500 - 301, 0)), 0),
    (2, 0, ((1501, 1033), (2999, 0)), 0),
    (2, 0, ((3001, 777), (TEXT - 300 - 3001 - 777, 0)), 300),
)


def _geometry() -> tuple[str, str]:
    body = ""
    final = ""
    for k, (num, space, widths, gutter) in enumerate(GEOMETRY):
        count = len(widths) or num
        for c in range(count):
            body += p(text(f"S{k} c{c} left"))
            last = c == count - 1
            body += p(text(f"S{k} c{c} right") + ("" if last else column_break()), jc="right")
        columns = cols(num, space, widths=widths)
        if k < len(GEOMETRY) - 1:
            body += p("", sect=section(columns, None if k == 0 else "nextPage", gutter))
        else:
            final = section(columns, "nextPage", gutter)
    return body, final


def _flow(name: str, count: int, columns: str) -> tuple[str, str]:
    body = ""
    for k in range(count):
        if k % 7 == 3:
            body += running(f"{name} H{k}", k, 3, keepNext=True,
                            spacing={"before": 240, "after": 120, "line": 240, "lineRule": "auto"})
        size = (9, 37, 64, 18, 120, 5, 46, 83)[k % 8]
        body += running(f"{name} P{k}", k, size, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"})
    return body, section(columns)


def _keeps() -> tuple[str, str]:
    body = ""
    two = cols(2, 720)
    cases = (
        ("orphan", COLUMN_LINES - 1, lambda: named("Orphan", 6)),
        ("widow", COLUMN_LINES - 5, lambda: named("Widow", 6)),
        ("next", COLUMN_LINES - 1, lambda: named("Heading", 1, keepNext=True) + named("Next", 4)),
        ("lines", COLUMN_LINES - 3, lambda: named("Kept", 8, keepLines=True)),
        ("off", COLUMN_LINES - 1, lambda: named("Off", 6, widowControl=False)),
        ("second", 2 * COLUMN_LINES - 1, lambda: named("Second", 6)),
    )
    for k, (name, count, target) in enumerate(cases):
        body += fillers(name, count) + target() + fillers(f"{name} after", 3)
        if k < len(cases) - 1:
            body += p("", sect=section(two, None if k == 0 else "nextPage"))
    return body, section(two, "nextPage")


def _breaks() -> tuple[str, str]:
    body = ""
    body += named("Start", 3)
    body += p(text("Mid before") + column_break() + text("Mid after"))
    body += named("Second column", 2)
    body += p(column_break() + text("Opening break"))
    body += named("Page two", 2)
    body += p(column_break())
    body += named("After alone", 2)
    body += p(text("Twice") + column_break() + column_break() + text("After twice"))
    body += named("Spaced", 2)
    body += p(column_break())
    body += p(text("Space before 360"), spacing={"before": 360, "after": 0, "line": 240, "lineRule": "auto"})
    body += named("Page break", 2)
    body += p(text("Paged") + page_break() + text("After page break"))
    body += named("Then", 2)
    body += p(column_break())
    body += p(text("Before break before"))
    body += p(text("Page break before"), pageBreakBefore=True)
    body += named("Last", 2, sect=section(cols(2, 720)))
    body += named("Three", 2)
    body += p(text("Three a") + column_break() + text("Three b") + column_break() + text("Three c"))
    body += p(column_break() + text("Three next page"))
    return body, section(cols(3, 401), "nextPage")


def _sep() -> tuple[str, str]:
    body = "".join(running(f"Sep P{k}", k, (40, 70, 25, 90)[k % 4]) for k in range(26))
    body += p("", sect=section(cols(2, 720, sep=True)))
    body += "".join(running(f"Three P{k}", k, (30, 50)[k % 2]) for k in range(6))
    body += p("", sect=section(cols(3, 401, sep=True), "nextPage"))
    body += "".join(running(f"Unequal P{k}", k, (30, 50)[k % 2]) for k in range(6))
    widths = ((2000, 900), (TEXT - 2900, 0))
    return body, section(cols(sep=True, widths=widths), "nextPage")


def _balance() -> tuple[str, str]:
    one = cols()
    out = ""
    groups = [
        # (columns, content); each followed by a one-column paragraph, all continuous.
        (cols(2, 720), named("Seven", 7)),
        (cols(2, 720), named("Three", 3, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"})
         + named("Four", 4, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"})
         + named("Two", 2, spacing={"before": 0, "after": 120, "line": 240, "lineRule": "auto"})),
        (cols(3, 401), "".join(p(text(f"Ten {k}")) for k in range(10))),
        (cols(2, 720), named("Five", 5)),
        ("page", ""),
        (cols(2, 720), named("Before kept", 4) + named("Kept", 5, keepLines=True)),
        (cols(2, 720), named("Small", 3) + named("Large", 1, sz=32, szCs=32) + named("Small after", 3)),
        (cols(2, 720), named("One", 1)),
        (cols(2, 720), named("Pair", 2)),
        (cols(3, 401), "".join(p(text(f"Four {k}")) for k in range(4))),
        (cols(2, 720), named("Broken", 2) + p(column_break() + text("After break")) + named("Broken tail", 1)),
        ("page", ""),
        (cols(2, 720, sep=True), named("Ruled", 9)),
        (cols(2, 720), running("Words", 0, 60) + running("Words", 3, 45)),
        (cols(2, 720), named("Triple", 3)),
        (cols(3, 401), named("Eleven", 11)),
        (cols(2, 720), "".join(p(text(f"Many {k}")) for k in range(6)) + named("Widowed", 3)),
    ]
    out += p(text("Balance opening one column line"))
    kind = None
    for k, (columns, content) in enumerate(groups):
        if columns == "page":
            kind = "nextPage"
            continue
        # The one-column paragraph before ends the one-column section.
        out += p(text(f"One column before {k}"), sect=section(one, kind))
        kind = "continuous"
        out += content
        out += p("", sect=section(columns, "continuous"))
    out += p(text("One column at the end"))
    return out, section(one, "continuous")


def _balance_long() -> tuple[str, str]:
    body = p(text("Long opening"), sect=section(cols()))
    body += "".join(p(text(f"Long {k}")) for k in range(140))
    body += p("", sect=section(cols(2, 720), "continuous"))
    body += p(text("After the long section"))
    return body, section(cols(), "continuous")


def _table() -> tuple[str, str]:
    def cell(value: str, width: int, kind: str = "dxa") -> str:
        return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="{kind}"/></w:tcPr>'
                f"{p(text(value))}</w:tc>")

    def table(rows: int, widths: tuple, kind: str = "dxa", total: str = "") -> str:
        borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
        # Fixed, but autofit where the width is in percent (a fixed table in percent is
        # not laid out: Tables -- measured).
        layout = "" if total else '<w:tblLayout w:type="fixed"/>'
        total = total or '<w:tblW w:w="0" w:type="auto"/>'
        tbl_pr = (f"<w:tblPr>{total}<w:tblBorders>{borders}</w:tblBorders>"
                  f'{layout}<w:tblLook w:val="0000"/></w:tblPr>')
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        body = "".join("<w:tr>" + "".join(cell(f"r{r} c{c}", w if kind == "dxa" else w, kind)
                                          for c, w in enumerate(widths)) + "</w:tr>" for r in range(rows))
        return f"<w:tbl>{tbl_pr}<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"

    body = named("Table opening", 2)
    body += table(3, (1800, 1800)) + p(text("After fixed"))
    body += table(2, (2111, 2111), total='<w:tblW w:w="5000" w:type="pct"/>') + p(text("After percent"))
    body += fillers("Table filler", COLUMN_LINES - 12)
    body += table(12, (1500, 2000)) + p(text("After long"))
    return body, section(cols(2, 720))


def _picture() -> tuple[str, str]:
    body = named("Picture opening", 2)
    body += p(picture_probe.picture(1, 1500000, 600000))
    body += p(text("Beside ") + picture_probe.picture(2, 900000, 300000) + text(" text"))
    body += fillers("Picture filler", COLUMN_LINES - 12)
    body += p(picture_probe.picture(3, 2000000, 1000000))
    body += named("Picture after", 3)
    return body, section(cols(2, 720))


def _header() -> tuple[str, str, list]:
    def story(root: str, value: str) -> str:
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:{root} xmlns:w="{probe_docx.W_NS}">'
                + p(text(value + " left")) + p(text(value + " right"), jc="right") + f"</w:{root}>")

    extra = [("word/header1.xml", HEADER_CONTENT_TYPE, HEADER_REL, story("hdr", "Header")),
             ("word/footer1.xml", FOOTER_CONTENT_TYPE, FOOTER_REL, story("ftr", "Footer"))]
    refs = '<w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/>'
    body = "".join(running(f"Header P{k}", k, (40, 70, 25, 90)[k % 4]) for k in range(22))
    return body, section(cols(2, 720), refs=refs), extra


class _Notes:
    """The footnotes a document references, in order: ``ref`` writes a reference to a new
    note of ``value`` and returns its run."""

    def __init__(self):
        self.notes: list = []

    def ref(self, value: str) -> str:
        self.notes.append(footnote_probe.Note((value,)))
        return footnote_probe.reference(len(self.notes))

    def parts(self) -> list:
        footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     f'<w:footnotes xmlns:w="{probe_docx.W_NS}">{footnote_probe.separators()}'
                     + "".join(n.xml(k + 1) for k, n in enumerate(self.notes)) + "</w:footnotes>")
        return [("word/footnotes.xml", footnote_probe.FOOTNOTES_CONTENT_TYPE, footnote_probe.FOOTNOTES_REL,
                 footnotes)]


def _footnote() -> tuple[str, str, list]:
    notes = _Notes()
    note = notes.ref
    body = p(text("Notes in columns") + note("first " + words(0, 6)) + text(" " + words(1, 8) + "."))
    body += fillers("Note filler", 30)
    body += p(text("Second") + note("second " + words(2, 30)) + text(" " + words(3, 12) + "."))
    body += fillers("Note filler later", 30)
    body += p(text("Third") + note("third " + words(4, 4)) + text(" " + words(5, 6) + "."))
    body += fillers("Note filler last", 20)
    return body, section(cols(2, 720)), notes.parts()


def _notes_right() -> tuple[str, str, list]:
    """Notes referenced only from the second column: what the first keeps free of them."""
    notes = _Notes()
    body = fillers("Left filler", COLUMN_LINES + 2)
    body += p(text("Right") + notes.ref("right " + words(6, 44)) + text(" " + words(7, 5) + "."))
    body += fillers("Right filler", 6)
    body += p(text("Right two") + notes.ref("two " + words(8, 12)) + text(" " + words(9, 3) + "."))
    body += p(text("Right three") + notes.ref("three " + words(10, 2)))
    body += fillers("Right after", 8)
    return body, section(cols(2, 720)), notes.parts()


def _notes_three() -> tuple[str, str, list]:
    """Three columns, notes from each: one of two lines, one of six and one of one, one of
    three."""
    notes = _Notes()
    body = p(text("Col one") + notes.ref("one " + words(0, 8)) + text(" " + words(1, 3) + "."))
    body += fillers("Three filler", 44)
    body += p(text("Col two") + notes.ref("two " + words(2, 24)) + text(" and") + notes.ref("short " + words(3, 1)))
    body += fillers("Three filler mid", 44)
    body += p(text("Col three") + notes.ref("three " + words(4, 12)) + text(" " + words(5, 2) + "."))
    body += fillers("Three filler end", 20)
    return body, section(cols(3, 401)), notes.parts()


def _notes_unequal() -> tuple[str, str, list]:
    """Unequal columns of 2200, 4000 and the rest, a note from each: a narrow column's
    notes, under a separator longer than it is wide."""
    notes = _Notes()
    body = p(text("Narrow") + notes.ref("narrow " + words(0, 10)) + text(" " + words(1, 2) + "."))
    body += fillers("Unequal filler", 48)
    body += p(text("Wide") + notes.ref("wide " + words(2, 30)) + text(" " + words(3, 4) + "."))
    body += fillers("Unequal filler mid", 46)
    body += p(text("Last") + notes.ref("last " + words(4, 14)) + text(" " + words(5, 2) + "."))
    body += fillers("Unequal filler end", 10)
    return body, section(cols(widths=((2200, 400), (4000, 300), (TEXT - 6900, 0)))), notes.parts()


def _notes_over() -> tuple[str, str, list]:
    """A note of some 25 lines referenced near the first column's foot: where the rest of
    it goes, and the next page's separator."""
    notes = _Notes()
    body = fillers("Over filler", 36)
    body += p(text("Long") + notes.ref("long " + words(0, 150)) + text(" " + words(1, 3) + "."))
    body += fillers("Over filler mid", 40)
    body += p(text("Later") + notes.ref("later " + words(2, 6)) + text(" " + words(3, 3) + "."))
    body += fillers("Over filler end", 70)
    return body, section(cols(2, 720)), notes.parts()


def _notes_region() -> tuple[str, str, list]:
    """A page of a one-column region, a balanced two-column region and a one-column region
    after it, a note referenced from each."""
    notes = _Notes()
    one, two = cols(), cols(2, 720)
    body = p(text("Region one") + notes.ref("above " + words(0, 9)) + text(" " + words(1, 3) + "."))
    body += named("Above", 3, sect=section(one))
    body += named("Columns", 6) + p(text("Column note") + notes.ref("columns " + words(2, 20)))
    body += named("Columns after", 5) + p("", sect=section(two, "continuous"))
    body += p(text("Below") + notes.ref("below " + words(3, 5)) + text(" " + words(4, 3) + "."))
    body += named("Below", 3)
    return body, section(one, "continuous"), notes.parts()


def _endnotes() -> tuple[str, str, list]:
    """Endnotes referenced from both columns: where Word sets them after the text."""
    notes: list = []

    def note(value) -> str:
        notes.append(endnote_probe.Note(value))
        return endnote_probe.reference(len(notes))

    body = p(text("Endnotes in columns") + note(("first " + words(0, 6),)) + text(" " + words(1, 8) + "."))
    body += fillers("End filler", 50)
    body += p(text("Second") + note(("second " + words(2, 40),)) + text(" " + words(3, 12) + "."))
    body += p(text("Third") + note(("third " + words(4, 4), "third again " + words(5, 9))) + text(" " + words(5, 6)))
    body += fillers("End filler last", 12)
    endnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:endnotes xmlns:w="{probe_docx.W_NS}">{endnote_probe.separators()}'
                + "".join(n.xml(k + 1) for k, n in enumerate(notes)) + "</w:endnotes>")
    extra = [("word/endnotes.xml", endnote_probe.ENDNOTES_CONTENT_TYPE, endnote_probe.ENDNOTES_REL, endnotes)]
    return body, section(cols(2, 720)), extra


def drop_cap(letter: str, kind: str = "drop") -> str:
    """A drop cap of three lines, as ``make_drop_cap_probe.py``'s first cases: 50 pt, an
    exact line of 806, lowered 8 half points."""
    cap = {"sz": 100, "szCs": 100, "position": -8}
    frame = f'<w:framePr w:dropCap="{kind}" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'
    spacing = {"before": 0, "after": 0, "line": 806, "lineRule": "exact"}
    ppr = (f"<w:pPr><w:keepNext/>{frame}{wml._ordered(wml.PPR_ORDER, {'spacing': spacing})}"
           f'<w:textAlignment w:val="baseline"/>{wml.rpr(**cap)}</w:pPr>')
    return f"<w:p>{ppr}{text(letter, **cap)}</w:p>"


def _drop_cap() -> tuple[str, str]:
    """Drop caps in columns: one near the first column's top, one dropped into the second
    after a column break, one in the margin of the second."""
    body = named("Cap opening", 2)
    body += drop_cap("D") + running("Dropped one", 0, 40)
    body += fillers("Cap filler", 20)
    body += p(text("Before break") + column_break())
    body += drop_cap("W") + running("Dropped two", 3, 50)
    body += fillers("Cap filler two", 6)
    body += drop_cap("M", "margin") + running("In the margin", 5, 45)
    body += fillers("Cap filler end", 4)
    return body, section(cols(2, 720))


def _floating() -> tuple[str, str]:
    """Floating pictures in columns: in the first column one text wraps above and below,
    at the column's left; in the second one text wraps around, aligned right in the
    column, and one behind the text 360 twips into it."""
    emu = 635
    top_bottom = anchor_probe.Anchor(("column", "offset", 0), ("paragraph", "offset", 0), 1500000, 600000,
                                     wrap="<wp:wrapTopAndBottom/>", dist=(0, 0, 0, 0))
    square = anchor_probe.Anchor(("column", "align", "right"), ("paragraph", "offset", 0), 1200000, 900000,
                                 wrap='<wp:wrapSquare wrapText="bothSides"/>', dist=(0, 0, 114300, 114300))
    behind = anchor_probe.Anchor(("column", "offset", 360 * emu), ("paragraph", "offset", 0), 900000, 300000,
                                 behind=True, dist=(0, 0, 0, 0))
    body = named("Float opening", 3)
    body += p(top_bottom.xml(1) + text("Above and below " + words(0, 30) + "."))
    body += "".join(running(f"Float P{k}", k, (40, 70, 25, 90)[k % 4]) for k in range(9))
    body += p(text("Before break") + column_break())
    body += named("Second opening", 2)
    body += p(square.xml(2) + text("Around " + words(3, 80) + "."))
    body += p(behind.xml(3) + text("Behind " + words(5, 30) + "."))
    body += fillers("Float filler", 6)
    return body, section(cols(2, 720))


def spaced(before: int, after: int) -> dict:
    return {"before": before, "after": after, "line": 240, "lineRule": "auto"}


def _regions() -> tuple[str, str]:
    one, two = cols(), cols(2, 720)
    body = p(text("Regions opening"), spacing=spaced(0, 240), sect=section(one))
    # a: space after above, space before below, the break on text paragraphs.
    body += named("Region a", 3, spacing=spaced(120, 0)) + p(text("Region a end"), spacing=spaced(0, 0),
                                                             sect=section(two, "continuous"))
    body += p(text("Between a"), spacing=spaced(120, 0))
    # b: empty paragraphs holding the breaks, spaced as a Normal of 120 and 120.
    body += p(text("Before b"), spacing=spaced(120, 120)) + p("", spacing=spaced(120, 120),
                                                              sect=section(one, "continuous"))
    body += named("Region b", 4, spacing=spaced(120, 120)) + p("", spacing=spaced(120, 120),
                                                               sect=section(two, "continuous"))
    body += p(text("Between b"), spacing=spaced(120, 120))
    # c: balanced, its second column opening a paragraph with space before; its columns
    # ending in space after.
    body += p("", spacing=spaced(0, 0), sect=section(one, "continuous"))
    body += named("Col c one", 3, spacing=spaced(0, 360)) + named("Col c two", 2, spacing=spaced(240, 120))
    body += p("", spacing=spaced(0, 0), sect=section(two, "continuous"))
    body += p(text("After c"), spacing=spaced(0, 0))
    # d: one paragraph over both columns, its space after under the second.
    body += p("", spacing=spaced(0, 0), sect=section(one, "continuous"))
    body += named("Col d", 4, spacing=spaced(0, 480)) + p("", spacing=spaced(0, 0), sect=section(two, "continuous"))
    body += p(text("After d"), spacing=spaced(0, 0))
    # e: the empty paragraph holding the break spaced 240 after, the next 0 before.
    body += p("", spacing=spaced(0, 0), sect=section(one, "continuous"))
    body += named("Col e", 3, spacing=spaced(0, 0)) + p("", spacing=spaced(0, 240), sect=section(two, "continuous"))
    body += p(text("After e"), spacing=spaced(0, 0), sect=section(one, "continuous"))
    # A new page of two columns: the second column's top opens a paragraph with space
    # before.
    body += fillers("Top", COLUMN_LINES) + p(text("Column top"), spacing=spaced(240, 0)) + fillers("Top after", 3)
    return body, section(two, "nextPage")


def content(kind: str) -> tuple[str, str, list]:
    """``(body, the final section, extra parts)``."""
    if kind == "geometry":
        return (*_geometry(), [])
    if kind == "flow2":
        return (*_flow("Two", 40, cols(2, 720)), [])
    if kind == "flow3":
        return (*_flow("Three", 30, cols(3, 401)), [])
    if kind == "unequal":
        return (*_flow("Unequal", 30, cols(widths=((2200, 400), (4000, 300), (TEXT - 6900, 0)))), [])
    if kind == "keeps":
        return (*_keeps(), [])
    if kind == "breaks":
        return (*_breaks(), [])
    if kind == "sep":
        return (*_sep(), [])
    if kind == "balance":
        return (*_balance(), [])
    if kind == "balancelong":
        return (*_balance_long(), [])
    if kind == "table":
        return (*_table(), [])
    if kind == "picture":
        return (*_picture(), [("word/media/image1.png", "image/png", picture_probe.IMAGE_REL, "")])
    if kind == "header":
        return _header()
    if kind == "footnote":
        return _footnote()
    if kind == "notesright":
        return _notes_right()
    if kind == "notes3":
        return _notes_three()
    if kind == "notesunequal":
        return _notes_unequal()
    if kind == "notesover":
        return _notes_over()
    if kind == "notesregion":
        return _notes_region()
    if kind == "endnotes":
        return _endnotes()
    if kind == "dropcap":
        return (*_drop_cap(), [])
    if kind == "floating":
        return (*_floating(), [("word/media/image1.png", "image/png", picture_probe.IMAGE_REL, "")])
    if kind == "regions":
        return (*_regions(), [])
    raise ValueError(kind)


def parts(name: str) -> tuple[str, str]:
    """``(kind, setting)`` of ``cols-<kind>-<setting>``."""
    _, kind, setting = name.split("-")
    return kind, setting


def build(name: str) -> bytes:
    """``name``: ``cols-<kind>-<setting>``."""
    kind, setting = parts(name)
    body, final, extra = content(kind)
    endnote_styles = [wml.style("paragraph", "EndnoteText", based_on="Normal", ppr_={"spacing": SPACING},
                                rpr_={"sz": 20, "szCs": 20}),
                      wml.style("character", "EndnoteReference", rpr_={"vertAlign": "superscript"})]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "FootnoteText", based_on="Normal", ppr_={"spacing": SPACING},
                   rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "FootnoteReference", rpr_={"vertAlign": "superscript"})]
        + (endnote_styles if kind == "endnotes" else []),
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    mode = SETTINGS[setting]
    if mode is not None:
        extra = extra + [wml.settings_part({"val": "en-GB"}, compatibility_mode=mode)]
    data = probe_docx.package(body, final_section=final, styles=styles, extra_parts=tuple(extra))
    if kind not in ("picture", "floating"):
        return data
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in source.infolist():
            content_ = picture_probe.png() if info.filename == "word/media/image1.png" else source.read(info.filename)
            archive.writestr(info, content_)
    return buffer.getvalue()


DOCUMENTS = tuple(f"cols-{kind}-{setting}" for kind in KINDS for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)
