#!/usr/bin/env python3
"""Tracked changes and content controls: what Word's final view draws.

docx-agent's roadmap found three gaps in the final view docx2svg drew: text inside
``w:moveTo`` was dropped, a paragraph whose mark is deleted was not joined to the next,
and the text of an inline content control was not drawn.  This probe measures what Word
draws instead, exported in its *No Markup* view (``oracle.export(view="final")``; with no
view Word prints the revisions in balloons: ``word_export_pdf.applescript``,
"Revisions"), and the same documents with every revision accepted (``view="accept"``).

Every case is a page: ``Case N``, the case's content, ``After N``.  Calibri 11 pt, no
paragraph spacing, unless the case says otherwise.  Families:

* ``move`` -- ``w:moveFrom`` / ``w:moveTo`` within paragraphs, with their range markers,
  a whole paragraph moved (its mark ``w:moveFrom`` / ``w:moveTo``), the range markers at
  body level, a move with no range markers, formatted moved text, a move between cells;
* ``mark`` -- a deleted paragraph mark (``w:pPr/w:rPr/w:del``) between paragraphs whose
  properties differ (:data:`PA` ... :data:`PD`): content kept or deleted, chains of three,
  paragraph styles and lists on either side, a large mark, before a table (with and
  without content, styled, spaced, the first cell right-aligned), before a block-level
  content control, in a cell and at a cell's end (styled; the next cell centred), an
  inserted mark as the control, and the second paragraph's spacing before and after
  with the join at a line's end, mid-line and in a chain;
* ``sdt`` -- content controls: run-level (plain, showing its placeholder, nested, holding
  revisions, inside a hyperlink), block-level (paragraphs, placeholder, a table, nested),
  cell-level and row-level;
* ``props`` -- property-change records (``w:rPrChange`` on a run and on a mark,
  ``w:pPrChange``, one on a list paragraph, ``w:tblPrChange`` / ``w:trPrChange`` /
  ``w:tcPrChange``) whose old properties differ visibly from the current;
* ``rows`` -- table rows inserted and deleted (``w:trPr/w:ins``, ``w:trPr/w:del``), cells
  inserted and deleted (``w:cellIns``, ``w:cellDel``) and ``w:cellMerge``;
* ``comment`` -- comments, which the final view does not show: a commented range, its
  reference at 8 pt and at 48 pt between words, and a 48 pt reference alone in a paragraph;
* ``section`` -- last, each case its own section: a deleted mark on the paragraph that
  ends a section, the wider margin on either side of the break, ``w:sectPrChange``, and
  the body's last paragraph with its mark deleted.

One document per compatibility setting (none, 14, 15).  Reader:
``read_revision_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import probe_docx
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
WHO = 'w:author="Probe" w:date="2026-01-01T00:00:00Z"'

#: Paragraph properties that tell the joined paragraph's source apart.
PA = {"spacing": {"before": 0, "after": 240, "line": 240, "lineRule": "auto"}, "ind": {"left": 720}, "jc": "center"}
PB = {"spacing": {"before": 120, "after": 0, "line": 240, "lineRule": "auto"}, "ind": {"right": 1440}, "jc": "right"}
PC = {"ind": {"left": 1440, "firstLine": 720}}
PD = {"ind": {"left": 360}, "jc": "both"}

TEXT_A = "Alpha paragraph says its first words, then a few more of them so that it fills a line. "
TEXT_B = "Bravo paragraph follows, with words enough to run across the line and onto another one."
TEXT_C = "Charlie paragraph is the third, and it too is long enough to wrap once on this page. "
TEXT_D = "Delta paragraph ends the chain, and its words carry on past the right margin here."


class _Ids:
    def __init__(self) -> None:
        self.next = 100

    def __call__(self) -> int:
        self.next += 1
        return self.next


ID = _Ids()


def rev(kind: str, content: str) -> str:
    """``content`` wrapped in ``w:ins``, ``w:del``, ``w:moveFrom`` or ``w:moveTo``."""
    return f'<w:{kind} w:id="{ID()}" {WHO}>{content}</w:{kind}>'


def mark(kind: str) -> str:
    """A paragraph mark's (or a row's) revision element: ``<w:del .../>``."""
    return f'<w:{kind} w:id="{ID()}" {WHO}/>'


def run(text: str, **props) -> str:
    return wml.run(text, **props)


def deleted(text: str, **props) -> str:
    """A run as Word writes one inside ``w:del``: its text ``w:delText``."""
    return f'<w:r>{wml.rpr(**props)}<w:delText xml:space="preserve">{probe_docx.escape(text)}</w:delText></w:r>'


def para(content: str = "", *, props: dict | None = None, mark_rev: str | None = None, mark_props: dict | None = None,
         mark_change: str = "", sect: str = "", change: str | None = None) -> str:
    """A ``w:p``: ``props`` its ``w:pPr`` (by element name), ``mark_rev`` a revision on its
    mark (``del``, ``ins``, ``moveFrom``, ``moveTo``), ``mark_props`` the mark's run
    properties, ``mark_change`` a ``w:rPrChange`` on the mark, ``sect`` a ``w:sectPr`` and
    ``change`` a ``w:pPrChange``'s old ``w:pPr`` content (``""``: an empty one)."""
    base = wml._ordered(wml.PPR_ORDER, props or {})
    mark_rpr = ""
    if mark_rev or mark_props or mark_change:
        inner = (mark(mark_rev) if mark_rev else "") + wml._ordered(wml.RPR_ORDER, mark_props or {})
        if mark_change:
            inner += f'<w:rPrChange w:id="{ID()}" {WHO}><w:rPr>{mark_change}</w:rPr></w:rPrChange>'
        mark_rpr = f"<w:rPr>{inner}</w:rPr>"
    pchange = f'<w:pPrChange w:id="{ID()}" {WHO}><w:pPr>{change}</w:pPr></w:pPrChange>' if change is not None else ""
    ppr = base + mark_rpr + sect + pchange
    return f"<w:p>{f'<w:pPr>{ppr}</w:pPr>' if ppr else ''}{content}</w:p>"


def range_marker(kind: str, name: str | None = None, identifier: int | None = None) -> str:
    """``w:moveFromRangeStart`` (with ``name``) or ``...End``."""
    if kind.endswith("Start"):
        return f'<w:{kind} w:id="{identifier}" w:name="{name}" {WHO}/>'
    return f'<w:{kind} w:id="{identifier}"/>'


def sdt(content: str, *, placeholder: bool = False, alias: str = "Control", block: bool = False) -> str:
    """A rich-text content control around ``content`` (runs, or blocks with ``block``)."""
    tag = alias.replace(" ", "")
    pr = f'<w:alias w:val="{alias}"/><w:tag w:val="{tag}"/><w:id w:val="{ID()}"/>'
    if placeholder:
        pr += "<w:showingPlcHdr/>"
    return f"<w:sdt><w:sdtPr>{pr}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"


PLACEHOLDER = "Click or tap here to enter text."


def _side(name: str) -> str:
    return f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'


BORDERS = "<w:tblBorders>" + "".join(_side(n) for n in ("top", "left", "bottom", "right", "insideH", "insideV")) \
    + "</w:tblBorders>"
LOOK = ('<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
        ' w:noHBand="1" w:noVBand="1"/>')


def cell(content: str, width: int = 3000, extra: str = "", change: str = "") -> str:
    """A ``w:tc``: ``extra`` goes after ``w:tcW`` in its ``w:tcPr`` (schema order is the
    caller's), ``change`` is a ``w:tcPrChange``'s old ``w:tcPr`` content."""
    pchange = f'<w:tcPrChange w:id="{ID()}" {WHO}><w:tcPr>{change}</w:tcPr></w:tcPrChange>' if change else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{extra}{pchange}</w:tcPr>{content}</w:tc>'


def row(cells: str, trpr: str = "") -> str:
    return f"<w:tr>{f'<w:trPr>{trpr}</w:trPr>' if trpr else ''}{cells}</w:tr>"


def table(rows: str, columns: int = 2, width: int = 3000, tblpr_change: str = "", grid: tuple | None = None) -> str:
    change = f'<w:tblPrChange w:id="{ID()}" {WHO}><w:tblPr>{tblpr_change}</w:tblPr></w:tblPrChange>' \
        if tblpr_change else ""
    widths = grid or (width,) * columns
    pr = (f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/>{BORDERS}<w:tblLayout w:type="fixed"/>{LOOK}{change}')
    cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{cols}</w:tblGrid>{rows}</w:tbl>"


def text_row(label: str, columns: int = 2) -> str:
    return row("".join(cell(para(run(f"{label} {side}"))) for side in "abc"[:columns]))


@dataclass(frozen=True)
class Case:
    family: str
    name: str
    body: Callable[[int], str]
    #: A case of the ``section`` family ends its own section: its ``After N`` carries the
    #: ``w:sectPr`` this makes (``None`` for the body's, on the last case).
    section: Callable[[], str] | None = None
    #: ``After N`` is centred and its mark deleted (the body's last paragraph).
    last_mark: bool = False

    @property
    def key(self) -> str:
        return f"{self.family} {self.name}"


def section(left: int = 1440, change: str = "") -> str:
    """A next-page ``w:sectPr`` on A4 with this left margin; ``change`` an old one in a
    ``w:sectPrChange``."""
    pchange = f'<w:sectPrChange w:id="{ID()}" {WHO}><w:sectPr>{change}</w:sectPr></w:sectPrChange>' if change else ""
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="{left}" w:header="720" w:footer="720"'
            f' w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/>{pchange}</w:sectPr>')


# -- the cases ----------------------------------------------------------------------------


def _move_inline(n: int) -> str:
    a, b = ID(), ID()
    return (para(run("Moved words start ") + range_marker("moveFromRangeStart", f"move{n}", a)
                 + rev("moveFrom", run("MOVED WORDS ")) + range_marker("moveFromRangeEnd", identifier=a)
                 + run("then the sentence goes on."))
            + para(run("Second paragraph takes ") + range_marker("moveToRangeStart", f"move{n}", b)
                   + rev("moveTo", run("MOVED WORDS ")) + range_marker("moveToRangeEnd", identifier=b)
                   + run("and ends.")))


def _move_paragraph(n: int) -> str:
    a, b = ID(), ID()
    return (para(run("First paragraph stays where it was."))
            + para(range_marker("moveFromRangeStart", f"move{n}", a) + rev("moveFrom", run("Moving paragraph text."))
                   + range_marker("moveFromRangeEnd", identifier=a), props={"ind": {"left": 1440}},
                   mark_rev="moveFrom")
            + para(run("Middle paragraph stays too."), props={"jc": "right"})
            + para(range_marker("moveToRangeStart", f"move{n}", b) + rev("moveTo", run("Moving paragraph text."))
                   + range_marker("moveToRangeEnd", identifier=b), props={"ind": {"left": 1440}},
                   mark_rev="moveTo"))


def _move_body_ranges(n: int) -> str:
    a, b = ID(), ID()
    return (range_marker("moveFromRangeStart", f"move{n}", a)
            + para(rev("moveFrom", run("Paragraph moved down the page.")), props=PC, mark_rev="moveFrom")
            + range_marker("moveFromRangeEnd", identifier=a)
            + para(run("Between the two ends of the move."))
            + range_marker("moveToRangeStart", f"move{n}", b)
            + para(rev("moveTo", run("Paragraph moved down the page.")), props=PC, mark_rev="moveTo")
            + range_marker("moveToRangeEnd", identifier=b))


def _move_bare(n: int) -> str:
    return (para(run("Plain words ") + rev("moveTo", run("ARRIVED")) + run(" then more."))
            + para(run("Other words ") + rev("moveFrom", run("LEFT ")) + run("and the end.")))


def _move_formatted(n: int) -> str:
    return para(run("Before the move ") + rev("moveTo", run("BOLD MOVED", b=True) + run(" and plain moved"))
                + run(" after it."))


def _move_cells(n: int) -> str:
    return table(row(cell(para(run("Cell one ") + rev("moveFrom", run("TRAVELLER ")) + run("stays.")))
                     + cell(para(run("Cell two ") + rev("moveTo", run("TRAVELLER ")) + run("arrives.")))))


def _joined(*parts) -> Callable[[int], str]:
    """Paragraphs ``(text, props, mark revision, content deleted, extra)``."""
    def body(n: int) -> str:
        out = ""
        for text, props, mark_rev, gone, extra in parts:
            content = rev("del", deleted(text)) if gone else run(text)
            if extra.get("break"):
                content += "<w:r><w:br/></w:r>"
            out += para(content, props={**props, **extra.get("props", {})}, mark_rev=mark_rev,
                        mark_props=extra.get("mark"))
        return out
    return body


def _numbered() -> dict:
    return {"numPr": '<w:ilvl w:val="0"/><w:numId w:val="1"/>'}


def _mark_before_table(gone: bool, props: dict = PA, first: dict | None = None) -> Callable[[int], str]:
    """A paragraph whose mark is deleted before a table; ``first`` the ``w:pPr`` of the
    table's first cell's paragraph."""
    def body(n: int) -> str:
        content = rev("del", deleted(TEXT_A)) if gone else run(TEXT_A)
        top = row(cell(para(run(f"{n}.t1 a"), props=first or {})) + cell(para(run(f"{n}.t1 b"))))
        return para(content, props=props, mark_rev="del") + table(top + text_row(f"{n}.t2"))
    return body


def _mark_before_sdt(n: int) -> str:
    return para(run(TEXT_A), props=PA, mark_rev="del") + sdt(para(run(TEXT_B), props=PB), block=True,
                                                            alias="Block after")


def _mark_in_cell(n: int) -> str:
    inner = para(run("Cell first part, "), props={"jc": "center"}, mark_rev="del") \
        + para(run("cell second part."), props={"jc": "right"})
    return table(row(cell(inner, 4000) + cell(para(run(f"{n}.side")), 4000)), grid=(4000, 4000))


def _mark_cell_end(props: dict | None = None, following: dict | None = None) -> Callable[[int], str]:
    """The last paragraph of a cell with its mark deleted (``props``, centred by default);
    ``following`` the next cell's paragraph's properties."""
    def body(n: int) -> str:
        inner = para(run("Cell first paragraph.")) + para(run("Cell last paragraph."),
                                                           props=props or {"jc": "center"}, mark_rev="del")
        return table(row(cell(inner, 4000) + cell(para(run(f"{n}.next"), props=following or {}), 4000)),
                     grid=(4000, 4000))
    return body


def _sdt_run(n: int) -> str:
    return para(run("Before the control ") + sdt(run("INSIDE THE CONTROL")) + run(" after it."))


def _sdt_placeholder(n: int) -> str:
    return para(run("Name: ") + sdt(run(PLACEHOLDER, rStyle="PlaceholderText"), placeholder=True) + run(" end."))


def _sdt_nested(n: int) -> str:
    inner = sdt(run("INNER"), alias="Inner")
    return para(run("Start ") + sdt(run("outer one ") + inner + run(" outer two"), alias="Outer") + run(" stop."))


def _sdt_revisions(n: int) -> str:
    return para(run("Lead ") + sdt(rev("ins", run("ADDED ")) + rev("del", deleted("REMOVED ")) + run("kept"))
                + run(" tail."))


def _sdt_hyperlink(n: int) -> str:
    return para(run("See ") + f'<w:hyperlink w:anchor="target">{sdt(run("LINKED CONTROL"))}</w:hyperlink>'
                + run(" here."))


def _sdt_block(n: int) -> str:
    return sdt(para(run("Block control first paragraph.")) + para(run("Block control second paragraph."),
                                                                  props={"jc": "center"}), block=True)


def _sdt_block_placeholder(n: int) -> str:
    return sdt(para(run(PLACEHOLDER, rStyle="PlaceholderText")), placeholder=True, block=True)


def _sdt_block_table(n: int) -> str:
    return sdt(table(text_row(f"{n}.r1") + text_row(f"{n}.r2")), block=True)


def _sdt_block_nested(n: int) -> str:
    return sdt(para(run("Outer block.")) + sdt(para(run("Inner block.")), block=True, alias="Inner"),
               block=True, alias="Outer")


def _sdt_cell(n: int) -> str:
    return table(row(cell(sdt(para(run("Cell control text.")), block=True)) + cell(para(run(f"{n}.plain")))))


def _sdt_row(n: int) -> str:
    return table(text_row(f"{n}.r1") + sdt(text_row(f"{n}.r2"), block=True) + text_row(f"{n}.r3"))


def _rpr_change(n: int) -> str:
    changed = (f'<w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/>'
               f'<w:rPrChange w:id="{ID()}" {WHO}><w:rPr/></w:rPrChange></w:rPr>'
               '<w:t xml:space="preserve">Bold and large now</w:t></w:r>')
    return para(run("Plain before ") + changed + run(" plain after."))


def _mark_rpr_change(n: int) -> str:
    return para(run("A paragraph whose mark is large."), mark_props={"sz": 72, "szCs": 72},
                mark_change='<w:sz w:val="22"/><w:szCs w:val="22"/>') + para(run("The next paragraph."))


def _ppr_change(n: int) -> str:
    return para(run(TEXT_A + TEXT_B), props={"ind": {"left": 1440}, "jc": "center"},
                change='<w:ind w:left="0"/><w:jc w:val="right"/>')


def _ppr_change_list(n: int) -> str:
    return para(run("A list paragraph now."), props=_numbered(), change="") + \
        para(run("No longer in a list."), change='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')


def _table_changes(n: int) -> str:
    rows = (row(cell(para(run(f"{n}.a1")), extra='<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
                     change='<w:tcW w:w="1500" w:type="dxa"/>') + cell(para(run(f"{n}.b1"))),
                trpr=f'<w:trHeight w:val="1200" w:hRule="exact"/><w:trPrChange w:id="{ID()}" {WHO}><w:trPr>'
                     '<w:trHeight w:val="300" w:hRule="exact"/></w:trPr></w:trPrChange>')
            + text_row(f"{n}.r2"))
    return table(rows, tblpr_change='<w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="center"/>')


def _row(label: str, kind: str | None, content_tracked: bool = True) -> str:
    def text(t):
        if kind == "del" and content_tracked:
            return rev("del", deleted(t))
        if kind == "ins" and content_tracked:
            return rev("ins", run(t))
        return run(t)
    cells = "".join(cell(para(text(f"{label} {side}"), mark_rev=kind if content_tracked else None)) for side in "ab")
    return row(cells, trpr=mark(kind) if kind else "")


def _rows(*kinds, content_tracked: bool = True) -> Callable[[int], str]:
    def body(n: int) -> str:
        return table("".join(_row(f"{n}.r{i + 1}", kind, content_tracked) for i, kind in enumerate(kinds)))
    return body


def _cell_revision(kind: str) -> Callable[[int], str]:
    def body(n: int) -> str:
        if kind == "cellIns":
            middle = cell(para(rev("ins", run(f"{n}.new")), mark_rev="ins"), extra=mark("cellIns"))
        else:
            middle = cell(para(rev("del", deleted(f"{n}.gone")), mark_rev="del"), extra=mark("cellDel"))
        rows = row(cell(para(run(f"{n}.a1"))) + middle + cell(para(run(f"{n}.c1")))) + text_row(f"{n}.r2", 3)
        return table(rows, columns=3)
    return body


def _cell_merge(vmerge: bool) -> Callable[[int], str]:
    def body(n: int) -> str:
        merge = f'<w:cellMerge w:id="{ID()}" {WHO} w:vMerge="cont" w:vMergeOrig="rest"/>'
        top = row(cell(para(run(f"{n}.a1")), extra='<w:vMerge w:val="restart"/>') + cell(para(run(f"{n}.b1"))))
        below = row(cell(para(run(f"{n}.a2")) if not vmerge else para(), extra=('<w:vMerge/>' if vmerge else "") + merge)
                    + cell(para(run(f"{n}.b2"))))
        return table(top + below + text_row(f"{n}.r3"))
    return body


COMMENTS: list[int] = []


def comment_reference(size: int | None = None) -> str:
    """A comment's reference run, as Word writes it (character style ``CommentReference``);
    the comment is added to :data:`COMMENTS`."""
    COMMENTS.append(len(COMMENTS))
    props = {"rStyle": "CommentReference"}
    if size:
        props.update(sz=size, szCs=size)
    return f'<w:r>{wml.rpr(**props)}<w:commentReference w:id="{COMMENTS[-1]}"/></w:r>'


def _comment_range(n: int) -> str:
    k = len(COMMENTS)
    return para(run("Commented ") + f'<w:commentRangeStart w:id="{k}"/>' + run("words in a range")
                + f'<w:commentRangeEnd w:id="{k}"/>' + comment_reference() + run(" and the end."))


def _comment_large(n: int) -> str:
    return para(run("Before a large reference") + comment_reference(96) + run(" after it."))


def _comment_alone(n: int) -> str:
    return para(comment_reference(96)) + para(run("Below the reference's paragraph."))


def comments_part() -> tuple[str, str, str, str]:
    """``word/comments.xml`` with every comment :data:`COMMENTS` holds."""
    body = "".join(f'<w:comment w:id="{k}" {WHO} w:initials="P"><w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/>'
                   f'</w:rPr><w:annotationRef/></w:r><w:r><w:t>Comment {k}.</w:t></w:r></w:p></w:comment>'
                   for k in COMMENTS)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<w:comments xmlns:w="{probe_docx.W_NS}">{body}</w:comments>')
    return ("word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", xml)


def _section_mark(n: int) -> str:
    return para(run("Section end paragraph whose mark is deleted. "), props=PA, mark_rev="del",
                sect=section(2880)) + para(run(TEXT_B), props=PB)


def _section_mark_after(n: int) -> str:
    return para(run("Section end paragraph whose mark is deleted. "), props=PA, mark_rev="del",
                sect=section()) + para(run(TEXT_B), props=PB)


#: The old page geometry of the ``sectPrChange`` case: a left margin of 2880.
OLD_MARGINS = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="2880" w:header="720" w:footer="720" w:gutter="0"/>')


def _section_change(n: int) -> str:
    return para(run("A section whose margins changed."))


def _cases() -> tuple[Case, ...]:
    out = [
        Case("move", "inline with ranges", _move_inline),
        Case("move", "paragraph", _move_paragraph),
        Case("move", "body-level ranges", _move_body_ranges),
        Case("move", "no ranges", _move_bare),
        Case("move", "formatted", _move_formatted),
        Case("move", "between cells", _move_cells),
        Case("mark", "content kept", _joined((TEXT_A, PA, "del", False, {}), (TEXT_B, PB, None, False, {}))),
        Case("mark", "content deleted", _joined((TEXT_A, PA, "del", True, {}), (TEXT_B, PB, None, False, {}))),
        Case("mark", "chain of three", _joined((TEXT_A, PA, "del", False, {}), (TEXT_B, PB, "del", False, {}),
                                               (TEXT_C, PC, "del", False, {}), (TEXT_D, PD, None, False, {}))),
        Case("mark", "chain of three deleted", _joined((TEXT_A, PA, "del", True, {}), (TEXT_B, PB, "del", True, {}),
                                                       (TEXT_C, PC, "del", True, {}), (TEXT_D, PD, None, False, {}))),
        Case("mark", "chain, middle kept", _joined((TEXT_A, PA, "del", True, {}), (TEXT_B, PB, "del", False, {}),
                                                   (TEXT_C, PC, None, False, {}))),
        Case("mark", "styled first", _joined((TEXT_A, {"pStyle": "Big"}, "del", False, {}),
                                             (TEXT_B, PB, None, False, {}))),
        Case("mark", "styled second", _joined((TEXT_A, PA, "del", False, {}),
                                              (TEXT_B, {"pStyle": "Big"}, None, False, {}))),
        Case("mark", "list first", _joined((TEXT_A, _numbered(), "del", False, {}), (TEXT_B, PB, None, False, {}))),
        Case("mark", "list second", _joined((TEXT_A, PA, "del", False, {}), (TEXT_B, _numbered(), None, False, {}))),
        Case("mark", "large mark", _joined((TEXT_A, {}, "del", False, {"mark": {"sz": 72, "szCs": 72}}),
                                           (TEXT_B, {}, None, False, {}))),
        Case("mark", "large mark kept", _joined((TEXT_A, {}, "del", False, {}),
                                                (TEXT_B, {}, None, False, {"mark": {"sz": 72, "szCs": 72}}))),
        Case("mark", "before a table", _mark_before_table(False)),
        Case("mark", "deleted before a table", _mark_before_table(True)),
        Case("mark", "before a block control", _mark_before_sdt),
        Case("mark", "in a cell", _mark_in_cell),
        Case("mark", "at a cell's end", _mark_cell_end()),
        Case("mark", "inserted", _joined((TEXT_A, PA, "ins", False, {}), (TEXT_B, PB, None, False, {}))),
        # Which properties a paragraph with no paragraph to join takes.
        Case("mark", "before a table, first cell right", _mark_before_table(False, first=PB)),
        Case("mark", "before a table, styled", _mark_before_table(False, props={"pStyle": "Big"})),
        Case("mark", "before a table, spaced", _mark_before_table(False, props=PB)),
        Case("mark", "at a cell's end, next centred", _mark_cell_end(PB, {"jc": "center"})),
        Case("mark", "at a cell's end, styled", _mark_cell_end({"pStyle": "Big"})),
        # Where the second paragraph's spacing before goes below mode 15.
        Case("mark", "short first", _joined(("Short alpha. ", PA, "del", False, {}), (TEXT_B, PB, None, False, {}))),
        Case("mark", "second spaced 480", _joined((TEXT_A, PA, "del", False, {}),
                                                  (TEXT_B, {"spacing": {"before": 480, "after": 0}}, None, False,
                                                   {}))),
        Case("mark", "chain, last spaced", _joined((TEXT_A, PA, "del", False, {}), (TEXT_B, PB, "del", False, {}),
                                                   (TEXT_C, {"spacing": {"before": 240, "after": 0}}, None, False,
                                                    {}))),
        Case("mark", "second spaced after", _joined((TEXT_A, PA, "del", False, {}),
                                                    (TEXT_B, {"spacing": {"before": 120, "after": 480}}, None,
                                                     False, {}))),
        Case("mark", "break, second spaced", _joined(("Short alpha.", PA, "del", False, {"break": True}),
                                                     (TEXT_B, PB, None, False, {}))),
        Case("mark", "break, second spaced 480", _joined(("Short alpha.", PA, "del", False, {"break": True}),
                                                         (TEXT_B, {**PB, "spacing": {"before": 480, "after": 0}},
                                                          None, False, {}))),
        Case("mark", "two lines and a break, second spaced",
             _joined((TEXT_A + "Two more.", PA, "del", False, {"break": True}), (TEXT_B, PB, None, False, {}))),
        Case("mark", "chain, breaks, last spaced", _joined(("Short alpha.", PA, "del", False, {"break": True}),
                                                           ("Short bravo.", PB, "del", False, {"break": True}),
                                                           (TEXT_C, {"spacing": {"before": 240, "after": 0}}, None,
                                                            False, {}))),
        Case("mark", "chain of four, breaks, last spaced",
             _joined(("Short alpha.", PA, "del", False, {"break": True}),
                     ("Short bravo.", PB, "del", False, {"break": True}),
                     ("Short charlie.", PC, "del", False, {"break": True}),
                     (TEXT_D, {"spacing": {"before": 240, "after": 0}}, None, False, {}))),
        Case("mark", "second spaced, first mid-line", _joined((TEXT_A + "Two more. ", PA, "del", False, {}),
                                                              (TEXT_B, PB, None, False, {}))),
        Case("sdt", "run", _sdt_run),
        Case("sdt", "run placeholder", _sdt_placeholder),
        Case("sdt", "run nested", _sdt_nested),
        Case("sdt", "run with revisions", _sdt_revisions),
        Case("sdt", "run in a hyperlink", _sdt_hyperlink),
        Case("sdt", "block", _sdt_block),
        Case("sdt", "block placeholder", _sdt_block_placeholder),
        Case("sdt", "block table", _sdt_block_table),
        Case("sdt", "block nested", _sdt_block_nested),
        Case("sdt", "cell", _sdt_cell),
        Case("sdt", "row", _sdt_row),
        Case("props", "rPrChange", _rpr_change),
        Case("props", "mark rPrChange", _mark_rpr_change),
        Case("props", "pPrChange", _ppr_change),
        Case("props", "pPrChange on a list", _ppr_change_list),
        Case("props", "table, row and cell changes", _table_changes),
        Case("rows", "middle row deleted", _rows(None, "del", None)),
        Case("rows", "middle row inserted", _rows(None, "ins", None)),
        Case("rows", "first row deleted", _rows("del", None, None)),
        Case("rows", "every row deleted", _rows("del", "del")),
        Case("rows", "row deleted, content not", _rows(None, "del", None, content_tracked=False)),
        Case("rows", "cell inserted", _cell_revision("cellIns")),
        Case("rows", "cell deleted", _cell_revision("cellDel")),
        Case("rows", "cellMerge with vMerge", _cell_merge(True)),
        Case("rows", "cellMerge alone", _cell_merge(False)),
        Case("comment", "range", _comment_range),
        Case("comment", "large reference", _comment_large),
        Case("comment", "reference alone", _comment_alone),
        # The section family: each case ends its own section (the one before them too).
        Case("section", "deleted mark, wider margin before", _section_mark, section=section),
        Case("section", "deleted mark, wider margin after", _section_mark_after, section=lambda: section(2880)),
        Case("section", "sectPrChange", _section_change, section=lambda: section(change=OLD_MARGINS)),
        Case("section", "the body's last mark deleted", lambda n: para(run("The body's last paragraph follows.")),
             last_mark=True),
    ]
    return tuple(out)


CASES = _cases()
#: The last case before the ``section`` family ends a section of its own.
FIRST_SECTION = next(i for i, case in enumerate(CASES) if case.family == "section")


def case_body(number: int, case: Case) -> str:
    # After a section break the case is on a new page already.
    props = {"pageBreakBefore": True} if number and number != FIRST_SECTION and case.family != "section" else {}
    if case.family == "section" and number > FIRST_SECTION:
        props = {}
    out = para(run(f"Case {number}"), props=props)
    out += case.body(number)
    sect = ""
    if number == FIRST_SECTION - 1:
        sect = section()
    elif case.family == "section" and case.section is not None:
        sect = case.section()
    if case.last_mark:
        return out + para(run(f"After {number}"), props={"jc": "center"}, mark_rev="del")
    return out + para(run(f"After {number}"), sect=sect)


def body() -> str:
    ID.next = 100
    COMMENTS.clear()
    return "".join(case_body(number, case) for number, case in enumerate(CASES))


def final_section() -> str:
    """The body's ``w:sectPr``: the last case's."""
    return section()


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "Big", name="Big", based_on="Normal",
                   ppr_={"spacing": {"before": 240, "after": 120, "line": 240, "lineRule": "auto"}},
                   rpr_={"b": True, "sz": 32, "szCs": 32}),
         wml.style("character", "PlaceholderText", name="Placeholder Text", rpr_={"color": "808080"}),
         wml.style("character", "CommentReference", name="annotation reference", rpr_={"sz": 16, "szCs": 16})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    numbering = wml.numbering_part([("decimal", "%1.", {"ind": {"left": 720, "hanging": 360}}, {})])
    document = body()
    extra = (numbering, comments_part())
    if mode is not None:
        extra += (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(document, final_section=final_section(), styles=styles(), extra_parts=extra)


DOCUMENTS = tuple(f"revision-{setting}" for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"revision-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
    print(len(CASES), "cases")
