"""Made-up reports of many pages, for the layout's page <-> block map and for laying a
document out again from a previous layout (round 19).

A report is a list of :class:`Block` values -- headings, paragraphs, a bulleted list, tables
with a header row, side panels (a text box anchored in its paragraph, text wrapping beside
it), page breaks -- that :func:`package` writes as a ``.docx`` with a header (the report's
name) and a footer (``Page N of M``: ``PAGE`` and ``NUMPAGES``).  :func:`report` makes one of
about ``pages`` A4 pages; the edits a test applies (:func:`edit`) change the list, so the
same document can be written before and after.  Every word is made up; the face is Calibri
in every slot (Carlito where Calibri is absent, as CI installs it).
"""

from __future__ import annotations

import io
import random
import zipfile
import zlib
from dataclasses import dataclass, replace

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."

WORDS = ("river stone lantern orchard meadow harbour quiet copper window garden bridge "
         "morning letter valley ribbon candle signal forest compass marble season "
         "travel number simple yellow corner pocket silver little gentle winter "
         "summer button kettle ladder pebble basket feather island market puzzle").split()


@dataclass(frozen=True)
class Block:
    """One body block: ``kind`` is ``heading``, ``para``, ``bullet``, ``table``, ``panel``
    (a paragraph anchoring a side text box), ``break`` (a paragraph holding a page break),
    ``noted`` (a paragraph with a footnote, whose text is ``note``), ``lead`` (a paragraph
    that keeps with the next) or ``raw`` (``text`` is the block's XML, written as it is); ``text`` its words (a table's: ``rows`` of cells).

    With fields (round 20): ``toc`` (a line of a table of contents: its heading's words, a
    dotted tab and ``PAGEREF target``; the first ``opens`` the ``TOC`` field), ``tocend``
    (the paragraph closing it), ``caption`` (``Table {SEQ Table}: text``, its label and
    number bookmarked as ``mark``), ``xref`` (``text`` and ``See {REF target}.``),
    ``pageref`` (``text (see page {PAGEREF target})``) and ``pagefield`` (``This paragraph
    is on page {PAGE}.`` and ``text``)."""

    kind: str
    text: str = ""
    rows: tuple = ()
    #: ``w14:paraId`` of a paragraph, as Word stamps it.
    para_id: str = ""
    #: A ``noted`` paragraph's footnote.
    note: str = ""
    #: A bookmark: around a ``heading``'s text, or a ``caption``'s label and number.
    mark: str = ""
    #: The bookmark a ``toc`` line's ``PAGEREF``, an ``xref``'s ``REF`` or a ``pageref``'s
    #: ``PAGEREF`` points at.
    target: str = ""
    #: A ``toc`` line that opens the ``TOC`` field (the first).
    opens: bool = False
    #: A table's rows that keep with the next (round 24, :func:`report`'s ``keeps``): ``all``,
    #: ``but last``, ``some`` (every other one) or ``label`` (all, and no header row).
    keep: str = ""


def _words(rng: random.Random, n: int) -> str:
    out = " ".join(rng.choice(WORDS) for _ in range(n))
    return out[0].upper() + out[1:] + "."


def _sentences(rng: random.Random, low: int, high: int) -> str:
    return " ".join(_words(rng, rng.randint(6, 14)) for _ in range(rng.randint(low, high)))


#: A paragraph ending a section: the next starts on a new page (``nextPage``) or on this one.
def section_break(kind: str = "nextPage") -> Block:
    return Block("raw", f'<w:p><w:pPr><w:sectPr><w:type w:val="{kind}"/><w:pgSz w:w="11906" w:h="16838"/>'
                        '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" '
                        'w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>')


def report(pages: int, seed: int = 0, *, notes: bool = False, sections: bool = False, toc: bool = False,
           toc_break: bool = True, fields: bool = False, keeps: bool = False) -> list[Block]:
    """A made-up report of about ``pages`` A4 pages: per part a heading, paragraphs, a
    bulleted list, a side panel, and every other part a table; every sixth part starts a
    page (a paragraph holding a page break before its heading).  ``notes``: every third
    part has a paragraph with a footnote too; ``sections``: every seventh part ends a
    section, the next starting on a new page, or (every other one) on the same page.

    ``toc``: a title page, then a table of contents -- a ``TOC`` field whose lines each
    hold a ``PAGEREF`` to a heading's bookmark: 12 of them from 40 pages, 7 from 16, fewer
    in a shorter report -- and (``toc_break``) a page break after it.  ``fields``: each table
    has a caption numbered by ``SEQ Table`` before it, every third part a paragraph
    referring to the last caption (``REF``), every fifth a paragraph showing a heading's
    page (``PAGEREF``) and every seventh one showing its own (``PAGE``).  ``keeps`` (round
    24): every table's rows keep with the next -- all of them, all but the last, every other
    one, or all with no header row, in turn -- tables of up to 16 rows, and every other
    one after a ``lead`` paragraph that keeps with the next."""
    rng = random.Random(seed)
    blocks: list[Block] = []
    serial = 0

    def pid() -> str:
        nonlocal serial
        serial += 1
        return f"{0x10000000 + serial * 7919:08X}"

    parts = pages * 8 // 5
    entries = 12 if pages >= 40 else 7 if pages >= 16 else min(parts, 4)
    step = -(-parts // entries)
    headings = [part for part in range(parts) if part % step == 0]
    if toc:
        blocks.append(Block("para", "Made-up quarterly report", para_id=pid()))
        blocks.append(Block("break", para_id=pid()))
        blocks.append(Block("para", "Contents", para_id=pid()))
        for k, part in enumerate(headings):
            blocks.append(Block("toc", "", para_id=pid(), target=f"_Toc{part + 1:04d}", opens=k == 0))
        blocks.append(Block("tocend", para_id=pid()))
        if toc_break:
            blocks.append(Block("break", para_id=pid()))
    captions = 0
    for part in range(parts):
        if part % 6 == 5:
            # A chapter starts a page.
            blocks.append(Block("break", para_id=pid()))
        mark = f"_Toc{part + 1:04d}" if toc or fields else ""
        blocks.append(Block("heading", f"Part {part + 1}: " + _words(rng, 3)[:-1], para_id=pid(), mark=mark))
        for _ in range(rng.randint(2, 4)):
            blocks.append(Block("para", _sentences(rng, 2, 6), para_id=pid()))
        if part % 3 == 1:
            blocks.append(Block("panel", _sentences(rng, 3, 5), para_id=pid()))
        if notes and part % 3 == 2:
            blocks.append(Block("noted", _sentences(rng, 1, 3), para_id=pid(), note=_sentences(rng, 1, 2)))
        for _ in range(rng.randint(0, 3)):
            blocks.append(Block("bullet", _words(rng, rng.randint(4, 12)), para_id=pid()))
        if part % 2 == 0:
            rows = tuple(tuple(_words(rng, rng.randint(1, 6 if r else 2)) for _ in range(3))
                         for r in range(rng.randint(3, 16 if keeps else 7)))
            if fields:
                captions += 1
                blocks.append(Block("caption", _words(rng, 4)[:-1], para_id=pid(), mark=f"_Ref{captions:04d}"))
            keep = KEEPS[part // 2 % len(KEEPS)] if keeps else ""
            if keeps and part % 4 == 0:
                blocks.append(Block("lead", _words(rng, 6), para_id=pid()))
            blocks.append(Block("table", rows=rows, keep=keep))
        if fields and captions and part % 3 == 2:
            blocks.append(Block("xref", _sentences(rng, 1, 2), para_id=pid(), target=f"_Ref{captions:04d}"))
        if fields and part % 5 == 4:
            blocks.append(Block("pageref", _sentences(rng, 1, 2), para_id=pid(),
                                target=f"_Toc{rng.randrange(parts) + 1:04d}"))
        if fields and part % 7 == 3:
            blocks.append(Block("pagefield", _sentences(rng, 1, 2), para_id=pid()))
        blocks.append(Block("para", _sentences(rng, 1, 4), para_id=pid()))
        if sections and part % 7 == 6:
            blocks.append(section_break("nextPage" if part % 14 == 6 else "continuous"))
    if toc:
        # Each line of the contents shows its heading's words.
        words = {block.mark: block.text for block in blocks if block.kind == "heading" and block.mark}
        blocks = [replace(block, text=words.get(block.target, "")) if block.kind == "toc" else block
                  for block in blocks]
    return blocks


# -- edits -------------------------------------------------------------------------------------


#: Every kind of edit :func:`edit` makes.
EDIT_KINDS = ("text", "grow", "cell", "insert", "delete", "row-insert", "row-delete", "break", "unbreak")
#: The edits of a report with fields (:func:`report`'s ``toc`` and ``fields``): a caption
#: inserted or deleted (the ``SEQ`` numbers after it change, and the ``REF`` to them), a line
#: of the contents inserted or deleted.
FIELD_EDIT_KINDS = ("caption-insert", "caption-delete", "toc-insert", "toc-delete")
#: The edits of a report whose header shows a ``STYLEREF`` (round 21): a paragraph made a
#: heading (its first words), or a heading a paragraph (its style changed).
STYLE_EDIT_KINDS = ("restyle",)
#: The blocks whose words an edit changes.
_WORDED = ("heading", "para", "bullet", "panel", "noted", "caption", "xref", "pageref", "pagefield", "lead")
#: How a table's rows keep with the next, in turn (:func:`report`'s ``keeps``).
KEEPS = ("all", "but last", "some", "label")


def _in_contents(blocks: list[Block], k: int) -> bool:
    """Whether a block inserted at ``k`` would stand inside the contents' ``TOC`` field."""
    return 0 < k < len(blocks) and blocks[k - 1].kind == "toc" and blocks[k].kind in ("toc", "tocend")


def edit(blocks: list[Block], rng: random.Random, kind: str | None = None,
         kinds: tuple | None = None) -> tuple[list[Block], str]:
    """One random edit of ``blocks``: change a paragraph's or a cell's text, insert or delete
    a paragraph, a table row or a page break.  ``(new blocks, what was done)``.  ``kind``
    (one of :data:`EDIT_KINDS` or :data:`FIELD_EDIT_KINDS`) makes an edit of that kind only;
    ``unbreak`` then needs a page break in ``blocks``, ``cell``, ``row-insert`` and
    ``row-delete`` a table, ``caption-delete`` a caption and ``toc-insert`` and
    ``toc-delete`` contents.  ``kinds``: the kinds to choose from (each as often as it is
    named), else the usual mix.  The contents' field is never split or emptied."""
    blocks = list(blocks)
    every = EDIT_KINDS + FIELD_EDIT_KINDS + STYLE_EDIT_KINDS
    if kind is not None and kind not in every:
        raise ValueError(f"no edit of kind {kind!r}: one of {', '.join(every)}")
    kinds = [kind] if kind else list(kinds or ["text", "text", "cell", "insert", "delete", "row-insert",
                                               "row-delete", "break", "unbreak", "grow"])
    while True:
        kind = rng.choice(kinds)
        k = rng.randrange(len(blocks))
        block = blocks[k]
        if kind in ("text", "grow") and block.kind in _WORDED:
            text = _words(rng, rng.randint(3, 10)) if kind == "text" and rng.random() < 0.5 else (
                block.text + " " + _sentences(rng, 1, 3 if kind == "grow" else 1))
            blocks[k] = replace(block, text=text)
            return blocks, f"{kind} {k}"
        if kind in ("cell", "row-insert", "row-delete") and block.kind == "table":
            rows = [list(row) for row in block.rows]
            r = rng.randrange(1, len(rows))
            if kind == "cell":
                rows[r][rng.randrange(3)] = _words(rng, rng.randint(1, 20))
            elif kind == "row-insert":
                rows.insert(r, [_words(rng, rng.randint(1, 8)) for _ in range(3)])
            elif len(rows) > 2:
                del rows[r]
            else:
                continue
            blocks[k] = replace(block, rows=tuple(tuple(row) for row in rows))
            return blocks, f"{kind} {k}.{r}"
        if kind in ("insert", "break", "caption-insert") and _in_contents(blocks, k):
            continue
        if kind == "insert":
            blocks.insert(k, Block("para", _sentences(rng, 1, 3), para_id=f"7{rng.randrange(1 << 28):07X}"))
            return blocks, f"insert {k}"
        if kind == "delete" and len(blocks) > 2 and block.kind not in ("toc", "tocend"):
            del blocks[k]
            return blocks, f"delete {k}"
        if kind == "break":
            blocks.insert(k, Block("break", para_id=f"6{rng.randrange(1 << 28):07X}"))
            return blocks, f"break {k}"
        if kind == "unbreak":
            breaks = [i for i, b in enumerate(blocks) if b.kind == "break"]
            if breaks:
                k = rng.choice(breaks)
                del blocks[k]
                return blocks, f"unbreak {k}"
        if kind == "restyle" and block.kind in ("heading", "para") and not block.mark:
            # A paragraph made a heading keeps its first words, as a heading's.
            text = block.text if block.kind == "heading" else " ".join(block.text.split()[:5])
            blocks[k] = replace(block, kind="para" if block.kind == "heading" else "heading", text=text)
            return blocks, f"restyle {k}"
        if kind == "caption-insert":
            mark = f"_Ref{rng.randrange(1 << 20) + 5000:07d}"
            blocks.insert(k, Block("caption", _words(rng, 4)[:-1], para_id=f"5{rng.randrange(1 << 28):07X}",
                                   mark=mark))
            return blocks, f"caption-insert {k}"
        if kind == "caption-delete":
            found = [i for i, b in enumerate(blocks) if b.kind == "caption"]
            if found:
                k = rng.choice(found)
                del blocks[k]
                return blocks, f"caption-delete {k}"
        if kind == "toc-insert":
            lines = [i for i, b in enumerate(blocks) if b.kind == "toc"]
            marks = [b for b in blocks if b.kind == "heading" and b.mark]
            if lines and marks:
                k = rng.choice(lines) + 1
                heading = rng.choice(marks)
                blocks.insert(k, Block("toc", heading.text, para_id=f"4{rng.randrange(1 << 28):07X}",
                                       target=heading.mark))
                return blocks, f"toc-insert {k}"
        if kind == "toc-delete":
            lines = [i for i, b in enumerate(blocks) if b.kind == "toc" and not b.opens]
            if lines:
                k = rng.choice(lines)
                del blocks[k]
                return blocks, f"toc-delete {k}"


# -- writing -----------------------------------------------------------------------------------


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'


def _para(text: str, style: str | None = None, para_id: str = "", extra_ppr: str = "", runs: str = "") -> str:
    ident = f' w14:paraId="{para_id}"' if para_id else ""
    ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + extra_ppr
    return f'<w:p{ident}>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{runs or (_run(text) if text else "")}</w:p>'


def _panel(text: str, number: int) -> str:
    """A side panel: a 2-inch wide text box at the column's right edge, text wrapping
    beside it (``wrapSquare``), anchored in its paragraph."""
    cx, cy = 1828800, 1371600
    box = (f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{WPS_NS}" xmlns:wps="{WPS_NS}">'
           f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
           '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EEF2F7"/>'
           '</a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="5B7DB1"/></a:solidFill></a:ln></wps:spPr>'
           f'<wps:txbx><w:txbxContent>{_para("Side note " + str(number), "Panel")}{_para(text[:120], "Panel")}'
           '</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" '
           'rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
           '</a:graphicData></a:graphic>')
    anchor = (f'<w:r><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}"><wp:anchor distT="45720" distB="45720" '
              f'distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + number}" behindDoc="0" '
              'locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
              '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
              '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
              f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
              f'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{number}" name="Text Box {number}"/>'
              f'<wp:cNvGraphicFramePr/>{box}</wp:anchor></w:drawing></w:r>')
    return anchor


def _kept(keep: str, r: int, rows: int) -> bool:
    """Whether row ``r`` of ``rows`` keeps with the next under ``keep`` (:attr:`Block.keep`)."""
    return keep in ("all", "label") or (keep == "but last" and r < rows - 1) or (keep == "some" and r % 2 == 1)


def _table(rows: tuple, keep: str = "") -> str:
    width = 9000
    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
           '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'
           + "".join(f'<w:gridCol w:w="{width // 3}"/>' for _ in range(3)) + "</w:tblGrid>")
    for r, row in enumerate(rows):
        first = r == 0 and keep != "label"
        header = "<w:trPr><w:tblHeader/></w:trPr>" if first else ""
        extra = "<w:keepNext/>" if _kept(keep, r, len(rows)) else ""
        cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{width // 3}" w:type="dxa"/>'
                        + ('<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if first else "")
                        + f'</w:tcPr>{_para(text, "Cell", extra_ppr=extra)}</w:tc>' for text in row)
        out += f"<w:tr>{header}{cells}</w:tr>"
    return out + "</w:tbl>"


def _marked(runs: str, name: str) -> str:
    """``runs`` inside a bookmark ``name`` (none where ``name`` is empty), its id found from
    its name: a bookmark inserted before it does not change it, as in Word."""
    if not name:
        return runs
    number = zlib.crc32(name.encode()) % 1000000
    return f'<w:bookmarkStart w:id="{number}" w:name="{name}"/>{runs}<w:bookmarkEnd w:id="{number}"/>'


def body_xml(blocks: list[Block]) -> str:
    out = ""
    panels = 0
    notes = 0
    for block in blocks:
        if block.kind == "heading":
            out += _para("", "Heading1", block.para_id, runs=_marked(_run(block.text), block.mark)) \
                if block.mark else _para(block.text, "Heading1", block.para_id)
        elif block.kind == "toc":
            opening = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> '
                       'TOC \\o "1-1" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/>'
                       '</w:r>') if block.opens else ""
            number = _field(f"PAGEREF {block.target} \\h", "1")
            link = (f'<w:hyperlink w:anchor="{block.target}" w:history="1">{_run(block.text)}<w:r><w:tab/></w:r>'
                    f'{number}</w:hyperlink>')
            out += _para("", "TOC1", block.para_id, runs=opening + link)
        elif block.kind == "tocend":
            out += _para("", None, block.para_id, runs='<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        elif block.kind == "caption":
            out += _para("", "Caption", block.para_id, runs=_marked(_run("Table ") + _field(
                "SEQ Table \\* ARABIC", "1"), block.mark) + _run(": " + block.text))
        elif block.kind == "xref":
            out += _para("", None, block.para_id, runs=_run(block.text + " See ") + _field(
                f"REF {block.target} \\h", "Table 1") + _run("."))
        elif block.kind == "pageref":
            out += _para("", None, block.para_id, runs=_run(block.text + " (see page ") + _field(
                f"PAGEREF {block.target} \\h", "1") + _run(")."))
        elif block.kind == "pagefield":
            out += _para("", None, block.para_id, runs=_run("This paragraph is on page ") + _field("PAGE", "1")
                         + _run(". " + block.text))
        elif block.kind == "para":
            out += _para(block.text, None, block.para_id)
        elif block.kind == "lead":
            out += _para(block.text, None, block.para_id, extra_ppr="<w:keepNext/>")
        elif block.kind == "bullet":
            out += _para(block.text, "ListBullet", block.para_id)
        elif block.kind == "panel":
            panels += 1
            out += _para("", None, block.para_id, runs=_panel(block.text, panels) + _run(block.text))
        elif block.kind == "break":
            out += _para("", None, block.para_id, runs='<w:r><w:br w:type="page"/></w:r>')
        elif block.kind == "table":
            out += _table(block.rows, block.keep)
        elif block.kind == "noted":
            notes += 1
            out += _para("", None, block.para_id, runs=_run(block.text) + (
                '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                f'<w:footnoteReference w:id="{notes}"/></w:r>'))
        elif block.kind == "raw":
            out += block.text
    return out


def footnotes_xml(blocks: list[Block]) -> str | None:
    """The footnotes of ``blocks``' ``noted`` paragraphs, numbered 1, 2... in order."""
    noted = [block for block in blocks if block.kind == "noted"]
    if not noted:
        return None
    return (f'<w:footnotes xmlns:w="{W_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/>'
            '</w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
            '<w:continuationSeparator/></w:r></w:p></w:footnote>'
            + "".join(f'<w:footnote w:id="{k}">' + _para("", "FootnoteText", runs=(
                '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                + _run(" " + block.note))) + "</w:footnote>" for k, block in enumerate(noted, start=1))
            + "</w:footnotes>")


def _field(instruction: str, result: str) -> str:
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t>{result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


FACE = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'


def _style(kind: str, style_id: str, name: str, ppr: str = "", rpr: str = "", based: str = "Normal") -> str:
    return (f'<w:style w:type="{kind}" w:styleId="{style_id}"><w:name w:val="{name}"/>'
            f'<w:basedOn w:val="{based}"/>' + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "")
            + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "") + "</w:style>")


STYLES = (
    f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>{FACE}<w:sz w:val="22"/><w:szCs w:val="22"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
    '</w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + _style("paragraph", "Heading1", "heading 1",
             '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/>',
             '<w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/>')
    + _style("paragraph", "ListBullet", "List Bullet",
             '<w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/><w:ind w:left="720" w:hanging="360"/>')
    + _style("paragraph", "Cell", "Cell", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
             '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    + _style("paragraph", "Panel", "Panel", '<w:spacing w:after="60" w:line="240" w:lineRule="auto"/>',
             '<w:color w:val="1F3864"/><w:sz w:val="18"/><w:szCs w:val="18"/>')
    + _style("paragraph", "FootnoteText", "footnote text", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
             '<w:sz w:val="18"/><w:szCs w:val="18"/>')
    + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
      '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
    + _style("paragraph", "TOC1", "toc 1",
             '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9062"/></w:tabs><w:spacing w:after="100"/>')
    + _style("paragraph", "Caption", "caption", '<w:keepNext/><w:spacing w:after="200" w:line="240" '
             'w:lineRule="auto"/>', '<w:i/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/>')
    + _style("paragraph", "Header", "header",
             '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs>'
             '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    + _style("paragraph", "Footer", "footer",
             '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs>'
             '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    + "</w:styles>")

NUMBERING = (
    f'<w:numbering xmlns:w="{W_NS}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def package(blocks: list[Block], *, title: str = "Made-up quarterly report", header: str | None = None,
            styles: str = "", even_odd: bool = False, title_page: bool = False, styleref: tuple = ()) -> bytes:
    """``blocks`` as a ``.docx``: A4, 2.5 cm margins, a header (``title``) and a footer
    (``Page N of M``), the footnotes of its ``noted`` paragraphs; ``styles``: more style
    definitions.  ``even_odd``: even pages have a header of their own
    (``w:evenAndOddHeaders``); ``title_page``: the first page has a header and a footer of its
    own, with no page number (``w:titlePg``).  ``styleref``: ``STYLEREF`` instructions the
    header shows after its title, each after a bar (round 21), cached as ``Chapter``."""
    header = header if header is not None else title
    references = '<w:headerReference w:type="default" r:id="rIdH"/>'
    if even_odd:
        references += '<w:headerReference w:type="even" r:id="rIdHe"/>'
    if title_page:
        references += '<w:headerReference w:type="first" r:id="rIdHf"/><w:footerReference w:type="first" r:id="rIdFf"/>'
    sect = (f'<w:sectPr>{references}'
            '<w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" '
            'w:gutter="0"/><w:cols w:space="708"/>' + ("<w:titlePg/>" if title_page else "")
            + '<w:docGrid w:linePitch="360"/></w:sectPr>')
    ns = f'xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:w14="{W14_NS}"'
    parts = {
        "word/document.xml": f'<w:document {ns}><w:body>{body_xml(blocks)}{sect}</w:body></w:document>',
        "word/styles.xml": STYLES.replace("</w:styles>", styles + "</w:styles>"),
        "word/numbering.xml": NUMBERING,
        "word/header1.xml": f'<w:hdr {ns}>' + _para(header, "Header", runs=_run(header) + "".join(
            _run(" | ") + _field(instruction, "Chapter") for instruction in styleref)
            if styleref else "") + "</w:hdr>",
        "word/footer1.xml": (f'<w:ftr {ns}>' + _para("", "Footer", runs=(
            '<w:r><w:tab/></w:r>' + _run("Page ") + _field("PAGE", "1") + _run(" of ")
            + _field("NUMPAGES", "1"))) + "</w:ftr>"),
    }
    rels = [("rIdS", "styles", "styles.xml"), ("rIdN", "numbering", "numbering.xml"),
            ("rIdH", "header", "header1.xml"), ("rIdF", "footer", "footer1.xml")]
    types = [("/word/document.xml", CT + "document.main+xml"), ("/word/styles.xml", CT + "styles+xml"),
             ("/word/numbering.xml", CT + "numbering+xml"), ("/word/header1.xml", CT + "header+xml"),
             ("/word/footer1.xml", CT + "footer+xml")]
    if even_odd:
        parts["word/header2.xml"] = f'<w:hdr {ns}>{_para("Even pages: " + header, "Header")}</w:hdr>'
        parts["word/settings.xml"] = f'<w:settings xmlns:w="{W_NS}"><w:evenAndOddHeaders/></w:settings>'
        rels += [("rIdHe", "header", "header2.xml"), ("rIdSt", "settings", "settings.xml")]
        types += [("/word/header2.xml", CT + "header+xml"), ("/word/settings.xml", CT + "settings+xml")]
    if title_page:
        parts["word/header3.xml"] = f'<w:hdr {ns}>{_para(title.upper(), "Header")}</w:hdr>'
        parts["word/footer2.xml"] = f'<w:ftr {ns}>{_para("A made-up report", "Footer")}</w:ftr>'
        rels += [("rIdHf", "header", "header3.xml"), ("rIdFf", "footer", "footer2.xml")]
        types += [("/word/header3.xml", CT + "header+xml"), ("/word/footer2.xml", CT + "footer+xml")]
    notes = footnotes_xml(blocks)
    if notes is not None:
        parts["word/footnotes.xml"] = notes
        rels.append(("rIdFn", "footnotes", "footnotes.xml"))
        types.append(("/word/footnotes.xml", CT + "footnotes+xml"))
    parts["word/_rels/document.xml.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{rid}" Type="{R_NS}/{kind}" Target="{target}"/>' for rid, kind, target in rels)
        + "</Relationships>")
    parts["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(f'<Override PartName="{name}" ContentType="{kind}"/>' for name, kind in types) + "</Types>")
    parts["_rels/.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R_NS}/officeDocument" Target="word/document.xml"/></Relationships>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data.encode("utf-8"))
    return buffer.getvalue()
