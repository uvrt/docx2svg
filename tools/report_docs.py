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
    (a paragraph anchoring a side text box) or ``break`` (a paragraph holding a page break);
    ``text`` its words (a table's: ``rows`` of cells)."""

    kind: str
    text: str = ""
    rows: tuple = ()
    #: ``w14:paraId`` of a paragraph, as Word stamps it.
    para_id: str = ""


def _words(rng: random.Random, n: int) -> str:
    out = " ".join(rng.choice(WORDS) for _ in range(n))
    return out[0].upper() + out[1:] + "."


def _sentences(rng: random.Random, low: int, high: int) -> str:
    return " ".join(_words(rng, rng.randint(6, 14)) for _ in range(rng.randint(low, high)))


def report(pages: int, seed: int = 0) -> list[Block]:
    """A made-up report of about ``pages`` A4 pages: per part a heading, paragraphs, a
    bulleted list, a side panel, and every other part a table."""
    rng = random.Random(seed)
    blocks: list[Block] = []
    serial = 0

    def pid() -> str:
        nonlocal serial
        serial += 1
        return f"{0x10000000 + serial * 7919:08X}"

    for part in range(pages * 19 // 10):
        blocks.append(Block("heading", f"Part {part + 1}: " + _words(rng, 3)[:-1], para_id=pid()))
        for _ in range(rng.randint(2, 4)):
            blocks.append(Block("para", _sentences(rng, 2, 6), para_id=pid()))
        if part % 3 == 1:
            blocks.append(Block("panel", _sentences(rng, 3, 5), para_id=pid()))
        for _ in range(rng.randint(0, 3)):
            blocks.append(Block("bullet", _words(rng, rng.randint(4, 12)), para_id=pid()))
        if part % 2 == 0:
            rows = tuple(tuple(_words(rng, rng.randint(1, 6 if r else 2)) for _ in range(3))
                         for r in range(rng.randint(3, 7)))
            blocks.append(Block("table", rows=rows))
        blocks.append(Block("para", _sentences(rng, 1, 4), para_id=pid()))
    return blocks


# -- edits -------------------------------------------------------------------------------------


def edit(blocks: list[Block], rng: random.Random) -> tuple[list[Block], str]:
    """One random edit of ``blocks``: change a paragraph's or a cell's text, insert or delete
    a paragraph, a table row or a page break.  ``(new blocks, what was done)``."""
    blocks = list(blocks)
    kinds = ["text", "text", "cell", "insert", "delete", "row-insert", "row-delete", "break", "unbreak", "grow"]
    while True:
        kind = rng.choice(kinds)
        k = rng.randrange(len(blocks))
        block = blocks[k]
        if kind in ("text", "grow") and block.kind in ("heading", "para", "bullet", "panel"):
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
        if kind == "insert":
            blocks.insert(k, Block("para", _sentences(rng, 1, 3), para_id=f"7{rng.randrange(1 << 28):07X}"))
            return blocks, f"insert {k}"
        if kind == "delete" and len(blocks) > 2:
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


def _table(rows: tuple) -> str:
    width = 9000
    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
           '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'
           + "".join(f'<w:gridCol w:w="{width // 3}"/>' for _ in range(3)) + "</w:tblGrid>")
    for r, row in enumerate(rows):
        header = "<w:trPr><w:tblHeader/></w:trPr>" if r == 0 else ""
        cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{width // 3}" w:type="dxa"/>'
                        + ('<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if r == 0 else "")
                        + f'</w:tcPr>{_para(text, "Cell")}</w:tc>' for text in row)
        out += f"<w:tr>{header}{cells}</w:tr>"
    return out + "</w:tbl>"


def body_xml(blocks: list[Block]) -> str:
    out = ""
    panels = 0
    for block in blocks:
        if block.kind == "heading":
            out += _para(block.text, "Heading1", block.para_id)
        elif block.kind == "para":
            out += _para(block.text, None, block.para_id)
        elif block.kind == "bullet":
            out += _para(block.text, "ListBullet", block.para_id)
        elif block.kind == "panel":
            panels += 1
            out += _para("", None, block.para_id, runs=_panel(block.text, panels) + _run(block.text))
        elif block.kind == "break":
            out += _para("", None, block.para_id, runs='<w:r><w:br w:type="page"/></w:r>')
        elif block.kind == "table":
            out += _table(block.rows)
    return out


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


def package(blocks: list[Block], *, title: str = "Made-up quarterly report", header: str | None = None) -> bytes:
    """``blocks`` as a ``.docx``: A4, 2.5 cm margins, a header (``title``) and a footer
    (``Page N of M``)."""
    header = header if header is not None else title
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'
            '<w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" '
            'w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    ns = f'xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:w14="{W14_NS}"'
    parts = {
        "word/document.xml": f'<w:document {ns}><w:body>{body_xml(blocks)}{sect}</w:body></w:document>',
        "word/styles.xml": STYLES,
        "word/numbering.xml": NUMBERING,
        "word/header1.xml": f'<w:hdr {ns}>{_para(header, "Header")}</w:hdr>',
        "word/footer1.xml": (f'<w:ftr {ns}>' + _para("", "Footer", runs=(
            '<w:r><w:tab/></w:r>' + _run("Page ") + _field("PAGE", "1") + _run(" of ")
            + _field("NUMPAGES", "1"))) + "</w:ftr>"),
    }
    rels = [("rIdS", "styles", "styles.xml"), ("rIdN", "numbering", "numbering.xml"),
            ("rIdH", "header", "header1.xml"), ("rIdF", "footer", "footer1.xml")]
    types = [("/word/document.xml", CT + "document.main+xml"), ("/word/styles.xml", CT + "styles+xml"),
             ("/word/numbering.xml", CT + "numbering+xml"), ("/word/header1.xml", CT + "header+xml"),
             ("/word/footer1.xml", CT + "footer+xml")]
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
