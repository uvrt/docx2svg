#!/usr/bin/env python3
"""How Word draws SmartArt (``dgm:relIds`` in a ``w:drawing``): the laid-out drawing it
caches for a diagram, drawn as Word draws it.

docx2svg draws a diagram from the drawing Word caches for it (``dsp:drawing``), through
``ooxml-common``'s scene renderers (``docx2svg.diagram``).  Word itself lays a diagram out
again from its data whenever it opens a document and writes that layout as the cache on
every save, so a document Word saved holds a cache that is Word's own layout.  What this
probe measures is the drawing of that cache: its shapes' colours and outlines, and above
all their text -- face, size, colour, where each line stands in its shape.

Every case is a page: a heading line, then one diagram -- inline in a paragraph of its
own, or floating -- and a line after it.  The diagrams are written here: the data model,
a colour definition of this file's own (``urn:docx2svg/colors/probe``: ``accent1`` fills,
``lt1`` lines and text) and, for the layout and the quick style, Office's built-in names
with nothing in them, which Word fills in with its own.  The cache this file writes is
:data:`CACHES`: Word's layout of the same data -- each shape's place, size and text
size, read off a copy Word saved (``read_smartart_probe.py --caches``) and written back
here as numbers -- so that the document Word exports and the one docx2svg draws hold the
same layout.  Cases (``CASES``):

* a Basic Block List of three short names, of three shorter words (which Word sets
  larger), and of five with a long name, whose text Word shrinks and wraps;
* one of two items, each with two items under it, which Word sets as bulleted
  paragraphs in the item's block;
* a floating one, in front of the text.

Only the Basic Block List (``layout/default``) is here, because it is the one Word lays
out from an empty definition: a definition it does not hold, or one named and left empty,
it replaces with that one -- every other layout needs its own written out, which this file
does not do.

The theme is Office's faces: Aptos for the minor font (SmartArt's text), Aptos Display
for the major.  Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_smartart_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from xml.sax.saxutils import escape

import make_anchor_probe as anchor_probe
import probe_docx
import wml
from make_anchor_probe import A_NS, R_NS, WP_NS

DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
DRAWING_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
TYPE = "application/vnd.openxmlformats-officedocument.drawingml."
DRAWING_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml"
OFFICE = "urn:microsoft.com/office/officeart/2005/8/"
SETTINGS = {"none": None, "14": 14, "15": 15}
FACE = {"ascii": "Georgia", "hAnsi": "Georgia", "eastAsia": "Georgia", "cs": "Georgia"}
THEME = wml.theme_part({"latin": "Aptos Display"}, {"latin": "Aptos"})
COLORS_ID = "urn:docx2svg/colors/probe"


@dataclass(frozen=True)
class Case:
    note: str
    layout: str
    #: The nodes: a name, or ``(name, (child, ...))``.
    nodes: tuple
    extent: tuple[int, int] = (5486400, 1600200)
    floating: bool = False
    extra: dict = field(default_factory=dict)


CASES = (
    Case("block list of three", "default", ("Alpha", "Beta", "Gamma")),
    Case("block list of five, one long", "default",
         ("Alpha", "Beta", "A much longer name than the others", "Delta", "Epsilon"), extent=(5486400, 2743200)),
    Case("block list of short words", "default", ("Plan", "Build", "Ship")),
    Case("block list with bullets", "default", (("First", ("one", "two")), ("Second", ("three", "four"))),
         extent=(5486400, 2286000)),
    Case("block list, floating", "default", ("Alpha", "Beta", "Gamma"), floating=True),
)


# -- the parts --------------------------------------------------------------------------


def _id(case: int, k: int, kind: int = 0) -> str:
    return "{%08X-0000-4000-8000-%012X}" % (case * 256 + kind, k)


def _text(name: str | None) -> str:
    if name is None:
        return '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>'
    return (f'<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{escape(name)}</a:t></a:r>'
            "</a:p></dgm:t>")


def data_model(number: int, case: Case, drawing_rid: str) -> str:
    """The data: a document point, a point per node, and the connections between them."""
    doc = _id(number, 0)
    points = (f'<dgm:pt modelId="{doc}" type="doc"><dgm:prSet loTypeId="{OFFICE}layout/{case.layout}" '
              f'qsTypeId="{OFFICE}quickstyle/simple1" csTypeId="{COLORS_ID}"/><dgm:spPr/>{_text(None)}</dgm:pt>')
    links = ""
    serial = [0]

    def add(parent: str, name: str, order: int) -> str:
        serial[0] += 1
        k = serial[0]
        node, link, par, sib = _id(number, k), _id(number, k, 1), _id(number, k, 2), _id(number, k, 3)
        nonlocal points, links
        empty = f"<dgm:prSet/><dgm:spPr/>{_text(None)}"
        points += (f'<dgm:pt modelId="{node}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/>{_text(name)}</dgm:pt>'
                   f'<dgm:pt modelId="{par}" type="parTrans" cxnId="{link}">{empty}</dgm:pt>'
                   f'<dgm:pt modelId="{sib}" type="sibTrans" cxnId="{link}">{empty}</dgm:pt>')
        links += (f'<dgm:cxn modelId="{link}" srcId="{parent}" destId="{node}" srcOrd="{order}" destOrd="0" '
                  f'parTransId="{par}" sibTransId="{sib}"/>')
        return node

    for order, node in enumerate(case.nodes):
        name, kids = (node, ()) if isinstance(node, str) else node
        parent = add(doc, name, order)
        for kid_order, kid in enumerate(kids):
            add(parent, kid, kid_order)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel xmlns:dgm="{DGM}" '
            f'xmlns:a="{A_NS}"><dgm:ptLst>{points}</dgm:ptLst><dgm:cxnLst>{links}</dgm:cxnLst><dgm:bg/><dgm:whole/>'
            '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">'
            f'<dsp:dataModelExt xmlns:dsp="{DSP}" relId="{drawing_rid}" '
            'minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>')


def layout_def(name: str) -> str:
    """A built-in layout, named and left empty: Word lays the data out with its own."""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:layoutDef xmlns:dgm="{DGM}" '
            f'xmlns:a="{A_NS}" uniqueId="{OFFICE}layout/{name}"><dgm:title val=""/><dgm:desc val=""/>'
            '<dgm:layoutNode name="diagram"/></dgm:layoutDef>')


STYLE = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:styleDef xmlns:dgm="{DGM}" xmlns:a="{A_NS}" '
         f'uniqueId="{OFFICE}quickstyle/simple1"><dgm:title val=""/><dgm:desc val=""/></dgm:styleDef>')


def _label(name: str, fill: str, line: str, text: str) -> str:
    def paint(tag: str, colour: str) -> str:
        return f'<dgm:{tag} meth="repeat">{colour}</dgm:{tag}>' if colour else f"<dgm:{tag}/>"
    return (f'<dgm:styleLbl name="{name}">' + paint("fillClrLst", fill) + paint("linClrLst", line)
            + "<dgm:effectClrLst/><dgm:txLinClrLst/>" + paint("txFillClrLst", text) + "<dgm:txEffectClrLst/>"
            "</dgm:styleLbl>")


def colors_def() -> str:
    """This file's own colours: shapes in ``accent1`` outlined and lettered in ``lt1``; the
    arrows between them in ``accent1`` at 60% tint; text outside a shape in ``tx1``."""
    accent, light, dark = '<a:schemeClr val="accent1"/>', '<a:schemeClr val="lt1"/>', '<a:schemeClr val="tx1"/>'
    tint = '<a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr>'
    labels = "".join([
        _label("node0", accent, light, light), _label("node1", accent, light, light),
        _label("alignNode1", accent, light, light), _label("lnNode1", accent, light, light),
        _label("sibTrans2D1", tint, tint, light), _label("sibTrans1D1", accent, accent, dark),
        _label("parChTrans1D1", accent, accent, dark), _label("bgShp", tint, "", dark),
        _label("revTx", "", "", dark), _label("txAccent1", "", "", dark),
        _label("bgAccFollowNode1", tint, tint, dark), _label("fgAcc1", light, accent, dark),
    ])
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:colorsDef xmlns:dgm="{DGM}" '
            f'xmlns:a="{A_NS}" uniqueId="{COLORS_ID}"><dgm:title val=""/><dgm:desc val=""/>{labels}</dgm:colorsDef>')


# -- the cache ---------------------------------------------------------------------------

def _recorded(key: str) -> dict[int, list]:
    import json
    from pathlib import Path

    path = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "smartart-observations.json"
    if not path.exists():
        return {}
    return {int(k): v for k, v in json.loads(path.read_text(encoding="utf-8")).get(key, {}).items()}


#: Word's layout of each case's data, as the shapes of its cached drawing: ``[geometry,
#: [x, y, cx, cy], rotation, fill, [line colour, width], [text box], [insets], anchor,
#: paragraphs]`` -- each paragraph ``[text, size in 1/100 pt, align, bullet, marL,
#: indent, line spacing, space after]`` -- and a colour ``accent1``, ``lt1``.  Read off a
#: copy Word saved (``read_smartart_probe.py --caches``) and kept with the recording,
#: ``tests/fixtures/smartart-observations.json``; empty for a case not yet read.
CACHES: dict[int, list] = _recorded("caches")
#: The ``wp:effectExtent`` Word gives each diagram's frame (``l t r b``, EMU), read with
#: the caches: Word sets it from its layout, and an inline diagram stands that far in.
EFFECTS: dict[int, list] = _recorded("effects")


def drawing(number: int, case: Case) -> str:
    """The cached drawing: :data:`CACHES`'s shapes, or none (Word lays it out anyway)."""
    shapes = "".join(cached_shape(number, k, shape) for k, shape in enumerate(CACHES.get(number, ())))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dsp:drawing xmlns:dgm="{DGM}" '
            f'xmlns:dsp="{DSP}" xmlns:a="{A_NS}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/>'
            f"<dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>{shapes}</dsp:spTree></dsp:drawing>")


def _colour(spec: str | None) -> str:
    """``accent1``, ``lt1``, ``accent1@60000`` (a tint) as a fill's colour."""
    if not spec:
        return "<a:noFill/>"
    name, _, tint = spec.partition("@")
    inner = f'<a:tint val="{tint}"/>' if tint else ""
    return f'<a:solidFill><a:schemeClr val="{name}">{inner}</a:schemeClr></a:solidFill>'


def cached_shape(number: int, k: int, shape: tuple) -> str:
    geometry, (x, y, cx, cy), rotation, fill, line, box, insets, anchor, paragraphs = shape
    rot = f' rot="{rotation}"' if rotation else ""
    line_xml = f'<a:ln w="{line[1]}">{_colour(line[0])}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
    body = ""
    if paragraphs:
        left, top, right, bottom = insets
        text = ""
        for words, size, align, bullet, margin, indent, spacing, after in paragraphs:
            bullet_xml = f'<a:buChar char="{bullet}"/>' if bullet else "<a:buNone/>"
            text += (f'<a:p><a:pPr marL="{margin}" lvl="{1 if bullet else 0}" indent="{indent}" algn="{align}">'
                     f'<a:lnSpc><a:spcPct val="{spacing}"/></a:lnSpc>'
                     f'<a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="{after}"/></a:spcAft>'
                     f'{bullet_xml}</a:pPr><a:r><a:rPr lang="en-US" sz="{size}" kern="1200"/>'
                     f"<a:t>{escape(words)}</a:t></a:r></a:p>")
        bx, by, bcx, bcy = box
        body = (f'<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="{left}" tIns="{top}" '
                f'rIns="{right}" bIns="{bottom}" numCol="1" spcCol="1270" anchor="{anchor}" anchorCtr="0">'
                f"<a:noAutofit/></a:bodyPr><a:lstStyle/>{text}</dsp:txBody>"
                f'<dsp:txXfrm><a:off x="{bx}" y="{by}"/><a:ext cx="{bcx}" cy="{bcy}"/></dsp:txXfrm>')
    return (f'<dsp:sp modelId="{_id(number, k, 9)}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/>'
            f'</dsp:nvSpPr><dsp:spPr><a:xfrm{rot}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>{_colour(fill)}{line_xml}</dsp:spPr>'
            '<dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1">'
            '<a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/>'
            '</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>'
            f"{body}</dsp:sp>")


# -- the document ------------------------------------------------------------------------


def graphic(rids: list[str]) -> str:
    dm, lo, qs, cs = rids
    return (f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{DGM}"><dgm:relIds xmlns:dgm="{DGM}" '
            f'r:dm="{dm}" r:lo="{lo}" r:qs="{qs}" r:cs="{cs}"/></a:graphicData></a:graphic>')


def inline(rids: list[str], number: int, cx: int, cy: int) -> str:
    left, top, right, bottom = EFFECTS.get(number - 1, (0, 0, 0, 0))
    return (f'<w:r><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}"><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="{left}" t="{top}" r="{right}" b="{bottom}"/>'
            f'<wp:docPr id="{number}" name="Diagram {number}"/><wp:cNvGraphicFramePr/>{graphic(rids)}'
            "</wp:inline></w:drawing></w:r>")


def _first_rid(setting: str) -> int:
    """Styles is ``rId1``, then the settings part (when there is one) and the theme."""
    return 2 + (1 if SETTINGS[setting] is not None else 0) + 1


def body(setting: str) -> str:
    out = ""
    first = _first_rid(setting)
    for number, case in enumerate(CASES):
        rids = [f"rId{first + 5 * number + k}" for k in range(4)]
        out += anchor_probe._p(f"Case {number} {case.note}", pageBreakBefore=True)
        cx, cy = case.extent
        if case.floating:
            run = anchor_probe.Anchor(("margin", "offset", 457200), ("paragraph", "offset", 190500), cx, cy,
                                      graphic=graphic(rids)).xml(number + 1)
            out += wml.paragraph(wml.run(f"Case{number} anchors a diagram here. ") + run
                                 + wml.run("Text goes on behind it. " * 6))
        else:
            out += wml.paragraph(inline(rids, number + 1, cx, cy))
        out += anchor_probe._p(f"Case {number} after the diagram.")
    return out


def parts(setting: str) -> tuple:
    first = _first_rid(setting)
    out = []
    for number, case in enumerate(CASES):
        drawing_rid = f"rId{first + 5 * number + 4}"
        out += [(f"word/diagrams/data{number + 1}.xml", TYPE + "diagramData+xml", REL + "diagramData",
                 data_model(number, case, drawing_rid)),
                (f"word/diagrams/layout{number + 1}.xml", TYPE + "diagramLayout+xml", REL + "diagramLayout",
                 layout_def(case.layout)),
                (f"word/diagrams/quickStyle{number + 1}.xml", TYPE + "diagramStyle+xml", REL + "diagramQuickStyle",
                 STYLE),
                (f"word/diagrams/colors{number + 1}.xml", TYPE + "diagramColors+xml", REL + "diagramColors",
                 colors_def()),
                (f"word/diagrams/drawing{number + 1}.xml", DRAWING_TYPE, DRAWING_REL, drawing(number, case))]
    return tuple(out)


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    settings = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),) if mode is not None else ()
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package(body(setting), styles=styles, extra_parts=settings + (THEME,) + parts(setting),
                              final_section=anchor_probe.section())


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"smartart-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
