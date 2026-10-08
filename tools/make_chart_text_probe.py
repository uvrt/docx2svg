#!/usr/bin/env python3
"""How Word draws a chart's titles -- the chart's, one with no text, the axes' -- and a
chart too short for its furniture.

``make_chart_probe.py`` measured Word's chart text on charts that state no ``c:txPr`` and
no chart style, and its titles all had text.  A chart Word 365 writes states both -- a
``c:txPr`` whose ``a:defRPr`` is empty, and ``c14:style`` / ``c:style`` -- and docx-agent
saw Word draw such a chart's title larger than that probe's rule, draw a title that has
no text, and draw axis titles, none of which docx2svg drew.  This probe separates the
inputs (ROADMAP.md, F.23).  Every case is a page: a heading line, then one inline chart
and a line after it.  Families (``CASES``):

* ``title`` -- a title whose paragraph states an empty ``a:defRPr`` (or none, or
  ``b="0"``), in a chart that states no ``c:txPr``, an empty one, one of 12 pt, and the
  chart style or not;
* ``auto`` -- a ``c:title`` with no text, over one series, two, a pie, with and without
  an empty ``c:txPr``, the chart style, ``c:txPr`` at 12 pt; and no ``c:title`` at all
  where ``c:autoTitleDeleted`` is ``0``;
* ``axis`` -- axis titles: the value axis', the category axis', both; in the empty form
  (an empty ``a:defRPr``, the value axis' turned by ``a:bodyPr rot="-5400000"``) and in
  Office's chart-style form (10 pt, ``595959``); with no ``a:defRPr``, unturned, at
  12 pt through ``c:txPr``, at 14 pt; with a title, a legend at the bottom, at the top
  and at the left; on horizontal bars and on lines; long;
* ``short`` -- a chart as Word 365 writes one (``c:txPr`` and ``c:spPr`` stated, the chart
  style, no tick marks, one series, a legend at the bottom) at 216 pt across and 40 to
  126 pt high, and at 126 x 63 pt; with values whose axis reaches 5 and 5,000; and
  without its legend.

The charts are written as ``make_chart_probe.py`` writes them (the same series, axes and
document), with no embedded workbook.  The theme is Office's faces: Aptos for the minor
font, Aptos Display for the major; the document's own text is Georgia, which no chart
here names.  Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_chart_text_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.sax.saxutils import escape

import make_anchor_probe as anchor_probe
import make_chart_probe as base
import probe_docx
import wml

SETTINGS = base.SETTINGS
EMU_PT = 12700

#: Word 365's chart style, as it writes it into every chart it makes.
STYLE = ("<mc:AlternateContent xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'>"
         "<mc:Choice Requires='c14' xmlns:c14='http://schemas.microsoft.com/office/drawing/2007/8/2/chart'>"
         "<c14:style val='102'/></mc:Choice><mc:Fallback><c:style val='2'/></mc:Fallback></mc:AlternateContent>")
#: A ``c:txPr`` whose ``a:defRPr`` states nothing (Word 365's).
TX_EMPTY = ("<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang='en-US'/></a:p>"
            "</c:txPr>")
#: Word 365's chart space: a white fill and a line of the text colour at 15%.
SPACE = ("<c:spPr><a:solidFill><a:schemeClr val='bg1'/></a:solidFill><a:ln w='9525' cap='flat' cmpd='sng' "
         "algn='ctr'><a:solidFill><a:schemeClr val='tx1'><a:lumMod val='15000'/><a:lumOff val='85000'/>"
         "</a:schemeClr></a:solidFill><a:round/></a:ln></c:spPr>")
#: Office's chart-style text: the text colour at 65%, the theme's minor face.
OFFICE_FILL = ("<a:solidFill><a:schemeClr val='tx1'><a:lumMod val='65000'/><a:lumOff val='35000'/></a:schemeClr>"
               "</a:solidFill><a:latin typeface='+mn-lt'/><a:ea typeface='+mn-ea'/><a:cs typeface='+mn-cs'/>")
TURNED = " rot='-5400000' vert='horz'"


#: The chart style each way it can be stated: Word 365's (both), ``c:style`` alone, and
#: ``c14:style`` alone.
STYLES = {True: STYLE, "c": "<c:style val='2'/>",
          "c14": STYLE.replace("<mc:Fallback><c:style val='2'/></mc:Fallback>", ""),
          "c14-101": STYLE.replace("val='102'", "val='101'").replace("val='2'", "val='1'")}


def space(body: str, *, tx_pr: str = "", style: bool | str = False, sp_pr: str = "") -> str:
    """``make_chart_probe.chart_space`` with a chart style where ``style`` (:data:`STYLES`)."""
    return ("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
            f"<c:chartSpace xmlns:c='{base.C_NS}' xmlns:a='{base.A_NS}' xmlns:r='{base.R_NS}'>"
            "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>" + STYLES.get(style, "")
            + f"<c:chart>{body}</c:chart>{sp_pr}{tx_pr}</c:chartSpace>")


def rich(text: str, *, def_rpr: str | None = "<a:defRPr/>", body: str = "") -> str:
    """``c:tx`` holding rich text; ``def_rpr`` the paragraph's whole ``a:defRPr`` or
    ``None`` for no ``a:pPr``; ``body`` ``a:bodyPr``'s attributes."""
    ppr = f"<a:pPr>{def_rpr}</a:pPr>" if def_rpr is not None else ""
    return (f"<c:tx><c:rich><a:bodyPr{body}/><a:lstStyle/><a:p>{ppr}<a:r><a:rPr lang='en-US'/>"
            f"<a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>")


def office(text: str, size: int, *, turned: bool = False) -> str:
    """A title in Office's chart-style form (docx-agent's template)."""
    body = (f" rot='{-5400000 if turned else 0}' spcFirstLastPara='1' vertOverflow='ellipsis' vert='horz' "
            "wrap='square' anchor='ctr' anchorCtr='1'")
    def_rpr = (f"<a:defRPr sz='{size}' b='0' i='0' u='none' strike='noStrike' kern='1200' spc='0' baseline='0'>"
               f"{OFFICE_FILL}</a:defRPr>")
    return rich(text, def_rpr=def_rpr, body=body)


def titled(tx: str | None) -> str:
    """``c:title`` holding ``tx`` (none: a title with no text), not overlaid."""
    return f"<c:title>{tx or ''}<c:overlay val='0'/></c:title>"


def axis_xml(kind: str, ax: int, cross: int, position: str, title: str = "", *, ticks: str = "out",
             gridlines: bool = False, between: str = "between") -> str:
    """A ``c:catAx`` or ``c:valAx`` in schema order, ``title`` its whole ``c:title``."""
    head = (f"<c:{kind}><c:axId val='{ax}'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='0'/><c:axPos val='{position}'/>" + ("<c:majorGridlines/>" if gridlines else "")
            + title + "<c:numFmt formatCode='General' sourceLinked='1'/>"
            f"<c:majorTickMark val='{ticks}'/><c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross}'/><c:crosses val='autoZero'/>")
    if kind == "catAx":
        return head + ("<c:auto val='1'/><c:lblAlgn val='ctr'/><c:lblOffset val='100'/><c:noMultiLvlLbl val='0'/>"
                       "</c:catAx>")
    return head + f"<c:crossBetween val='{between}'/></c:valAx>"


def bars(names=base.SERIES, *, direction: str = "col", cats=base.REGIONS, gap: str = "") -> str:
    sers = "".join(base.series(k, name, values, cats=cats) for k, (name, values) in enumerate(names))
    gap = gap or "<c:gapWidth val='150'/>"
    return (f"<c:barChart><c:barDir val='{direction}'/><c:grouping val='clustered'/><c:varyColors val='0'/>{sers}"
            f"{gap}<c:axId val='1'/><c:axId val='2'/></c:barChart>")


def lines(names=base.SERIES) -> str:
    sers = "".join(base.series(k, name, values, invert=False, after_val="<c:smooth val='0'/>")
                   for k, (name, values) in enumerate(names))
    return (f"<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>{sers}"
            "<c:marker val='1'/><c:axId val='1'/><c:axId val='2'/></c:lineChart>")


def chart(*, title: str | None = None, deleted: bool | None = None, group: str | None = None,
          cat_title: str = "", val_title: str = "", direction: str = "col", legend: str | None = "r",
          tx_pr: str = "", style: bool | str = False, sp_pr: str = "", ticks: str = "out", gridlines: bool = True,
          names=base.SERIES) -> str:
    """A chart: ``title`` its whole ``c:title`` (``None``: none), ``deleted`` its
    ``c:autoTitleDeleted`` (``None``: ``1`` without a title and ``0`` with one;
    ``"absent"``: no element), the group (bars of ``names`` by default) and its two axes
    with their titles."""
    if deleted is None:
        deleted = title is None
    group = group or bars(names, direction=direction)
    cat_pos, val_pos = ("l", "b") if direction == "bar" else ("b", "l")
    axes = (axis_xml("catAx", 1, 2, cat_pos, cat_title, ticks=ticks)
            + axis_xml("valAx", 2, 1, val_pos, val_title, ticks=ticks, gridlines=gridlines))
    flag = "" if deleted == "absent" else f"<c:autoTitleDeleted val='{int(deleted)}'/>"
    body = (title or "") + flag + base.plot(group, axes)
    return space(body + base.tail(base.legend(legend)), tx_pr=tx_pr, style=style, sp_pr=sp_pr)


def pie(title: str | None, *, tx_pr: str = "") -> str:
    sers = base.series(0, "Share", (5, 3, 2, 1), invert=False)
    groups = f"<c:pieChart><c:varyColors val='1'/>{sers}<c:firstSliceAng val='0'/></c:pieChart>"
    body = (title or "") + f"<c:autoTitleDeleted val='{int(title is None)}'/>" + base.plot(groups, "")
    return space(body + base.tail(base.legend("r")), tx_pr=tx_pr)


ONE = (("Plan", (4, 2, 3, 5)),)
#: Values whose axis labels are four digits wide.
WIDE = (("Plan", (4000, 2000, 3000, 5000)), ("Actual", (2000, 4000, 1000, 3000)))
EMPTY = "<a:defRPr/>"
VALUE = titled(rich("Revenue", body=TURNED))
CATEGORY = titled(rich("Region"))
VALUE_OFFICE = titled(office("Revenue", 1000, turned=True))
CATEGORY_OFFICE = titled(office("Region", 1000))
TITLE = titled(rich("Sales by region"))
LETTERS = ("A", "B", "C")
GAP = "<c:gapWidth val='219'/><c:overlap val='-27'/>"


def short(height_pt: float, *, values=(1, 3, 2), legend: str | None = "b", width_pt: float = 216,
          size: int | None = None) -> "Case":
    """A chart as Word 365 writes one, ``width_pt`` x ``height_pt``; ``size`` its
    ``c:txPr``'s, in hundredths of a point."""
    group = bars((("Header", values),), cats=LETTERS, gap=GAP)
    xml = chart(group=group, legend=legend, tx_pr=base.text_size(size) if size else TX_EMPTY, style=True,
                sp_pr=SPACE, ticks="none")
    note = (f"{width_pt:g} x {height_pt:g} pt, values {'/'.join(map(str, values))}" + ("" if legend else ", no legend")
            + (f", txPr {size // 100} pt" if size else ""))
    return Case("short", note, xml, (round(width_pt * EMU_PT), round(height_pt * EMU_PT)))


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    chart: str
    extent: tuple[int, int] = (base.CX, base.CY)


def _cases() -> tuple[Case, ...]:
    def title_case(note: str, tx: str, **kw) -> Case:
        return Case("title", note, chart(title=titled(tx), **kw))

    def auto(note: str, **kw) -> Case:
        return Case("auto", note, chart(title=titled(None), **kw))

    def axis(note: str, **kw) -> Case:
        return Case("axis", note, chart(**kw))

    plain = rich("Sales by region")
    sizes = (("txPr empty", TX_EMPTY), ("txPr 8 pt", base.text_size(800)), ("txPr 10 pt", base.text_size(1000)),
             ("txPr 12 pt", base.text_size(1200)), ("txPr 18 pt", base.text_size(1800)))
    return (
        title_case("empty defRPr", plain),
        title_case("empty defRPr, style", plain, style=True),
        title_case("empty defRPr, c:style alone", plain, style="c"),
        title_case("empty defRPr, c14:style alone", plain, style="c14"),
    ) + tuple(
        title_case(f"empty defRPr, {note}" + (", style" if style else ""), plain, tx_pr=tx, style=style)
        for style in (False, True) for note, tx in sizes
    ) + (
        title_case("no defRPr, txPr empty, style", rich("Sales by region", def_rpr=None), tx_pr=TX_EMPTY,
                   style=True),
        title_case("defRPr b=0, txPr empty, style", rich("Sales by region", def_rpr="<a:defRPr b='0'/>"),
                   tx_pr=TX_EMPTY, style=True),
        title_case("empty defRPr, txPr empty, style, legend at the bottom", plain, tx_pr=TX_EMPTY, style=True,
                   legend="b"),
        auto("one series", names=ONE),
        auto("one series, txPr empty, style", names=ONE, tx_pr=TX_EMPTY, style=True),
        auto("one series, txPr 12 pt", names=ONE, tx_pr=base.text_size(1200)),
        auto("one series, txPr 12 pt, style", names=ONE, tx_pr=base.text_size(1200), style=True),
        auto("two series"),
        auto("two series, style", style=True),
        auto("one series, autoTitleDeleted 1", names=ONE, deleted=True),
        Case("auto", "pie", pie(titled(None))),
        Case("auto", "no c:title, autoTitleDeleted 0, one series", chart(deleted=False, names=ONE)),
        Case("auto", "no c:title, no autoTitleDeleted, one series", chart(deleted="absent", names=ONE)),
        Case("auto", "no c:title, autoTitleDeleted 0, two series", chart(deleted=False)),
        axis("value", val_title=VALUE),
        axis("category", cat_title=CATEGORY),
        axis("both, txPr empty", val_title=VALUE, cat_title=CATEGORY, tx_pr=TX_EMPTY),
        axis("both, txPr empty, style", val_title=VALUE, cat_title=CATEGORY, tx_pr=TX_EMPTY, style=True),
        axis("both, Office's form", val_title=VALUE_OFFICE, cat_title=CATEGORY_OFFICE),
        axis("both, Office's form, title, legend at the bottom", val_title=VALUE_OFFICE, cat_title=CATEGORY_OFFICE,
             title=titled(office("Sales by region", 1400)), legend="b"),
        axis("bars, both", val_title=titled(rich("Revenue")), cat_title=titled(rich("Region", body=TURNED)),
             direction="bar"),
        axis("value, no defRPr", val_title=titled(rich("Revenue", def_rpr=None, body=TURNED))),
        axis("value, not turned", val_title=titled(rich("Revenue"))),
        axis("value, rot 0", val_title=titled(rich("Revenue", body=" rot='0' vert='horz'"))),
        axis("value, txPr 12 pt", val_title=VALUE, tx_pr=base.text_size(1200)),
        axis("value, txPr 12 pt, style", val_title=VALUE, tx_pr=base.text_size(1200), style=True),
        axis("both, long", val_title=titled(rich("Revenue in thousands of euros", body=TURNED)),
             cat_title=titled(rich("Sales region, as the plan names it"))),
        Case("axis", "lines, both, legend at the top", chart(group=lines(), val_title=VALUE, cat_title=CATEGORY,
                                                             legend="t")),
        axis("value, legend at the left", val_title=VALUE, legend="l"),
        axis("value, 14 pt Arial bold", val_title=titled(rich(
            "Revenue", def_rpr="<a:defRPr sz='1400' b='1'><a:latin typeface='Arial'/></a:defRPr>", body=TURNED))),
        axis("both, no legend", val_title=VALUE, cat_title=CATEGORY, legend=None),
        axis("value, wide labels", val_title=VALUE, names=WIDE),
        axis("category, 14 pt Arial bold", cat_title=titled(rich(
            "Region", def_rpr="<a:defRPr sz='1400' b='1'><a:latin typeface='Arial'/></a:defRPr>"))),
        axis("category, no defRPr", cat_title=titled(rich("Region", def_rpr=None))),
        axis("category, legend at the bottom", cat_title=CATEGORY, legend="b"),
    ) + tuple(short(h) for h in (40, 50, 63, 72, 81, 90, 108, 126)) + (
        short(63, values=(1, 4321, 2)),
        short(90, values=(1, 4321, 2)),
        short(63, legend=None),
        short(63, width_pt=126),
    ) + tuple(short(h) for h in (45, 55, 58, 60, 66, 69)) + tuple(short(h, legend=None) for h in (30, 36, 45)) + (
        title_case("text, autoTitleDeleted 1", plain, deleted=True),
        title_case("empty defRPr, c14:style 101", plain, style="c14-101"),
        Case("short", "216 x 63 pt, legend at the right", short(63, legend="r").chart, short(63).extent),
    ) + tuple(short(h, size=800) for h in (50, 58, 66)) + tuple(short(h, size=1400) for h in (58, 66, 74, 82))


CASES = _cases()


def body(setting: str) -> str:
    out = ""
    first = base._first_rid(setting)
    for number, case in enumerate(CASES):
        rid = f"rId{first + number}"
        out += anchor_probe._p(f"Case {number} {case.family} {case.note}", pageBreakBefore=True)
        out += wml.paragraph(base.inline(rid, number + 1, *case.extent))
        out += anchor_probe._p(f"Case {number} after the chart.")
    return out


def parts() -> tuple:
    return tuple((f"word/charts/chart{k + 1}.xml", base.CHART_TYPE, base.CHART_REL, case.chart)
                 for k, case in enumerate(CASES))


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    settings = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),) if mode is not None else ()
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": base.FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package(body(setting), styles=styles, extra_parts=settings + (base.THEME,) + parts(),
                              final_section=anchor_probe.section())


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"chart-text-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
