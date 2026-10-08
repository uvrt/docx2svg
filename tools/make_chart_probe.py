#!/usr/bin/env python3
"""How Word draws a chart (``c:chart`` in a ``w:drawing``): its default text, its frame,
its plot area and legend, its colours and lines, inline and floating.

docx2svg draws a chart with ``ooxml-common``'s chart layout, whose every constant was
measured on PowerPoint (pptx2svg ROADMAP.md, Phase 3).  Word draws charts with the same
Office chart engine, but it is a different application with its own defaults, so this
probe measures Word on the same markup: where it differs, the difference is a field of
``ooxml_common.chart.rules.ChartRules`` (its ``WORD`` variant), not a guess.

Every case is a page: a heading line, then one chart -- inline in a paragraph of its own,
or floating (``wp:anchor``) beside a paragraph of text -- and a line after it.  The charts
are written here, part by part (``word/charts/chartN.xml``, related from the document),
with their numbers in ``c:numCache`` / ``c:strCache`` and no embedded workbook: Word draws
a chart from its cache.  The values are whole numbers, so every axis label is one (Word
writes a fraction with the host's decimal separator, which is the host's and not Word's).
Families (``CASES``):

* ``text`` -- a chart that states no text properties at all; ``c:txPr`` at 8 and 12 pt;
  titles whose ``a:defRPr`` states a size, bold, a face or ``+mj-lt``;
* ``legend`` -- the legend at the right, bottom, top (under a title and alone), left and
  top right, and none; a side legend at 6, 12, 14 and 18 pt and with long names;
* ``type`` -- horizontal bars (clustered, stacked, legend below), stacked and 100%
  columns, three series, lines with and without markers, an area, a pie, a doughnut, a
  scatter, a radar;
* ``paint`` -- the chart's and the plot area's fill and line stated, each alone, and
  stated as none; series' own fills and lines (a line with no width, lines of 2.25 pt),
  data labels;
* ``place`` -- a floating chart in front of the text, one text wraps around, and an inline
  chart whose extent is not in whole twips, on a centred line.

The theme is Office's faces: Aptos for the minor (body) font, Aptos Display for the
major; the document's own text is Georgia, which no chart here names.  Three documents: no ``settings.xml``, mode 14
and mode 15.  Reader: ``read_chart_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from xml.sax.saxutils import escape

import make_anchor_probe as anchor_probe
import probe_docx
import wml
from make_anchor_probe import A_NS, R_NS, WP_NS

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
SETTINGS = {"none": None, "14": 14, "15": 15}
#: The probe's own text: a face no chart here is drawn in, so the reader can tell them apart.
FACE = {"ascii": "Georgia", "hAnsi": "Georgia", "eastAsia": "Georgia", "cs": "Georgia"}
THEME = wml.theme_part({"latin": "Aptos Display"}, {"latin": "Aptos"})

REGIONS = ("North", "South", "East", "West")
SERIES = (("Plan", (4, 2, 3, 5)), ("Actual", (2, 4, 1, 3)))
LONG = (("Planned revenue", (4, 2, 3, 5)), ("Actual revenue", (2, 4, 1, 3)))
THREE = SERIES + (("Forecast", (3, 3, 4, 4)),)
#: The chart's extent, EMU: 6 x 3.5 inches.
CX, CY = 5486400, 3200400


# -- chart parts ------------------------------------------------------------------------


def _points(values) -> str:
    return "".join(f"<c:pt idx='{i}'><c:v>{escape(str(v))}</c:v></c:pt>" for i, v in enumerate(values))


def categories(names, tag: str = "cat") -> str:
    return (f"<c:{tag}><c:strRef><c:f>Sheet1!$A$2:$A${len(names) + 1}</c:f><c:strCache>"
            f"<c:ptCount val='{len(names)}'/>{_points(names)}</c:strCache></c:strRef></c:{tag}>")


def numbers(values, tag: str = "val") -> str:
    return (f"<c:{tag}><c:numRef><c:f>Sheet1!$B$2:$B${len(values) + 1}</c:f><c:numCache>"
            f"<c:formatCode>General</c:formatCode><c:ptCount val='{len(values)}'/>{_points(values)}"
            f"</c:numCache></c:numRef></c:{tag}>")


def series_name(name: str) -> str:
    return (f"<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{escape(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>")


def series(index: int, name: str, values, *, cats=REGIONS, sp_pr: str = "", before_cat: str = "",
           after_val: str = "", invert: bool = True) -> str:
    """One ``c:ser``: idx, order, tx, spPr, (invertIfNegative / marker / dLbls), cat, val,
    (smooth)."""
    return (f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>{series_name(name)}{sp_pr}"
            + ("<c:invertIfNegative val='0'/>" if invert else "") + before_cat
            + categories(cats) + numbers(values) + after_val + "</c:ser>")


def cat_axis(ax: int, cross: int, *, position: str = "b", delete: bool = False) -> str:
    return (f"<c:catAx><c:axId val='{ax}'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='{int(delete)}'/><c:axPos val='{position}'/>"
            "<c:numFmt formatCode='General' sourceLinked='1'/><c:majorTickMark val='out'/>"
            "<c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross}'/><c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
            "<c:lblOffset val='100'/><c:noMultiLvlLbl val='0'/></c:catAx>")


def val_axis(ax: int, cross: int, *, position: str = "l", gridlines: bool = True, between: str = "between") -> str:
    return (f"<c:valAx><c:axId val='{ax}'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='0'/><c:axPos val='{position}'/>" + ("<c:majorGridlines/>" if gridlines else "")
            + "<c:numFmt formatCode='General' sourceLinked='1'/><c:majorTickMark val='out'/>"
            "<c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross}'/><c:crosses val='autoZero'/><c:crossBetween val='{between}'/></c:valAx>")


def title(text: str, *, def_rpr: str = "", run_rpr: str = "") -> str:
    """``c:title`` with rich text; ``def_rpr`` is the paragraph's whole ``a:defRPr``, or
    nothing at all, and ``run_rpr`` the run's ``a:rPr``."""
    ppr = f"<a:pPr>{def_rpr}</a:pPr>" if def_rpr else ""
    return (f"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{ppr}<a:r>{run_rpr}<a:t>{escape(text)}</a:t></a:r>"
            "</a:p></c:rich></c:tx><c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>")


NO_TITLE = "<c:autoTitleDeleted val='1'/>"


def legend(position: str | None) -> str:
    return f"<c:legend><c:legendPos val='{position}'/><c:overlay val='0'/></c:legend>" if position else ""


def solid(colour: str) -> str:
    return f"<a:solidFill><a:srgbClr val='{colour}'/></a:solidFill>"


def chart_space(body: str, *, tx_pr: str = "", sp_pr: str = "") -> str:
    """A ``c:chartSpace``: ``c:chart`` holding ``body``, then the space's ``c:spPr`` and
    ``c:txPr``, schema order."""
    return ("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
            f"<c:chartSpace xmlns:c='{C_NS}' xmlns:a='{A_NS}' xmlns:r='{R_NS}'>"
            "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
            f"<c:chart>{body}</c:chart>{sp_pr}{tx_pr}</c:chartSpace>")


def text_size(hundredths: int) -> str:
    return (f"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='{hundredths}'/></a:pPr>"
            "<a:endParaRPr lang='en-US'/></a:p></c:txPr>")


def plot(groups: str, axes: str, *, sp_pr: str = "") -> str:
    return f"<c:plotArea><c:layout/>{groups}{axes}{sp_pr}</c:plotArea>"


def tail(legend_xml: str) -> str:
    return f"{legend_xml}<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/>"


def columns(*, direction: str = "col", grouping: str = "clustered", overlap: str = "", ser_sp=("", "", ""),
            labels: str = "", names=SERIES) -> str:
    sers = "".join(series(k, name, values, sp_pr=ser_sp[k], before_cat=labels) for k, (name, values) in
                   enumerate(names))
    return (f"<c:barChart><c:barDir val='{direction}'/><c:grouping val='{grouping}'/><c:varyColors val='0'/>{sers}"
            f"<c:gapWidth val='150'/>{overlap}<c:axId val='1'/><c:axId val='2'/></c:barChart>")


def bar_axes(direction: str = "col") -> str:
    if direction == "bar":
        return cat_axis(1, 2, position="l") + val_axis(2, 1, position="b")
    return cat_axis(1, 2) + val_axis(2, 1)


def column_chart(*, heading: str | None = "Sales by region", def_rpr: str = "", run_rpr: str = "",
                 where: str | None = "r", tx_pr: str = "", sp_pr: str = "", plot_sp: str = "", **kw) -> str:
    direction = kw.get("direction", "col")
    body = (title(heading, def_rpr=def_rpr, run_rpr=run_rpr) if heading else NO_TITLE)
    body += plot(columns(**kw), bar_axes(direction), sp_pr=plot_sp) + tail(legend(where))
    return chart_space(body, tx_pr=tx_pr, sp_pr=sp_pr)


def line_chart(markers: bool, sp_pr: str = "") -> str:
    marker = "" if markers else "<c:marker><c:symbol val='none'/></c:marker>"
    sers = "".join(series(k, name, values, invert=False, sp_pr=sp_pr, before_cat=marker,
                          after_val="<c:smooth val='0'/>") for k, (name, values) in enumerate(SERIES))
    groups = (f"<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>{sers}"
              f"<c:marker val='1'/><c:axId val='1'/><c:axId val='2'/></c:lineChart>")
    return chart_space(NO_TITLE + plot(groups, bar_axes()) + tail(legend("b")))


def area_chart() -> str:
    sers = "".join(series(k, name, values, invert=False) for k, (name, values) in enumerate(SERIES))
    groups = (f"<c:areaChart><c:grouping val='standard'/><c:varyColors val='0'/>{sers}"
              "<c:axId val='1'/><c:axId val='2'/></c:areaChart>")
    return chart_space(title("Area") + plot(groups, cat_axis(1, 2) + val_axis(2, 1, between="midCat"))
                       + tail(legend("r")))


def pie_chart(kind: str) -> str:
    sers = series(0, "Share", (5, 3, 2, 1), invert=False)
    hole = "<c:holeSize val='50'/>" if kind == "doughnutChart" else ""
    groups = f"<c:{kind}><c:varyColors val='1'/>{sers}<c:firstSliceAng val='0'/>{hole}</c:{kind}>"
    return chart_space(title("Share") + plot(groups, "") + tail(legend("r")))


def scatter_chart() -> str:
    sers = "".join(
        f"<c:ser><c:idx val='{k}'/><c:order val='{k}'/>{series_name(name)}"
        + numbers((1, 2, 3, 4), "xVal") + numbers(values, "yVal") + "<c:smooth val='0'/></c:ser>"
        for k, (name, values) in enumerate(SERIES))
    groups = (f"<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>{sers}"
              "<c:axId val='1'/><c:axId val='2'/></c:scatterChart>")
    axes = val_axis(1, 2, position="b", gridlines=False, between="midCat") + val_axis(2, 1, between="midCat")
    return chart_space(NO_TITLE + plot(groups, axes) + tail(legend("r")))


def radar_chart() -> str:
    sers = "".join(series(k, name, values, invert=False) for k, (name, values) in enumerate(SERIES))
    groups = (f"<c:radarChart><c:radarStyle val='marker'/><c:varyColors val='0'/>{sers}"
              "<c:axId val='1'/><c:axId val='2'/></c:radarChart>")
    return chart_space(NO_TITLE + plot(groups, bar_axes()) + tail(legend("b")))


# -- the cases -------------------------------------------------------------------------


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    chart: str
    #: ``inline``, ``float`` (``wrapNone`` in front, against the margin) or ``square``
    #: (``wrapSquare``, text beside it).
    where: str = "inline"
    extent: tuple[int, int] = (CX, CY)
    jc: str | None = None
    extra: dict = field(default_factory=dict)


LINE = "<a:ln w='19050'>" + solid("7030A0") + "</a:ln>"
SPACE_PAINT = "<c:spPr>" + solid("F2F2F2") + "<a:ln w='12700'>" + solid("4472C4") + "</a:ln></c:spPr>"
SPACE_NONE = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>"
PLOT_PAINT = "<c:spPr>" + solid("FFF2CC") + "<a:ln w='9525'>" + solid("C00000") + "</a:ln></c:spPr>"
LABELS = ("<c:dLbls><c:showLegendKey val='0'/><c:showVal val='1'/><c:showCatName val='0'/>"
          "<c:showSerName val='0'/><c:showPercent val='0'/><c:showBubbleSize val='0'/></c:dLbls>")


def _cases() -> tuple[Case, ...]:
    return (
        Case("text", "nothing stated", column_chart()),
        Case("text", "txPr 12 pt", column_chart(tx_pr=text_size(1200))),
        Case("text", "txPr 8 pt", column_chart(tx_pr=text_size(800))),
        Case("text", "title 14 pt bold +mj-lt", column_chart(
            def_rpr="<a:defRPr sz='1400' b='1'><a:latin typeface='+mj-lt'/></a:defRPr>")),
        Case("legend", "bottom", column_chart(where="b")),
        Case("legend", "top", column_chart(where="t")),
        Case("legend", "left", column_chart(where="l")),
        Case("legend", "top right", column_chart(where="tr")),
        Case("legend", "none, no title", column_chart(where=None, heading=None)),
        Case("type", "bars", column_chart(direction="bar")),
        Case("type", "stacked columns", column_chart(grouping="stacked", overlap="<c:overlap val='100'/>")),
        Case("type", "100% columns", column_chart(grouping="percentStacked", overlap="<c:overlap val='100'/>")),
        Case("type", "lines with markers", line_chart(True)),
        Case("type", "lines", line_chart(False)),
        Case("type", "area", area_chart()),
        Case("type", "pie", pie_chart("pieChart")),
        Case("type", "doughnut", pie_chart("doughnutChart")),
        Case("type", "scatter", scatter_chart()),
        Case("type", "radar", radar_chart()),
        Case("paint", "space and plot area painted", column_chart(sp_pr=SPACE_PAINT, plot_sp=PLOT_PAINT)),
        Case("paint", "space and plot area none", column_chart(sp_pr=SPACE_NONE,
                                                                plot_sp="<c:spPr><a:noFill/></c:spPr>")),
        Case("paint", "series fills, a line with no width",
             column_chart(ser_sp=("<c:spPr>" + solid("70AD47") + "<a:ln>" + solid("000000") + "</a:ln></c:spPr>",
                                  "<c:spPr>" + solid("FFC000") + LINE + "</c:spPr>"))),
        Case("paint", "data labels", column_chart(labels=LABELS)),
        Case("place", "floating, in front", column_chart(), where="float"),
        Case("place", "floating, text wraps around", column_chart(), where="square", extent=(3657600, 2286000)),
        Case("place", "inline, off the twip grid, centred", column_chart(), extent=(CX + 1, CY + 317), jc="center"),
        Case("paint", "space fill alone", column_chart(sp_pr="<c:spPr>" + solid("F2F2F2") + "</c:spPr>")),
        Case("paint", "space line alone", column_chart(
            sp_pr="<c:spPr><a:ln w='12700'>" + solid("4472C4") + "</a:ln></c:spPr>")),
        Case("paint", "plot area line alone", column_chart(
            plot_sp="<c:spPr><a:ln w='12700'>" + solid("C00000") + "</a:ln></c:spPr>")),
        Case("legend", "right, txPr 6 pt", column_chart(tx_pr=text_size(600))),
        Case("legend", "right, txPr 14 pt", column_chart(tx_pr=text_size(1400))),
        Case("legend", "right, txPr 18 pt", column_chart(tx_pr=text_size(1800))),
        Case("legend", "right, long names", column_chart(names=LONG)),
        Case("legend", "left, txPr 12 pt", column_chart(where="l", tx_pr=text_size(1200))),
        Case("legend", "top, no title", column_chart(where="t", heading=None)),
        Case("legend", "bottom, txPr 12 pt", column_chart(where="b", tx_pr=text_size(1200))),
        Case("type", "stacked bars", column_chart(direction="bar", grouping="stacked",
                                                  overlap="<c:overlap val='100'/>")),
        Case("type", "bars, legend at the bottom", column_chart(direction="bar", where="b")),
        Case("type", "three series", column_chart(names=THREE)),
        Case("paint", "lines 2.25 pt", line_chart(False, "<c:spPr><a:ln w='28575'/></c:spPr>")),
        Case("text", "title Arial 14", column_chart(
            def_rpr="<a:defRPr sz='1400'><a:latin typeface='Arial'/></a:defRPr>")),
        Case("text", "title Aptos 18", column_chart(
            def_rpr="<a:defRPr sz='1800'><a:latin typeface='Aptos'/></a:defRPr>")),
        Case("text", "title Arial 24", column_chart(def_rpr="<a:defRPr sz='2400'/>")),
        Case("text", "title defRPr empty", column_chart(def_rpr="<a:defRPr/>")),
        Case("text", "title defRPr b=0", column_chart(def_rpr="<a:defRPr b='0'/>")),
        Case("text", "title run 14 pt, no defRPr", column_chart(run_rpr="<a:rPr lang='en-US' sz='1400'/>")),
        Case("text", "title, txPr Calibri", column_chart(
            tx_pr="<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface='Calibri'/></a:defRPr>"
                  "</a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>")),
        Case("legend", "left, txPr 8 pt", column_chart(where="l", tx_pr=text_size(800))),
        Case("legend", "left, txPr 18 pt", column_chart(where="l", tx_pr=text_size(1800))),
        Case("type", "stacked columns, legend at the bottom",
             column_chart(grouping="stacked", overlap="<c:overlap val='100'/>", where="b")),
    ) + tuple(
        Case("text", f"title {face} {size}", column_chart(
            def_rpr=f"<a:defRPr sz='{size * 100}' b='0'><a:latin typeface='{face}'/></a:defRPr>"))
        for face, size in (("Aptos", 8), ("Aptos", 12), ("Aptos", 16), ("Aptos", 20), ("Aptos", 28),
                           ("Aptos", 36), ("Arial", 10), ("Arial", 28))
    )


CASES = _cases()


# -- the document ----------------------------------------------------------------------

WORDS = ("Text goes on beside the chart, in words of several lengths, so that each line ends at a different "
         "place and the wrap shows where the chart's box is. ") * 4


def graphic(rid: str) -> str:
    return (f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{C_NS}"><c:chart xmlns:c="{C_NS}" r:id="{rid}"/>'
            "</a:graphicData></a:graphic>")


def inline(rid: str, number: int, cx: int, cy: int) -> str:
    return (f'<w:r><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}"><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{number}" name="Chart {number}"/><wp:cNvGraphicFramePr/>{graphic(rid)}'
            "</wp:inline></w:drawing></w:r>")


def floating(rid: str, number: int, cx: int, cy: int, square: bool) -> str:
    wrap = '<wp:wrapSquare wrapText="bothSides"/>' if square else "<wp:wrapNone/>"
    return anchor_probe.Anchor(("margin", "offset", 914400 if square else 457200),
                               ("paragraph", "offset", 0 if square else 190500), cx, cy, wrap=wrap,
                               graphic=graphic(rid)).xml(number)


def _first_rid(setting: str) -> int:
    """The first chart's relationship id: styles is ``rId1``, then the settings part (when
    there is one) and the theme (:func:`probe_docx.package` numbers them in order)."""
    return 2 + (1 if SETTINGS[setting] is not None else 0) + 1


def body(setting: str) -> str:
    out = ""
    first = _first_rid(setting)
    for number, case in enumerate(CASES):
        rid = f"rId{first + number}"
        out += anchor_probe._p(f"Case {number} {case.family} {case.note}", pageBreakBefore=True)
        cx, cy = case.extent
        if case.where == "inline":
            out += wml.paragraph(inline(rid, number + 1, cx, cy), jc=case.jc) if case.jc else wml.paragraph(
                inline(rid, number + 1, cx, cy))
        else:
            out += wml.paragraph(wml.run(f"Case{number} anchors a chart here. ")
                                 + floating(rid, number + 1, cx, cy, case.where == "square") + wml.run(WORDS))
        out += anchor_probe._p(f"Case {number} after the chart.")
    return out


def parts() -> tuple:
    return tuple((f"word/charts/chart{k + 1}.xml", CHART_TYPE, CHART_REL, case.chart) for k, case in enumerate(CASES))


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    settings = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),) if mode is not None else ()
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package(body(setting), styles=styles, extra_parts=settings + (THEME,) + parts(),
                              final_section=anchor_probe.section())


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"chart-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
