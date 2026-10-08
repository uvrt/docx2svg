#!/usr/bin/env python3
"""Tab leaders: which glyphs Word draws in a tab's room, where, and in what.

A tab stop's ``w:leader`` (``dot``, ``hyphen``, ``underscore``, ``heavy``, ``middleDot``)
fills the room a tab to it takes; nothing of it was drawn (``tab-leader-not-drawn``).
The render probe's one dot leader showed its dots on a grid of their advance from the
page's left edge; this probe asks the rest, one family per page, each line a paragraph
of its own (Calibri 11 pt, A4, the left margin off the pixel grid and off every leader's
grid):

* ``phase`` -- a left stop at 6,000 twips with a dot leader, after ``Lead`` and 0 to 39
  ``i``: where the first glyph goes against the text's end, as it moves through a dot;
* ``end`` -- a right stop at 8,000 twips with a dot leader, before ``H`` and 0 to 39
  ``i``: where the last glyph goes against the text's start;
* ``kinds`` -- every leader, before left, right, centre and decimal stops;
* ``faces`` -- every leader in Times New Roman 12, Arial 9, Georgia 14, Cambria 26 and
  Calibri 7 pt, the whole line in the face;
* ``runs`` -- the tab in a run of its own, formatted otherwise than the text around it
  (20 pt, bold, italic, red, Times New Roman, 8 pt, underlined, raised as a
  superscript), and the text around it formatted otherwise than the tab; an underlined
  tab with no leader; a leader tab ending its line, plain and underlined;
* ``indent`` -- the paragraph indented 357 and 1,000 twips, and hanging;
* ``short`` -- a stop so near the text's end that no glyph or one glyph fits, and two
  leader tabs on a line;
* ``align`` -- right-aligned, centred and justified paragraphs with a leader tab.

One document per setting: no ``settings.xml``, mode 14 and mode 15.

Reader: ``read_tab_leader_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
LEADERS = ("dot", "hyphen", "underscore", "heavy", "middleDot")
FAMILIES = ("phase", "end", "kinds", "faces", "runs", "indent", "short", "align")
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1440, "right": 1300, "bottom": 1440, "left": 1442, "header": 700, "footer": 650, "gutter": 0}


def fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


FACES = (("Times New Roman", 24), ("Arial", 18), ("Georgia", 28), ("Cambria", 52), ("Calibri", 14))
#: The tab run's own formatting, family ``runs``.
RUN_FORMATS = {
    "large": {"sz": 40, "szCs": 40},
    "bold": {"b": True},
    "italic": {"i": True},
    "red": {"color": "C00000"},
    "times": {"rFonts": fonts("Times New Roman")},
    "small": {"sz": 16, "szCs": 16},
    "underlined": {"u": "single"},
    "raised": {"vertAlign": "superscript"},
}


def tabs(*stops: tuple[str, int, str | None]) -> str:
    out = ""
    for alignment, position, leader in stops:
        leader_xml = f' w:leader="{leader}"' if leader else ""
        out += f'<w:tab w:val="{alignment}"{leader_xml} w:pos="{position}"/>'
    return out


def _p(runs: str, stops: str, mark: dict | None = None, **props) -> str:
    props.setdefault("spacing", SPACING)
    return wml.paragraph(runs, mark=mark or {}, tabs=stops, **props)


def run(text: str, **props) -> str:
    return wml.run(text, **props)


def lines(family: str) -> list[tuple[str, str]]:
    """``(case, w:p)`` of a family."""
    out: list[tuple[str, str]] = []
    if family == "phase":
        for k in range(40):
            out.append((f"phase/{k}", _p(run("Lead" + "i" * k) + run("\t") + run("After"),
                                         tabs(("left", 6000, "dot")))))
    elif family == "end":
        for k in range(40):
            out.append((f"end/{k}", _p(run("Lead") + run("\t") + run("H" + "i" * k),
                                       tabs(("right", 8000, "dot")))))
    elif family == "kinds":
        for leader in LEADERS:
            for alignment, position, text in (("left", 6000, "After"), ("right", 8000, "Right 12"),
                                              ("center", 6000, "Centred text"), ("decimal", 7000, "1234.5")):
                out.append((f"kinds/{leader}/{alignment}",
                            _p(run("Lead text") + run("\t") + run(text), tabs((alignment, position, leader)))))
    elif family == "faces":
        for face, size in FACES:
            for leader in LEADERS:
                props = {"rFonts": fonts(face), "sz": size, "szCs": size}
                out.append((f"faces/{face}/{size}/{leader}",
                            _p(run("Lead", **props) + run("\t", **props) + run("After 7", **props),
                               tabs(("right", 8000, leader)), mark=props)))
    elif family == "runs":
        for name, props in RUN_FORMATS.items():
            for leader in ("dot", "underscore"):
                out.append((f"runs/{name}/{leader}", _p(run("Lead") + run("\t", **props) + run("After 7"),
                                                        tabs(("right", 8000, leader)))))
        # An underlined tab with no leader, and a leader tab that ends its line (plain and
        # underlined).
        out.append(("runs/underlined/none", _p(run("Lead") + run("\t", u="single") + run("After 7"),
                                               tabs(("right", 8000, None)))))
        out.append(("runs/trailing/dot", _p(run("Lead") + run("\t"), tabs(("right", 8000, "dot")))))
        out.append(("runs/trailing-underlined/dot", _p(run("Lead") + run("\t", u="single"),
                                                       tabs(("left", 7000, "dot")))))
        around = {"sz": 40, "szCs": 40, "rFonts": fonts("Times New Roman")}
        out.append(("runs/around/dot", _p(run("Lead", **around) + run("\t") + run("After 7", **around),
                                          tabs(("right", 8000, "dot")))))
        one = (f'<w:r>{wml.rpr(sz=28, szCs=28)}<w:t xml:space="preserve">Lead</w:t><w:tab/>'
               '<w:t xml:space="preserve">After 7</w:t></w:r>')
        out.append(("runs/in-text/dot", _p(one, tabs(("right", 8000, "dot")))))
    elif family == "indent":
        for name, ind in (("357", {"left": 357}), ("1000", {"left": 1000}), ("first", {"firstLine": 500}),
                          ("hanging", {"left": 1440, "hanging": 720})):
            for leader in ("dot", "middleDot"):
                out.append((f"indent/{name}/{leader}", _p(run("Lead") + run("\t") + run("After 7"),
                                                          tabs(("left", 6000, leader)), ind=ind)))
    elif family == "short":
        for k, position in enumerate((700, 720, 740, 760, 780, 800, 820, 850)):
            out.append((f"short/{k}", _p(run("Lead") + run("\t") + run("After"), tabs(("left", position, "dot")))))
        out.append(("short/two", _p(run("One") + run("\t") + run("Two") + run("\t") + run("Three"),
                                    tabs(("left", 3000, "dot"), ("right", 8000, "hyphen")))))
    elif family == "align":
        for jc in ("right", "center", "both"):
            out.append((f"align/{jc}", _p(run("Lead") + run("\t") + run("After 7 and more words"),
                                          tabs(("left", 5000, "dot")), jc=jc)))
    return out


def body() -> list[tuple[str, str]]:
    """``(case, w:p)`` of every paragraph, each family led by an anchor that starts its page."""
    out: list[tuple[str, str]] = []
    for family in FAMILIES:
        out.append((f"{family}/anchor", wml.paragraph(run(f"Family {family}"), mark={}, pageBreakBefore=True,
                                                      spacing=SPACING)))
        out.extend(lines(family))
    return out


def section() -> str:
    m = MARGINS
    return (f"<w:sectPr><w:pgSz w:w=\"{PAGE['w']}\" w:h=\"{PAGE['h']}\"/>"
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def build(name: str) -> bytes:
    """``name``: ``leader-<setting>``."""
    setting = name.split("-", 1)[1]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package("".join(xml for _, xml in body()), final_section=section(), styles=styles,
                              extra_parts=extra)


DOCUMENTS = tuple(f"leader-{setting}" for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)
