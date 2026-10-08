#!/usr/bin/env python3
"""How Word counts list items: instances, overrides, restarts, legal numbering, stories.

docx-agent found docx2svg drawing ``1.`` where Word draws ``7.``: ``w:startOverride`` was
ignored, and items were counted per ``w:num`` where docx-agent's model counts per
``w:abstractNum``.  This probe measures the counting itself.  Every case is a page:
``Case N``, then its blocks; a list item's text names its instance and level (``B1``:
instance ``B``, level 1), so a line of Word's PDF reads ``7.B0``.  Calibri 11 pt, no
paragraph spacing.  Every case has abstract definitions of its own (three levels:
``%1.`` decimal, ``%2.`` lowerLetter, ``%3.`` lowerRoman unless the case says otherwise),
so no count runs from one case into the next.  Families:

* ``override`` -- ``w:lvlOverride``: ``w:startOverride`` (alone; on an instance after
  another instance; on two instances interleaved; at level 1, around level 0 items and
  with another instance's level-1 item; an instance used again after another; its first
  item deeper than the override), a whole ``w:lvl`` in the override (its format and its
  ``w:start``, at level 0 and level 1), both together;
* ``instance`` -- two ``w:num`` over one ``w:abstractNum`` with no override (in turn and
  interleaved; a level-0 item of one between level-1 items of the other), two abstract
  definitions alike (and with one ``w:nsid``, alike and not), ``w:numStyleLink`` /
  ``w:styleLink``, ``numId`` 0 between items;
* ``restart`` -- deeper levels reset by a shallower item, ``w:lvlRestart`` 0 and 1, a
  level skipped, a first item at level 1 (with and without ``%1`` in its label), a
  level's ``w:start`` other than 1;
* ``legal`` -- ``%1.%2.`` labels: upperRoman and lowerLetter parents, ``w:isLgl`` on the
  child level and on level 0;
* ``story`` -- last, the body's list continued through a table's cells, one text box and
  two, one footnote and two, and (in a section of their own) a header holding an item
  of the body's list.

One document per compatibility setting (none, 14, 15).  Reader:
``read_numbering_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_footnote_draw_probe as notes_probe
import probe_docx
import story_docx
import wml
from make_anchor_probe import Anchor
from make_drawing_probe import NO_LINE, graphic, preset, xfrm

SETTINGS = {"none": None, "14": 14, "15": 15}
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
#: The numbering style ``w:numStyleLink`` names.
LINKED_STYLE = "LinkedList"


@dataclass(frozen=True)
class Level:
    fmt: str
    text: str
    start: int = 1
    restart: int | None = None  # w:lvlRestart
    legal: bool = False  # w:isLgl

    def xml(self, ilvl: int) -> str:
        out = f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="{self.start}"/><w:numFmt w:val="{self.fmt}"/>'
        if self.restart is not None:
            out += f'<w:lvlRestart w:val="{self.restart}"/>'
        if self.legal:
            out += "<w:isLgl/>"
        left = 720 + 720 * ilvl
        return out + (f'<w:lvlText w:val="{self.text}"/><w:lvlJc w:val="left"/>'
                      f'<w:pPr><w:ind w:left="{left}" w:hanging="720"/></w:pPr></w:lvl>')


DEFAULT_LEVELS = (Level("decimal", "%1."), Level("lowerLetter", "%2."), Level("lowerRoman", "%3."))
LEGAL_LEVELS = (Level("decimal", "%1."), Level("decimal", "%1.%2."), Level("decimal", "%1.%2.%3."))


@dataclass(frozen=True)
class Abstract:
    levels: tuple[Level, ...] = DEFAULT_LEVELS
    #: ``w:nsid``: ``None`` gives the definition one of its own.
    nsid: str | None = None
    #: ``w:styleLink`` (this definition is the numbering style's) or ``w:numStyleLink``
    #: (this definition is the numbering style's definition's), naming :data:`LINKED_STYLE`.
    link: str | None = None


@dataclass(frozen=True)
class Num:
    abstract: str
    #: ilvl -> (startOverride or None, a replacing Level or None).
    overrides: dict = field(default_factory=dict, hash=False)


@dataclass(frozen=True)
class Case:
    family: str
    name: str
    abstracts: dict = field(hash=False)
    nums: dict = field(hash=False)
    #: ``("i", num, level)`` a list item; ``("p", text)`` a paragraph; ``("x", text)`` a
    #: paragraph with ``numId`` 0; ``("t", [[blocks], ...])`` a one-row table;
    #: ``("box", [blocks])`` a text box anchored in a paragraph; ``("note", [blocks])`` a
    #: paragraph whose footnote holds ``blocks`` after a first line.
    blocks: tuple = ()
    #: The header's blocks: the case starts a section of its own with this header.
    header: tuple | None = None

    @property
    def key(self) -> str:
        return f"{self.family} {self.name}"


def _cases() -> tuple[Case, ...]:
    X = {"X": Abstract()}
    A = {"A": Num("X")}
    AB = {"A": Num("X"), "B": Num("X")}

    def i(num: str, level: int = 0):
        return ("i", num, level)

    def so(value: int, level: int = 0) -> dict:
        return {level: (value, None)}

    roman4 = Level("upperRoman", "%1.", start=4)
    out = [
        # -- override ----------------------------------------------------------------------
        Case("override", "start 7", X, {"B": Num("X", so(7))}, (i("B"), i("B"), i("B"))),
        Case("override", "start 5", X, {"B": Num("X", so(5))}, (i("B"), i("B"))),
        Case("override", "after another instance", X, {"A": Num("X"), "B": Num("X", so(5))},
             (i("A"), i("A"), i("B"), i("B"), i("A"), i("A"))),
        Case("override", "restart at 1 after another", X, {"A": Num("X"), "B": Num("X", so(1))},
             (i("A"), i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("override", "used again after another", X, {"A": Num("X"), "B": Num("X", so(5))},
             (i("B"), i("A"), i("B"), i("A"), i("B"))),
        Case("override", "then an instance without", X, {"B": Num("X", so(5)), "C": Num("X")},
             (i("B"), i("B"), i("C"), i("C"), i("B"))),
        Case("override", "two overridden, interleaved", X, {"B": Num("X", so(5)), "C": Num("X", so(10))},
             (i("B"), i("C"), i("B"), i("C"))),
        Case("override", "at level 1", X, {"B": Num("X", so(3, 1))},
             (i("B"), i("B", 1), i("B", 1), i("B"), i("B", 1), i("B", 1))),
        Case("override", "at level 0, deeper items shared", X, {"A": Num("X"), "B": Num("X", so(5))},
             (i("A"), i("A", 1), i("A", 1), i("B", 1), i("B"), i("B", 1), i("A", 1), i("A"))),
        Case("override", "first item deeper than the override", X, {"B": Num("X", so(5))},
             (i("B", 1), i("B"), i("B", 1), i("B"))),
        Case("override", "a whole level", X, {"A": Num("X"), "B": Num("X", {0: (None, roman4)})},
             (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("override", "a whole level and a start", X, {"A": Num("X"), "B": Num("X", {0: (9, roman4)})},
             (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("override", "start over an abstract start", {"X": Abstract((Level("decimal", "%1.", start=3),
                                                                         *DEFAULT_LEVELS[1:]))},
             {"A": Num("X"), "B": Num("X", so(5))}, (i("A"), i("A"), i("B"), i("A"))),
        Case("override", "at level 1, then an instance without", X, {"A": Num("X"), "B": Num("X", so(3, 1))},
             (i("B"), i("B", 1), i("B", 1), i("A"), i("A", 1), i("B", 1))),
        Case("override", "a whole level at level 1", X,
             {"B": Num("X", {1: (None, Level("lowerLetter", "%2.", start=3))})},
             (i("B"), i("B", 1), i("B", 1), i("B"), i("B", 1))),
        # -- instance ----------------------------------------------------------------------
        Case("instance", "two, no override", X, AB, (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("instance", "two, interleaved", X, AB, (i("A"), i("B"), i("A"), i("B"))),
        Case("instance", "level 0 of the other between level 1", X, AB,
             (i("A"), i("A", 1), i("A", 1), i("B"), i("A", 1), i("B", 1))),
        Case("instance", "two definitions alike", {"X": Abstract(), "Y": Abstract()},
             {"A": Num("X"), "B": Num("Y")}, (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("instance", "two definitions, one nsid", {"X": Abstract(nsid="1A2B3C4D"), "Y": Abstract(nsid="1A2B3C4D")},
             {"A": Num("X"), "B": Num("Y")}, (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("instance", "one nsid, other levels",
             {"X": Abstract(nsid="2B3C4D5E"),
              "Y": Abstract((Level("upperRoman", "%1."), *DEFAULT_LEVELS[1:]), nsid="2B3C4D5E")},
             {"A": Num("X"), "B": Num("Y")}, (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("instance", "numStyleLink", {"X": Abstract(link="style"), "Y": Abstract(levels=(), link="num")},
             {"A": Num("X"), "B": Num("Y")}, (i("A"), i("A"), i("B"), i("B"), i("A"))),
        Case("instance", "numId 0 between", X, A, (i("A"), ("x", "Off"), i("A"), ("p", "Plain"), i("A"))),
        # -- restart -----------------------------------------------------------------------
        Case("restart", "deeper reset", X, A, (i("A"), i("A", 1), i("A", 1), i("A"), i("A", 1))),
        Case("restart", "lvlRestart 0", {"X": Abstract((DEFAULT_LEVELS[0], Level("lowerLetter", "%2.", restart=0),
                                                        DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A", 1), i("A", 1), i("A"), i("A", 1))),
        Case("restart", "lvlRestart 1 at level 2", {"X": Abstract((*DEFAULT_LEVELS[:2],
                                                                   Level("lowerRoman", "%3.", restart=1)))},
             A, (i("A"), i("A", 1), i("A", 2), i("A", 2), i("A", 1), i("A", 2), i("A"), i("A", 1), i("A", 2))),
        Case("restart", "a level skipped", {"X": Abstract(LEGAL_LEVELS)}, A,
             (i("A"), i("A"), i("A", 2), i("A", 2), i("A", 1), i("A", 2))),
        Case("restart", "first item at level 1", {"X": Abstract(LEGAL_LEVELS)}, A, (i("A", 1), i("A", 1), i("A"))),
        Case("restart", "first item at level 1, no %1", X, A, (i("A", 1), i("A", 1), i("A"))),
        Case("restart", "reset level skipped after", {"X": Abstract(LEGAL_LEVELS)}, A,
             (i("A"), i("A", 1), i("A", 1), i("A"), i("A", 2))),
        Case("restart", "level start 3", {"X": Abstract((DEFAULT_LEVELS[0], Level("lowerLetter", "%2.", start=3),
                                                         DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A", 1), i("A", 1), i("A"), i("A", 1))),
        # -- legal -------------------------------------------------------------------------
        Case("legal", "roman parent", {"X": Abstract((Level("upperRoman", "%1."), Level("decimal", "%1.%2."),
                                                      DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A", 1), i("A"), i("A", 1))),
        Case("legal", "roman parent, isLgl", {"X": Abstract((Level("upperRoman", "%1."),
                                                             Level("decimal", "%1.%2.", legal=True),
                                                             DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A", 1), i("A"), i("A", 1))),
        Case("legal", "letters, isLgl on a letter level",
             {"X": Abstract((Level("lowerLetter", "%1."), Level("upperLetter", "%1.%2.", legal=True),
                             DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A", 1), i("A", 1))),
        Case("legal", "isLgl on level 0", {"X": Abstract((Level("upperRoman", "%1.", legal=True),
                                                          Level("lowerLetter", "%1.%2."), DEFAULT_LEVELS[2]))},
             A, (i("A"), i("A"), i("A", 1))),
        Case("legal", "three levels", {"X": Abstract((Level("upperLetter", "%1."), Level("lowerRoman", "%1.%2."),
                                                      Level("decimal", "%1.%2.%3.")))},
             A, (i("A"), i("A", 1), i("A", 2), i("A", 2), i("A", 1), i("A", 2))),
        # -- story -------------------------------------------------------------------------
        Case("story", "table cells", X, A,
             (i("A"), ("t", [[i("A"), i("A")], [i("A")]]), i("A"))),
        Case("story", "text box", X, A, (i("A"), ("box", [i("A"), i("A")]), i("A"), i("A"))),
        Case("story", "footnote", X, A, (i("A"), ("note", [i("A"), i("A")]), i("A"), i("A"))),
        Case("story", "two footnotes", X, A,
             (i("A"), ("note", [i("A"), i("A")]), i("A"), ("note", [i("A")]), i("A"))),
        Case("story", "two text boxes", X, A,
             (i("A"), ("box", [i("A"), i("A")]), i("A"), ("box", [i("A")]), i("A"))),
        Case("story", "header", X, A, (i("A"), i("A")), header=(i("A"), i("A"))),
        Case("story", "header, second page", {}, {}, (("p", "Second page under the header"),)),
    ]
    return tuple(out)


CASES = _cases()
#: The first case in the header's section.
HEADER_CASE = next(k for k, case in enumerate(CASES) if case.header is not None)


class Ids:
    """Document-wide abstractNumIds and numIds for each case's local names."""

    def __init__(self) -> None:
        self.abstract: dict = {}
        self.num: dict = {}
        for number, case in enumerate(CASES):
            for name in case.abstracts:
                self.abstract[number, name] = len(self.abstract)
            for name in case.nums:
                self.num[number, name] = len(self.num) + 1


IDS = Ids()


def _nsid(number: int, name: str) -> str:
    return f"{0x10000000 + number * 64 + IDS.abstract[number, name]:08X}"


def numbering() -> str:
    abstracts = nums = ""
    for number, case in enumerate(CASES):
        for name, abstract in case.abstracts.items():
            nsid = abstract.nsid or _nsid(number, name)
            body = f'<w:nsid w:val="{nsid}"/><w:multiLevelType w:val="multilevel"/>'
            if abstract.link == "style":
                body += f'<w:styleLink w:val="{LINKED_STYLE}"/>'
            elif abstract.link == "num":
                body += f'<w:numStyleLink w:val="{LINKED_STYLE}"/>'
            body += "".join(level.xml(k) for k, level in enumerate(abstract.levels))
            abstracts += f'<w:abstractNum w:abstractNumId="{IDS.abstract[number, name]}">{body}</w:abstractNum>'
        for name, num in case.nums.items():
            body = f'<w:abstractNumId w:val="{IDS.abstract[number, num.abstract]}"/>'
            for ilvl, (start, level) in sorted(num.overrides.items()):
                inner = (f'<w:startOverride w:val="{start}"/>' if start is not None else "")
                inner += level.xml(ilvl) if level is not None else ""
                body += f'<w:lvlOverride w:ilvl="{ilvl}">{inner}</w:lvlOverride>'
            nums += f'<w:num w:numId="{IDS.num[number, name]}">{body}</w:num>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:numbering xmlns:w="{probe_docx.W_NS}">{abstracts}{nums}</w:numbering>')


def linked_num() -> int:
    """The numId the numbering style names: the ``numStyleLink`` case's ``A`` (over the
    definition that carries ``w:styleLink``)."""
    number = next(k for k, case in enumerate(CASES) if case.name == "numStyleLink")
    return IDS.num[number, "A"]


def p(text: str, **props) -> str:
    return wml.paragraph(wml.run(text), mark={}, **props)


class Builder:
    def __init__(self) -> None:
        self.notes: list[str] = []
        self.drawings = 0

    def item(self, number: int, num: str, level: int) -> str:
        return p(f"{num}{level}", numPr=f'<w:ilvl w:val="{level}"/><w:numId w:val="{IDS.num[number, num]}"/>')

    def blocks(self, number: int, blocks) -> str:
        out = ""
        for block in blocks:
            kind = block[0]
            if kind == "i":
                out += self.item(number, block[1], block[2])
            elif kind == "p":
                out += p(block[1])
            elif kind == "x":
                out += p(block[1], numPr='<w:ilvl w:val="0"/><w:numId w:val="0"/>')
            elif kind == "t":
                cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>'
                                f"{self.blocks(number, content)}</w:tc>" for content in block[1])
                out += ('<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
                        ' w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid>'
                        + '<w:gridCol w:w="3600"/>' * len(block[1]) + f"</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>")
            elif kind == "box":
                self.drawings += 1
                content = self.blocks(number, block[1])
                shape = (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>{xfrm(0, 0, 2400000, 900000)}'
                         f'{preset("rect")}<a:noFill/>{NO_LINE}</wps:spPr>'
                         f"<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>"
                         '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440"'
                         ' bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>')
                anchor = Anchor(("margin", "offset", 3200000), ("paragraph", "offset", 1800000), 2400000, 900000,
                                graphic=graphic(shape, "wps"))
                out += wml.paragraph(wml.run("Anchor") + anchor.xml(self.drawings), mark={})
            elif kind == "note":
                note_id = len(self.notes) + 1
                ref = f'<w:r>{wml.rpr(vertAlign="superscript")}<w:footnoteReference w:id="{note_id}"/></w:r>'
                first = (f'<w:p><w:r>{wml.rpr(vertAlign="superscript")}<w:footnoteRef/></w:r>'
                         f'{wml.run(" Note")}</w:p>')
                self.notes.append(f'<w:footnote w:id="{note_id}">{first}{self.blocks(number, block[1])}</w:footnote>')
                out += wml.paragraph(wml.run("Noted") + ref, mark={})
            else:
                raise ValueError(kind)
        return out


def parts() -> tuple[str, str, str, str]:
    """The body, the final section, the footnotes and the header."""
    builder = Builder()
    body = ""
    header = ""
    for number, case in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number != HEADER_CASE else {}
        content = p(f"Case {number}", **props) + builder.blocks(number, case.blocks)
        if case.header is not None:
            header = builder.blocks(number, case.header)
        if number == HEADER_CASE - 1:
            # The last paragraph ends the section before the header's.
            head, _, tail = content.rpartition("</w:pPr>")
            content = head + plain_section() + "</w:pPr>" + tail
        body += content
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<w:footnotes xmlns:w="{probe_docx.W_NS}">{notes_probe.separators()}'
                 + "".join(builder.notes) + "</w:footnotes>")
    return body, header_section(), footnotes, header


def plain_section() -> str:
    return story_docx.section()


#: The header part's relationship id: ``probe_docx.package`` numbers extra parts from 2,
#: and the header is the first.
HEADER_RID = "rId2"


def header_section() -> str:
    return story_docx.section(references=f'<w:headerReference w:type="default" r:id="{HEADER_RID}"/>')


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("numbering", LINKED_STYLE, name="Linked List",
                   ppr_={"numPr": f'<w:numId w:val="{linked_num()}"/>'})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    body, final, footnotes, header = parts()
    extra = [("word/header1.xml", story_docx.HEADER_CONTENT_TYPE, story_docx.HEADER_REL,
              story_docx.story_part("header", header)),
             ("word/numbering.xml", NUMBERING_CONTENT_TYPE, NUMBERING_REL, numbering()),
             ("word/footnotes.xml", notes_probe.FOOTNOTES_CONTENT_TYPE, notes_probe.FOOTNOTES_REL, footnotes)]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body, final_section=final, styles=styles(), extra_parts=tuple(extra))


DOCUMENTS = tuple(f"numbering-{setting}" for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"numbering-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
    print(len(CASES), "cases")
