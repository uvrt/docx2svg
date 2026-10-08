#!/usr/bin/env python3
"""Endnotes: where Word puts them, how it numbers them, and how it sets their text.

An endnote reference (``w:endnoteReference``) draws its note's number where it stands,
and the note (``word/endnotes.xml``) goes after the text: at the end of the document by
default, or at the end of each section (``w:endnotePr/w:pos`` ``sectEnd``, a setting of
the document's), under a separator (the ``separator`` note's paragraph) and, where the
notes go on to another page, a continuation separator.  The number's format, first value
and restart are ``w:endnotePr``'s, the document's (in ``settings.xml``) or a section's.
None of it was laid out.

The notes' formatting is stated as Word writes it: every note paragraph is
``EndnoteText`` (Calibri 10 pt, single, no spacing) opening with its number
(``w:endnoteRef``) in ``EndnoteReference`` (superscript), and the separators are one
paragraph each (single, no space after).  The body is Calibri 11 pt, A4, the left margin
off the pixel grid.  Documents (``DOCUMENTS``), each in three settings -- no
``settings.xml``, mode 14 and mode 15 -- unless it needs the settings part:

* ``base`` -- references in running text (mid-line, near a line's end, adjacent, in a
  24 pt paragraph, one not in ``EndnoteReference``), and notes of one line, three lines,
  two paragraphs, one in ``Normal``, one with space before and after, one whose number
  is not superscript; the notes after the last paragraph, on its page;
* ``overflow`` -- a page nearly full, then notes that go on to the next page (and a
  three-line note split there): the continuation separator;
* ``format`` -- three sections (``nextPage``), each with its own ``w:endnotePr``:
  ``decimal`` from 5, ``upperRoman`` restarting each section, ``lowerLetter`` from 3
  restarting each section; the notes at the end;
* ``sectend`` (modes 14 and 15 only) -- ``w:pos`` ``sectEnd``: three sections, the
  second ending in a ``continuous`` break, each section's notes after its text;
* ``docpr`` (modes 14 and 15 only) -- the document's ``w:endnotePr``: ``upperLetter``
  from 2, one section overriding it with ``decimal``;
* ``separators`` -- ``overflow``'s notes under separators formatted as a document may
  format them: the separator's paragraph at 20 pt Times New Roman, indented 720 and with
  120 before, the continuation separator's at 8 pt, right-aligned (which Word, with no
  settings part, was found to ignore);
* ``normal`` -- a Normal style as a filesamples document has it (12 pt, every first line
  indented 432, ``auto`` 276 by default), ``EndnoteText`` single: what the separators and
  the notes take from it;
* ``unnamed`` (modes 14 and 15 only) -- ``separators`` with a settings part whose
  ``w:endnotePr`` does not name the separator notes.

Reader: ``read_endnote_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
SETTINGS = {"none": None, "14": 14, "15": 15}
KINDS = ("base", "overflow", "format", "sectend", "docpr", "separators", "normal", "unnamed")
#: The kinds that need a settings part (the document's ``w:endnotePr``).
SETTINGS_ONLY = ("sectend", "docpr", "unnamed")
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}
ENDNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
ENDNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"

WORDS = ("amber", "birch", "cedar", "delta", "ember", "fjord", "grove", "heron", "inlet", "juniper", "kestrel",
         "larch", "maple", "nectar", "osprey", "plover", "quartz", "raven", "sable", "thistle", "umber",
         "vireo", "willow", "yarrow", "zephyr")


def words(start: int, count: int) -> str:
    return " ".join(WORDS[(start + k) % len(WORDS)] for k in range(count))


def text(value: str, **props) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{probe_docx.escape(value)}</w:t></w:r>'


def reference(note: int, styled: bool = True, **props) -> str:
    rpr = wml.rpr(rStyle="EndnoteReference", **props) if styled else wml.rpr(**props)
    return f"<w:r>{rpr}<w:endnoteReference w:id=\"{note}\"/></w:r>"


@dataclass(frozen=True)
class Note:
    #: Its paragraphs' texts.
    paragraphs: tuple[str, ...]
    style: str | None = "EndnoteText"
    spacing: dict | None = None
    #: Whether its number is in ``EndnoteReference``.
    styled: bool = True

    def xml(self, note_id: int) -> str:
        out = ""
        for k, value in enumerate(self.paragraphs):
            props = {}
            if self.style:
                props["pStyle"] = self.style
            if self.spacing:
                props["spacing"] = self.spacing
            number = ""
            if k == 0:
                rpr = wml.rpr(rStyle="EndnoteReference") if self.styled else ""
                number = f"<w:r>{rpr}<w:endnoteRef/></w:r>"
            out += f"<w:p>{wml.ppr(**props)}{number}{text((' ' if k == 0 else '') + value)}</w:p>"
        return f'<w:endnote w:id="{note_id}">{out}</w:endnote>'


def separators(styled: bool = False) -> str:
    """The two separators, as Word writes them; ``styled``: formatted (``separators``)."""
    out = ""
    for kind, note_id in (("separator", -1), ("continuationSeparator", 0)):
        ppr = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        rpr = ""
        if styled and kind == "separator":
            ppr = '<w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="720"/>'
            rpr = wml.rpr(rFonts={k: "Times New Roman" for k in FACE}, sz=40, szCs=40)
        elif styled:
            ppr += '<w:jc w:val="right"/>'
            rpr = wml.rpr(sz=16, szCs=16)
        out += (f'<w:endnote w:type="{kind}" w:id="{note_id}"><w:p><w:pPr>{ppr}{rpr}</w:pPr>'
                f"<w:r>{rpr}<w:{kind}/></w:r></w:p></w:endnote>")
    return out


def _p(runs: str, **props) -> str:
    props.setdefault("spacing", SPACING)
    return wml.paragraph(runs, mark={}, **props)


def section(endnote_pr: str = "", kind: str | None = None) -> str:
    m = MARGINS
    kind_xml = f'<w:type w:val="{kind}"/>' if kind else ""
    return (f"<w:sectPr>{endnote_pr}{kind_xml}<w:pgSz w:w=\"{PAGE['w']}\" w:h=\"{PAGE['h']}\"/>"
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def endnote_pr(fmt: str | None = None, start: int | None = None, restart: str | None = None,
               pos: str | None = None, special: bool = False) -> str:
    inner = ""
    if pos:
        inner += f'<w:pos w:val="{pos}"/>'
    if fmt:
        inner += f'<w:numFmt w:val="{fmt}"/>'
    if start is not None:
        inner += f'<w:numStart w:val="{start}"/>'
    if restart:
        inner += f'<w:numRestart w:val="{restart}"/>'
    if special:
        inner += '<w:endnote w:id="-1"/><w:endnote w:id="0"/>'
    return f"<w:endnotePr>{inner}</w:endnotePr>" if inner else ""


def content(kind: str) -> tuple[str, list[Note], str, str]:
    """``(body, notes, the final section, the document's endnotePr)``."""
    notes: list[Note] = []

    def note(value: Note, styled: bool = True, **props) -> str:
        notes.append(value)
        return reference(len(notes), styled, **props)

    body = ""
    final = section()
    doc_pr = ""
    if kind == "base":
        body += _p(text("Endnotes base"))
        body += _p(text("Alpha") + note(Note(("one line " + words(0, 4),))) + text(" " + words(1, 6) + "."))
        body += _p(text(words(2, 9) + " " + words(5, 2)) + note(Note((words(3, 40),)))
                   + text(" " + words(4, 30)) + note(Note(("first " + words(6, 5), "second " + words(7, 12))))
                   + text(" " + words(8, 20) + "."))
        body += _p(text("Large " + words(9, 5), sz=48, szCs=48) + note(Note(("in Normal " + words(10, 5),), style=None),
                                                                        sz=48, szCs=48)
                   + text(" " + words(11, 3), sz=48, szCs=48))
        body += _p(text("Plain") + note(Note(("spaced " + words(12, 6),),
                                             spacing={"before": 120, "after": 120, "line": 240, "lineRule": "auto"}),
                                        styled=False) + text(" reference, not superscript."))
        body += _p(text("Adjacent") + note(Note(("unstyled number " + words(13, 4),), styled=False))
                   + note(Note(("next " + words(14, 3),))) + text(" " + words(15, 5) + "."))
        for k in range(6):
            body += _p(text(f"Filler {k} " + words(16 + k, 7)) + (note(Note((f"note {k} " + words(k, 8),)))
                                                                  if k % 2 else ""))
        body += _p(text("The end."))
    elif kind in ("overflow", "separators", "unnamed"):
        body += _p(text(f"Endnotes {kind}"))
        for k in range(40):
            body += _p(text(f"Line {k} " + words(k, 6)))
        refs = "".join(note(Note((f"long {k} " + words(k, 30 if k % 3 else 50),))) for k in range(8))
        body += _p(text("References") + refs + text(" end."))
    elif kind == "format":
        specs = ((dict(fmt="decimal", start=5), "nextPage"), (dict(fmt="upperRoman", restart="eachSect"), "nextPage"),
                 (dict(fmt="lowerLetter", start=3, restart="eachSect"), None))
        for k, (pr, _) in enumerate(specs):
            body += _p(text(f"Section {k}") + note(Note((f"section {k} first " + words(k, 5),))) + text(" and")
                       + note(Note((f"section {k} second " + words(k + 3, 5),))) + text("."))
            if k < len(specs) - 1:
                body += _p(text(f"End of section {k}"), sect=section(endnote_pr(**pr), specs[k + 1][1]))
            else:
                body += _p(text(f"End of section {k}"))
                final = section(endnote_pr(**pr))
    elif kind == "sectend":
        doc_pr = endnote_pr(pos="sectEnd", special=True)
        breaks = ("continuous", "nextPage")
        for k in range(3):
            body += _p(text(f"Section {k} " + words(k, 6)) + note(Note((f"section {k} " + words(k + 1, 6),)))
                       + text(" " + words(k + 2, 4)) + note(Note((f"section {k} again " + words(k + 5, 20),))))
            body += _p(text(f"After section {k} " + words(k + 7, 5)),
                       **({"sect": section(kind=breaks[k])} if k < 2 else {}))
    elif kind == "normal":
        body += wml.paragraph(text("Endnotes under a Normal style " + words(0, 20)) + note(Note(("normal " + words(1, 30),)))
                              + text(" " + words(2, 12)) + note(Note(("again " + words(3, 4),))) + text("."), mark={})
        body += wml.paragraph(text("The end " + words(4, 6) + "."), mark={})
    elif kind == "docpr":
        doc_pr = endnote_pr(fmt="upperLetter", start=2, special=True)
        body += _p(text("Document format") + note(Note(("document " + words(0, 5),))) + note(Note(("again",))))
        body += _p(text("End of section 0"), sect=section(endnote_pr(fmt="decimal"), "nextPage"))
        body += _p(text("Section 1") + note(Note(("section 1 " + words(3, 5),))) + text("."))
        body += _p(text("The end."))
    return body, notes, final, doc_pr


def build(name: str) -> bytes:
    """``name``: ``endnote-<kind>-<setting>``."""
    _, kind, setting = name.split("-")
    body, notes, final, doc_pr = content(kind)
    normal = kind == "normal"
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True,
                   ppr_={"ind": {"firstLine": 432}} if normal else None, rpr_={"sz": 24} if normal else None),
         wml.style("paragraph", "EndnoteText", based_on="Normal",
                   ppr_={"spacing": {"line": 240, "lineRule": "auto"}} if normal else {"spacing": SPACING},
                   rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "EndnoteReference", rpr_={"vertAlign": "superscript"})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"line": 276, "lineRule": "auto"}} if normal else {"spacing": SPACING},
    )
    endnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:endnotes xmlns:w="{probe_docx.W_NS}">{separators(kind in ("separators", "unnamed"))}'
                + "".join(n.xml(k + 1) for k, n in enumerate(notes)) + "</w:endnotes>")
    extra = [("word/endnotes.xml", ENDNOTES_CONTENT_TYPE, ENDNOTES_REL, endnotes)]
    mode = SETTINGS[setting]
    if mode is not None:
        name_, content_type, rel, xml = wml.settings_part({"val": "en-GB"}, compatibility_mode=mode)
        if kind != "unnamed":
            xml = xml.replace("<w:compat>", (doc_pr or endnote_pr(special=True)) + "<w:compat>")
        extra.append((name_, content_type, rel, xml))
    return probe_docx.package(body, final_section=final, styles=styles, extra_parts=tuple(extra))


DOCUMENTS = tuple(f"endnote-{kind}-{setting}" for kind in KINDS for setting in SETTINGS
                  if not (kind in SETTINGS_ONLY and SETTINGS[setting] is None))


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)
