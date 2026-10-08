#!/usr/bin/env python3
"""Footnotes drawn: their references' numbers, the notes at the page foot, the separators.

``make_footnote_probe.py`` measured the room a footnote takes (ROADMAP.md 4.3); nothing of
a footnote was drawn.  A reference (``w:footnoteReference``) draws its note's number where
it stands, and the note (``word/footnotes.xml``) is set at the foot of the reference's
page, under a separator, opening with its own number (``w:footnoteRef``); where a note
goes on to the next page, that page's notes open with the continuation separator.  The
number's format, first value and restart (``continuous``, ``eachSect``, ``eachPage``) are
``w:footnotePr``'s, the document's (in ``settings.xml``) or a section's; ``w:pos``
(``pageBottom``, ``beneathText``) says where the notes go.

The notes' formatting is stated as Word writes it: every note paragraph is
``FootnoteText`` (Calibri 10 pt, single, no spacing) opening with its number in
``FootnoteReference`` (superscript), and the separators are one paragraph each (single, no
space after).  The body is Calibri 11 pt, A4, the left margin off the pixel grid.
Documents (``DOCUMENTS``), each in four settings -- no ``settings.xml``, a settings part
that names the separators but states no compatibility mode (``named``), mode 14 and mode
15 -- unless it needs the settings part:

* ``base`` -- a short page: references in running text (mid-line, near a line's end,
  adjacent, in a 24 pt paragraph, one not in ``FootnoteReference``), and notes of one
  line, three lines, two paragraphs, one in ``Normal``, one with space before and after,
  one whose number is not superscript;
* ``overflow`` -- a page nearly full, references late on it with long notes, one too
  long for the page: its rest goes to the next page, under the continuation separator;
* ``format`` -- four sections (``nextPage``), each with its own ``w:footnotePr``:
  ``decimal`` from 5, ``upperRoman`` restarting each section, ``lowerLetter`` from 3
  restarting each section, ``chicago`` going on;
* ``chain`` -- four sections: one stating nothing, one ``decimal`` from 7 going on, one
  restarting, one stating nothing: where a section that goes on counts from;
* ``eachpage`` -- ``w:numRestart`` ``eachPage`` in the section: three pages of two and
  three references each, the later pages' references past ``9`` were they continuous;
* ``docpr`` (settings only) -- the document's ``w:footnotePr``: ``upperLetter`` from 2,
  a second section overriding it with ``decimal``;
* ``docpage`` (settings only) -- the document's ``w:footnotePr`` restarting each page,
  the sections stating nothing;
* ``separators`` -- ``overflow``'s notes under separators formatted as a document may
  format them: the separator's paragraph at 20 pt Times New Roman, indented 720 and with
  120 before, the continuation separator's at 8 pt, right-aligned;
* ``unnamed`` (settings only) -- ``separators`` with a settings part whose
  ``w:footnotePr`` names no separator note;
* ``normal`` -- a Normal style as a filesamples document has it (12 pt, every first line
  indented 432, ``auto`` 276 by default), ``FootnoteText`` single with no first-line
  indent: what the separators and the notes take from it;
* ``beneath`` (settings only) -- ``w:pos`` ``beneathText`` in the document's
  ``w:footnotePr``: ``base``'s short page, the notes under its text;
* ``sectbeneath`` -- the same stated in the section's ``w:footnotePr``;
* ``after`` -- pages whose last note has space after it: two paragraphs each 160 after
  (Word 365's default), 120 before and 240 after, 240 before and none after, and a spaced
  note followed by a plain one.

Reader: ``read_footnote_draw_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
#: Setting -> (whether a settings part is written, its compatibility mode).
SETTINGS = {"none": (False, None), "named": (True, None), "14": (True, 14), "15": (True, 15)}
KINDS = ("base", "overflow", "format", "chain", "eachpage", "docpr", "docpage", "separators", "unnamed", "normal",
         "beneath", "sectbeneath", "after")
#: The kinds that need a settings part (the document's ``w:footnotePr``).
SETTINGS_ONLY = ("docpr", "docpage", "unnamed", "beneath")
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}
FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
FOOTNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

WORDS = ("amber", "birch", "cedar", "delta", "ember", "fjord", "grove", "heron", "inlet", "juniper", "kestrel",
         "larch", "maple", "nectar", "osprey", "plover", "quartz", "raven", "sable", "thistle", "umber",
         "vireo", "willow", "yarrow", "zephyr")


def words(start: int, count: int) -> str:
    return " ".join(WORDS[(start + k) % len(WORDS)] for k in range(count))


def text(value: str, **props) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{probe_docx.escape(value)}</w:t></w:r>'


def reference(note: int, styled: bool = True, **props) -> str:
    rpr = wml.rpr(rStyle="FootnoteReference", **props) if styled else wml.rpr(**props)
    return f"<w:r>{rpr}<w:footnoteReference w:id=\"{note}\"/></w:r>"


@dataclass(frozen=True)
class Note:
    #: Its paragraphs' texts.
    paragraphs: tuple[str, ...]
    style: str | None = "FootnoteText"
    spacing: dict | None = None
    #: Whether its number is in ``FootnoteReference``.
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
                rpr = wml.rpr(rStyle="FootnoteReference") if self.styled else ""
                number = f"<w:r>{rpr}<w:footnoteRef/></w:r>"
            out += f"<w:p>{wml.ppr(**props)}{number}{text((' ' if k == 0 else '') + value)}</w:p>"
        return f'<w:footnote w:id="{note_id}">{out}</w:footnote>'


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
        out += (f'<w:footnote w:type="{kind}" w:id="{note_id}"><w:p><w:pPr>{ppr}{rpr}</w:pPr>'
                f"<w:r>{rpr}<w:{kind}/></w:r></w:p></w:footnote>")
    return out


def _p(runs: str, **props) -> str:
    props.setdefault("spacing", SPACING)
    return wml.paragraph(runs, mark={}, **props)


def section(footnote_pr: str = "", kind: str | None = None) -> str:
    m = MARGINS
    kind_xml = f'<w:type w:val="{kind}"/>' if kind else ""
    return (f"<w:sectPr>{footnote_pr}{kind_xml}<w:pgSz w:w=\"{PAGE['w']}\" w:h=\"{PAGE['h']}\"/>"
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def footnote_pr(fmt: str | None = None, start: int | None = None, restart: str | None = None,
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
        inner += '<w:footnote w:id="-1"/><w:footnote w:id="0"/>'
    return f"<w:footnotePr>{inner}</w:footnotePr>" if inner else ""


def content(kind: str) -> tuple[str, list[Note], str, str]:
    """``(body, notes, the final section, the document's footnotePr)``."""
    notes: list[Note] = []

    def note(value: Note, styled: bool = True, **props) -> str:
        notes.append(value)
        return reference(len(notes), styled, **props)

    body = ""
    final = section()
    doc_pr = ""
    if kind in ("base", "beneath", "sectbeneath"):
        if kind == "beneath":
            doc_pr = footnote_pr(pos="beneathText", special=True)
        elif kind == "sectbeneath":
            final = section(footnote_pr(pos="beneathText"))
        body += _p(text(f"Footnotes {kind}"))
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
        body += _p(text("The end."))
    elif kind in ("overflow", "separators", "unnamed"):
        body += _p(text(f"Footnotes {kind}"))
        for k in range(38):
            # A six-line note, then one of some twenty-five lines that cannot all stay.
            size = {24: 90, 30: 380}.get(k)
            body += _p(text(f"Line {k} " + words(k, 6))
                       + (note(Note((f"long {k} " + words(k, size),))) if size else ""))
        for k in range(38, 60):
            body += _p(text(f"Line {k} " + words(k, 6))
                       + (note(Note((f"later {k} " + words(k, 12),))) if k in (40, 50) else ""))
    elif kind == "format":
        specs = ((dict(fmt="decimal", start=5), "nextPage"), (dict(fmt="upperRoman", restart="eachSect"), "nextPage"),
                 (dict(fmt="lowerLetter", start=3, restart="eachSect"), "nextPage"), (dict(fmt="chicago"), None))
        for k, (pr, _) in enumerate(specs):
            body += _p(text(f"Section {k}") + note(Note((f"section {k} first " + words(k, 5),))) + text(" and")
                       + note(Note((f"section {k} second " + words(k + 3, 5),))) + text(" more")
                       + note(Note((f"section {k} third " + words(k + 6, 3),))) + text("."))
            if k < len(specs) - 1:
                body += _p(text(f"End of section {k}"), sect=section(footnote_pr(**pr), specs[k + 1][1]))
            else:
                body += _p(text(f"End of section {k}"))
                final = section(footnote_pr(**pr))
    elif kind == "chain":
        specs = ((dict(), "nextPage"), (dict(fmt="decimal", start=7), "nextPage"),
                 (dict(restart="eachSect"), "nextPage"), (dict(), None))
        for k, (pr, _) in enumerate(specs):
            body += _p(text(f"Section {k}") + note(Note((f"section {k} first " + words(k, 5),))) + text(" and")
                       + note(Note((f"section {k} second " + words(k + 3, 5),))) + text("."))
            if k < len(specs) - 1:
                body += _p(text(f"End of section {k}"), sect=section(footnote_pr(**pr), specs[k + 1][1]))
            else:
                body += _p(text(f"End of section {k}"))
                final = section(footnote_pr(**pr))
    elif kind in ("eachpage", "docpage"):
        if kind == "docpage":
            doc_pr = footnote_pr(restart="eachPage", special=True)
        else:
            final = section(footnote_pr(restart="eachPage"))
        count = 0
        for page, refs in enumerate((2, 3, 9, 2)):
            for k in range(refs):
                props = {"pageBreakBefore": True} if page and not k else {}
                body += _p(text(f"Page {page} reference {k} " + words(count, 4))
                           + note(Note((f"page {page} note {k} " + words(count + 2, 3),))) + text("."), **props)
                count += 1
    elif kind == "normal":
        body += wml.paragraph(text("Footnotes under a Normal style " + words(0, 20)) + note(Note(("normal " + words(1, 30),)))
                              + text(" " + words(2, 12)) + note(Note(("again " + words(3, 4),))) + text("."), mark={})
        body += wml.paragraph(text("The end " + words(4, 6) + "."), mark={})
    elif kind == "after":
        word365 = {"before": 0, "after": 160, "line": 278, "lineRule": "auto"}
        pages = ((Note(("plain " + words(0, 4),)), Note(("first " + words(1, 6), "second " + words(2, 5)),
                                                        spacing=word365)),
                 (Note(("before and after " + words(3, 8),),
                       spacing={"before": 120, "after": 240, "line": 240, "lineRule": "auto"}),),
                 (Note(("before only " + words(4, 8),), spacing={"before": 240, "after": 0, "line": 240,
                                                                 "lineRule": "auto"}),),
                 (Note(("spaced " + words(5, 6), "again " + words(6, 4)), spacing=word365),
                  Note(("plain after it " + words(7, 4),))))
        for page, notes_ in enumerate(pages):
            props = {"pageBreakBefore": True} if page else {}
            body += _p(text(f"Space after, page {page} " + words(page, 5))
                       + "".join(note(n) + text(" " + words(page + k, 2)) for k, n in enumerate(notes_)) + text("."),
                       **props)
    elif kind == "docpr":
        doc_pr = footnote_pr(fmt="upperLetter", start=2, special=True)
        body += _p(text("Document format") + note(Note(("document " + words(0, 5),))) + note(Note(("again",))))
        body += _p(text("End of section 0"), sect=section(footnote_pr(fmt="decimal"), "nextPage"))
        body += _p(text("Section 1") + note(Note(("section 1 " + words(3, 5),))) + text("."))
        body += _p(text("The end."))
    return body, notes, final, doc_pr


def parts(name: str) -> tuple[str, str]:
    """``(kind, setting)`` of ``fndraw-<kind>-<setting>``."""
    _, kind, setting = name.split("-")
    return kind, setting


def build(name: str) -> bytes:
    """``name``: ``fndraw-<kind>-<setting>``."""
    kind, setting = parts(name)
    body, notes, final, doc_pr = content(kind)
    normal = kind == "normal"
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True,
                   ppr_={"ind": {"firstLine": 432}} if normal else None, rpr_={"sz": 24} if normal else None),
         wml.style("paragraph", "FootnoteText", based_on="Normal",
                   ppr_=({"spacing": {"line": 240, "lineRule": "auto"}, "ind": {"firstLine": 0}} if normal
                         else {"spacing": SPACING}),
                   rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "FootnoteReference", rpr_={"vertAlign": "superscript"})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"line": 276, "lineRule": "auto"}} if normal else {"spacing": SPACING},
    )
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<w:footnotes xmlns:w="{probe_docx.W_NS}">{separators(kind in ("separators", "unnamed"))}'
                 + "".join(n.xml(k + 1) for k, n in enumerate(notes)) + "</w:footnotes>")
    extra = [("word/footnotes.xml", FOOTNOTES_CONTENT_TYPE, FOOTNOTES_REL, footnotes)]
    written, mode = SETTINGS[setting]
    if written:
        name_, content_type, rel, xml = wml.settings_part({"val": "en-GB"}, compatibility_mode=mode)
        named = "" if kind == "unnamed" else (doc_pr or footnote_pr(special=True))
        xml = xml.replace("<w:compat>", named + "<w:compat>") if "<w:compat>" in xml else xml.replace(
            "<w:themeFontLang", named + "<w:themeFontLang")
        extra.append((name_, content_type, rel, xml))
    return probe_docx.package(body, final_section=final, styles=styles, extra_parts=tuple(extra))


DOCUMENTS = tuple(f"fndraw-{kind}-{setting}" for kind in KINDS for setting in SETTINGS
                  if not (kind in SETTINGS_ONLY and not SETTINGS[setting][0]))


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)
