#!/usr/bin/env python3
"""What Word draws for a field: which it computes, how it formats a page number, and in
which run's format.

Every field here carries a cached result Word could not have computed (``999``, ``1
January 2001``...), so the export shows which fields Word recomputes and which it draws
as cached.  Two documents.  **``computed``** holds what the renderer computes or draws as
cached and should draw as Word does; **``word``** holds what only Word can know -- the
fields it computes again on export and the number forms in the application's language --
and pins how far the renderer's drawing is from it.  Sections of two pages, each with a
footer of page-number fields:

* ``w:pgNumType`` -- ``@w:fmt`` decimal, upperRoman, lowerRoman, upperLetter,
  lowerLetter, numberInDash, decimalZero, ordinal, cardinalText, ordinalText, hex,
  chicago, and ``@w:start`` at values where a format changes shape (4, 49, 26 -> Z AA, 52
  -> ZZ AAA, 3999, 4000, 0);
* ``PAGE``, ``NUMPAGES`` (a ``w:fldSimple``) and ``SECTIONPAGES``, each also with the
  switches ``\\* roman``, ``Roman``, ``alphabetic``, ``ALPHABETIC``, ``Arabic``,
  ``ArabicDash``, ``Ordinal``, ``CardText``, ``OrdText``, ``Hex``, ``MERGEFORMAT``,
  ``\\# "000"`` and ``\\* roman \\* Upper``.

The header carries ``PAGE`` whose runs differ in format -- the ``begin`` run bold, the
result italic -- plain, with ``MERGEFORMAT`` and with ``CHARFORMAT``; one with no cached
result; a ``w:fldSimple`` whose result run is bold; ``NUMPAGES`` with a red instruction
and a blue result.

The body's first pages carry the fields a document holds besides page numbers (``DATE``,
``TIME``, ``AUTHOR``, ``FILENAME``, ``TITLE``, ``NUMWORDS``, ``REF``, ``PAGEREF``,
``SEQ``, ``=``, ``IF``, ``HYPERLINK``, ``DOCPROPERTY``, ``QUOTE``, an ``IF`` holding a
nested ``PAGE`` in its instruction, and a ``TOC`` whose result spans two paragraphs),
each with a wrong cached result; a hyperlink field with a 300-character instruction (if
the instruction were measured, the line would wrap); and, in the section numbered from
1,000,000, paragraphs whose last word is a ``PAGE`` or ``NUMPAGES`` field cached as ``1``
and filled so the computed value does not fit on the line where the cached one would.
"""

from __future__ import annotations

import functools

import story_docx
import wml
from story_docx import field, simple_field

#: The number forms Word writes in the application's language (``word``).
LANGUAGE_FORMATS = (
    ({"fmt": "ordinal", "start": 1}, "ordinal"),
    ({"fmt": "cardinalText", "start": 21}, "cardinalText"),
    ({"fmt": "ordinalText", "start": 12}, "ordinalText"),
)
FORMATS = (
    ({}, "decimal"),
    ({"fmt": "upperRoman", "start": 4}, "upperRoman"),
    ({"fmt": "lowerRoman", "start": 49}, "lowerRoman"),
    ({"fmt": "upperLetter", "start": 26}, "upperLetter"),
    ({"fmt": "lowerLetter", "start": 52}, "lowerLetter"),
    ({"fmt": "numberInDash", "start": 9}, "numberInDash"),
    ({"fmt": "decimalZero", "start": 7}, "decimalZero"),
    ({"fmt": "hex", "start": 255}, "hex"),
    ({"fmt": "chicago", "start": 1}, "chicago"),
    ({"fmt": "upperRoman", "start": 3999}, "upperRoman 3999"),
    ({"fmt": "upperRoman", "start": 4000}, "upperRoman 4000"),
    ({"start": 0}, "decimal 0"),
    ({"start": 1000000}, "wide"),
)

SWITCHES = (
    ("r", "PAGE \\* roman"), ("R", "PAGE \\* Roman"), ("a", "PAGE \\* alphabetic"), ("A", "PAGE \\* ALPHABETIC"),
    ("d", "PAGE \\* Arabic"), ("x", "PAGE \\* ArabicDash"), ("h", "PAGE \\* Hex"), ("m", "PAGE \\* MERGEFORMAT"),
    ("z", 'PAGE \\# "000"'), ("u", "PAGE \\* roman \\* Upper"), ("nr", "NUMPAGES \\* roman"),
    ("sR", "SECTIONPAGES \\* Roman"), ("l", "PAGE \\* Lower"), ("fc", "PAGE \\* alphabetic \\* FirstCap"),
    ("cs", "NUMPAGES \\* ALPHABETIC \\* Lower"),
)
#: The switches in the application's language (``word``).
LANGUAGE_SWITCHES = (("o", "PAGE \\* Ordinal"), ("c", "PAGE \\* CardText"), ("t", "PAGE \\* OrdText"))

BOLD = {"b": True}
ITALIC = {"i": True}


#: The switches whose results stay short at any value, for the sections numbered from
#: 3,999 up: there Word draws no footer paragraph holding a ``roman`` or ``alphabetic``
#: result of hundreds of letters (``\\* alphabetic`` of 3,999 is 154 letters; the whole
#: paragraph is left out) -- not measured further.
SHORT = (("d", "PAGE \\* Arabic"), ("x", "PAGE \\* ArabicDash"), ("h", "PAGE \\* Hex"),
         ("m", "PAGE \\* MERGEFORMAT"), ("z", 'PAGE \\# "000"'))


def footer(number: int, short: bool = False, bare: bool = False, switches=None) -> str:
    first = (wml.run(f"Sec{number} P[") + field("PAGE", "999") + wml.run("] N[")
             + simple_field("NUMPAGES", "888") + wml.run("] S[") + field("SECTIONPAGES", "777") + wml.run("]"))
    out = wml.paragraph(first, mark={}, pStyle="Footer")
    if bare:
        return out
    switches = switches or (SHORT if short else SWITCHES)
    for k in range(0, len(switches), 5):
        runs = ""
        for tag, instruction in switches[k:k + 5]:
            runs += wml.run(f" {tag}[") + field(instruction, "999") + wml.run("]")
        out += wml.paragraph(runs, mark={}, pStyle="Footer")
    return out


def header() -> str:
    runs = (wml.run("f1[") + field("PAGE", "9", rpr=BOLD, result_rpr=ITALIC) + wml.run("] f2[")
            + field("PAGE \\* MERGEFORMAT", "9", rpr=BOLD, result_rpr=ITALIC) + wml.run("] f3[")
            + field("PAGE \\* CHARFORMAT", "9", rpr=BOLD, result_rpr=ITALIC) + wml.run("] f4[")
            + field("PAGE", None) + wml.run("] f5[") + simple_field("PAGE", "9", rpr=BOLD) + wml.run("] f6[")
            + field("NUMPAGES", "9", rpr={}, instr_rpr={"color": "FF0000"}, result_rpr={"color": "0000FF"})
            + wml.run("]"))
    return wml.paragraph(runs, mark={}, pStyle="Header")


#: Fields Word draws as cached (``computed``).
OTHERS = (
    ("AUTHOR", "Cached Author"), ("FILENAME", "cached-name.docx"), ("TITLE", "Cached Title"),
    ("NUMWORDS", "12345"), ("= 2 + 3", "99"), ('HYPERLINK "http://example.com/"', "Link text"),
    ("DOCPROPERTY Company", "Cached Co"), ('QUOTE "quoted"', "cachedquote"),
)
#: Fields Word computes again on export (``word``).
RECOMPUTED = (
    ('DATE \\@ "d MMMM yyyy"', "1 January 2001"), ("TIME", "01:02"), ("REF bm1 \\h", "cached ref"),
    ("PAGEREF bm2 \\h", "77"), ("SEQ Figure \\* ARABIC", "42"), ('IF 1 = 1 "yes" "no"', "maybe"),
)


def _body_fields(fields=OTHERS, nested_if: bool = False) -> str:
    out = wml.paragraph('<w:bookmarkStart w:id="1" w:name="bm1"/>' + wml.run("Bookmarked text")
                        + '<w:bookmarkEnd w:id="1"/>', mark={})
    for k, (instruction, cached) in enumerate(fields):
        out += wml.paragraph(wml.run(f"o{k}[") + field(instruction, cached) + wml.run("]"), mark={})
    if nested_if:
        nested = (wml.run("n[") + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                  + '<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>'
                  + field("PAGE", "1")
                  + '<w:r><w:instrText xml:space="preserve"> = 1 "one" "other" </w:instrText></w:r>'
                  + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + wml.run("neither")
                  + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + wml.run("]"))
        out += wml.paragraph(nested, mark={})
    toc_begin = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    out += wml.paragraph(wml.run("toc[") + toc_begin + wml.run("Cached contents line one"), mark={})
    out += wml.paragraph(wml.run("Cached contents line two") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                         + wml.run("]"), mark={})
    long_instruction = 'HYPERLINK "http://example.com/' + "segment/" * 38 + '"'
    out += wml.paragraph(wml.run("long[") + field(long_instruction, "short result") + wml.run("] after"), mark={})
    out += wml.paragraph(wml.run("Second page of the fields ") + '<w:bookmarkStart w:id="2" w:name="bm2"/>'
                         + wml.run("marked") + '<w:bookmarkEnd w:id="2"/>', mark={}, pageBreakBefore=True)
    return out


#: The filler before a wide ``PAGE`` / ``NUMPAGES`` in the section numbered from
#: 1,000,000: word counts chosen by the model so the cached ``1`` fits on the first line
#: and the computed ``1000000`` does not, and neighbours of them.
WIDE_WORDS = tuple(range(14, 20))


def _wide(number: int) -> str:
    out = ""
    for words in WIDE_WORDS:
        for instruction in ("PAGE", "NUMPAGES"):
            filler = " ".join(["measured"] * words) + " " + "w" * (words % 5)
            out += wml.paragraph(wml.run(f"W{words}{instruction[0]} {filler} ") + field(instruction, "1")
                                 + wml.run(" tail"), mark={})
    return out


DOCUMENTS = ("computed", "word")


@functools.lru_cache(maxsize=None)
def parts(document: str = "computed") -> tuple[str, str, story_docx.Parts]:
    stories = story_docx.Parts()
    header_id = stories.add("header", header())
    body = ""
    final = ""
    word = document == "word"
    formats = (({}, "decimal"),) + LANGUAGE_FORMATS if word else FORMATS
    for number, (page_numbers, name) in enumerate(formats):
        # Numbered from 1,000,000, Word draws no footer paragraph of ``\\* Arabic``,
        # ``ArabicDash``, ``Hex``, ``MERGEFORMAT`` and ``\\# "000"`` results: that section's
        # footer is the first paragraph alone.
        short = name in ("upperRoman 3999", "upperRoman 4000")
        footer_xml = footer(number, short, bare=name == "wide", switches=LANGUAGE_SWITCHES if word else None)
        references = stories.reference("footer", "default", stories.add("footer", footer_xml))
        if number == 0:
            references = stories.reference("header", "default", header_id) + references
        sect = story_docx.section(references=references, page_numbers=page_numbers or None)
        content = (_body_fields(RECOMPUTED, nested_if=True) if word else _body_fields()) if number == 0 else ""
        if name == "wide":
            content += _wide(number)
        content += wml.paragraph(wml.run(f"Section {number} ({name}) first page body ") + field("PAGE", "0"),
                                 mark={}, pageBreakBefore=number == 0)
        last = number == len(formats) - 1
        closing = wml.run(f"Section {number} second page body ") + field("PAGE", "0")
        if last:
            body += content + wml.paragraph(closing, mark={}, pageBreakBefore=True)
            final = sect
        else:
            body += content + wml.paragraph(closing, mark={}, pageBreakBefore=True, sect=sect)
    return body, final, stories


def build(document: str = "computed") -> bytes:
    body, final, stories = parts(document)
    return story_docx.package(body, final, stories, compatibility_mode=15)
