#!/usr/bin/env python3
"""Phase 4's "done when": multi-page documents whose every paragraph must land on Word's page.

One body of text -- 260 paragraphs drawn from a seeded generator -- set on **three page
sizes** (Letter, A4, A5) with **three margin sets** (1 inch all round; 0.5 inch; and an
asymmetric set, 1800 top, 1080 bottom, 2160 left, 1440 right), nine documents in
compatibility mode 15; the Letter / 1 inch document again with no ``settings.xml`` and in
modes 12 and 14 (where ``keepNext`` and borders behave otherwise); and a document of four
sections that change size and margins at ``nextPage``, ``oddPage`` and ``continuous``
breaks.

What the body holds, so that every rule of the paginator is exercised by text it did not
choose: paragraphs of 1 to 140 words in five faces (Calibri, Times New Roman, Arial,
Cambria, Georgia) at six sizes, under five line rules (``auto`` 240, 276 and 360, ``exact``
300 twips, ``atLeast`` 320), with space before and after, left-aligned;
headings with ``keepNext`` (so chains of two where two headings meet); ``keepLines``
paragraphs; paragraphs with widow control off; bottom borders; numbered list items;
empty paragraphs; ``pageBreakBefore`` headings and manual page breaks.

Which page each line of each paragraph is on is read from Word's export
(``read_pagination_probe.py``); every page break is the model's.
"""

from __future__ import annotations

import functools
import random
from dataclasses import dataclass

import probe_docx
import wml

SIZES = {"letter": (12240, 15840), "a4": (11906, 16838), "a5": (8391, 11906)}
#: name -> (top, right, bottom, left)
MARGINS = {"normal": (1440, 1440, 1440, 1440), "narrow": (720, 720, 720, 720),
           "wide": (1800, 1440, 1080, 2160)}
FACES = ("Calibri", "Times New Roman", "Arial", "Cambria", "Georgia")
SIZES_HP = (18, 20, 21, 22, 24, 26)
RULES = ((240, "auto"), (276, "auto"), (360, "auto"), (300, "exact"), (320, "atLeast"))

WORDS = """the of and to in is that it for as with was on be by this are or from at an which
not have has but they their one all were can more will been would there when what its also
into about than other some these may only time over such new like any most after first
between where many those through each well before while should because both under same
page line word break paragraph measure model probe column margin section height width
layout rule space border heading list table figure number letter document printer
oracle baseline pitch ascent descent glyph advance kerning budget greedy inclusive exact
compatibility mode setting widow orphan keep chain footnote header footer device pixel
twelve fourteen fifteen sixteen seventy hundred thousand million quarter half double
""".split()


def face(name: str) -> dict:
    return {"ascii": name, "hAnsi": name, "eastAsia": name, "cs": name}


@dataclass(frozen=True)
class Document:
    name: str
    size: str
    margins: str
    mode: int | None = 15
    sections: bool = False


DOCUMENTS = tuple(
    [Document(f"{size}-{margins}", size, margins) for size in SIZES for margins in MARGINS]
    + [Document(f"letter-normal-{mode or 'none'}", "letter", "normal", mode) for mode in (None, 12, 14)]
    + [Document("sections", "letter", "normal", 15, sections=True)]
)


def _sentence(rng: random.Random, words: int) -> str:
    out = []
    for k in range(words):
        word = rng.choice(WORDS)
        if k == 0:
            word = word.capitalize()
        if rng.random() < 0.04:
            word = f"{rng.randint(2, 999)}"
        if rng.random() < 0.02:
            word += "-" + rng.choice(WORDS)
        if k < words - 1 and rng.random() < 0.08:
            word += ","
        out.append(word)
    return " ".join(out) + "."


def _run(text: str, props: dict) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{text}</w:t></w:r>'


@functools.lru_cache(maxsize=None)
def body(seed: int = 4, count: int = 260) -> tuple[str, ...]:
    """The paragraphs, as ``w:p`` XML."""
    rng = random.Random(seed)
    out = []
    while len(out) < count:
        kind = rng.choices(("body", "heading", "keeplines", "nowidow", "border", "list", "empty",
                            "pagebreak", "manual"),
                           weights=(58, 12, 5, 5, 4, 8, 4, 2, 2))[0]
        name = rng.choice(FACES)
        size = rng.choice(SIZES_HP)
        run = {"rFonts": face(name), "sz": size, "szCs": size}
        line, rule = rng.choice(RULES)
        spacing = {"before": rng.choice((0, 0, 60, 120)), "after": rng.choice((0, 120, 160, 200)),
                   "line": line, "lineRule": rule}
        props: dict = {"spacing": spacing}
        # Justified text was drawn here once: in mode 15 Word fits more words on a
        # justified line than on a left-aligned one (a line-breaking question, Phase 3's
        # "justification", not measured), so the draw is kept and its result unused.
        rng.random()
        text = _sentence(rng, rng.randint(1, 140))
        if kind in ("heading", "pagebreak"):
            run = {"rFonts": face(name), "b": True, "sz": rng.choice((28, 32)), "szCs": 28}
            text = _sentence(rng, rng.randint(2, 8))[:-1]
            props = {"keepNext": True, "spacing": {"before": 240, "after": 60, "line": 240, "lineRule": "auto"}}
            if kind == "pagebreak":
                props["pageBreakBefore"] = True
        elif kind == "keeplines":
            props["keepLines"] = True
            text = _sentence(rng, rng.randint(40, 140))
        elif kind == "nowidow":
            props["widowControl"] = False
            text = _sentence(rng, rng.randint(40, 140))
        elif kind == "border":
            props["pBdr"] = '<w:bottom w:val="single" w:sz="6" w:space="2" w:color="000000"/>'
        elif kind == "list":
            props = {"numPr": '<w:ilvl w:val="0"/><w:numId w:val="1"/>',
                     "spacing": {"before": 0, "after": 60, "line": 240, "lineRule": "auto"}}
            run = {"rFonts": face("Calibri"), "sz": 22, "szCs": 22}
            text = _sentence(rng, rng.randint(3, 40))
        elif kind == "empty":
            out.append(wml.paragraph("", mark=run, **props))
            continue
        elif kind == "manual":
            words = text.split(" ")
            cut = rng.randint(0, len(words))
            out.append(wml.paragraph(
                _run(" ".join(words[:cut]) + (" " if cut else ""), run)
                + f'<w:r>{wml.rpr(**run)}<w:br w:type="page"/></w:r>'
                + _run(" ".join(words[cut:]), run), mark=run, **props))
            continue
        out.append(wml.paragraph(_run(text, run), mark=run, **props))
    return tuple(out)


def _section(size: str, margins: str, kind: str | None = None) -> str:
    width, height = SIZES[size]
    top, right, bottom, left = MARGINS[margins]
    return (
        "<w:sectPr>"
        + (f'<w:type w:val="{kind}"/>' if kind else "")
        + f'<w:pgSz w:w="{width}" w:h="{height}"/>'
        f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    )


#: The sections document: (size, margins, how the section starts), and how many of the
#: body's paragraphs each takes.
SECTIONS = (("letter", "normal", None, 60), ("a5", "narrow", "nextPage", 60),
            ("a4", "wide", "oddPage", 60), ("a4", "normal", "continuous", 80))


def build(document: Document) -> bytes:
    paragraphs = list(body())
    if document.sections:
        xml, index = "", 0
        for number, (size, margins, kind, count) in enumerate(SECTIONS):
            chunk = paragraphs[index:index + count]
            index += count
            if number < len(SECTIONS) - 1:
                # A section's properties go in its last paragraph, and its w:type says
                # how that section itself starts.
                chunk[-1] = _with_section(chunk[-1], _section(size, margins, kind))
            xml += "".join(chunk)
        final = _section(*SECTIONS[-1][:3])
    else:
        xml = "".join(paragraphs)
        final = _section(document.size, document.margins)
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": face("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    numbering = wml.numbering_part([("decimal", "%1.", {"ind": {"left": 720, "hanging": 360}},
                                     {"rFonts": face("Calibri")})])
    extra = [numbering]
    if document.mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=document.mode))
    return probe_docx.package(xml, final_section=final, styles=styles, extra_parts=tuple(extra))


def _with_section(paragraph: str, section: str) -> str:
    """``paragraph`` with ``section`` as the last child of its ``w:pPr``."""
    if "<w:pPr>" in paragraph:
        head, rest = paragraph.split("<w:pPr>", 1)
        inner, tail = rest.split("</w:pPr>", 1)
        # w:sectPr follows w:rPr inside w:pPr.
        return f"{head}<w:pPr>{inner}{section}</w:pPr>{tail}"
    return paragraph.replace("<w:p>", f"<w:p><w:pPr>{section}</w:pPr>", 1)
