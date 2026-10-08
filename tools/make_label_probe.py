#!/usr/bin/env python3
"""What a list label adds to the height of its line.

``make_mixed_line_probe.py`` put a SymbolMT bullet in front of Cambria and found the
line takes the largest ascent and the largest descent over everything on it, the label
included.  Cambria's descent (455/2048 em) exceeds SymbolMT's (450), so that probe could
not see what ``filesamples/sample1`` showed: bullets in SymbolMT, Courier New and
Wingdings under Ubuntu (descent 0.189 em) make lines ~1 px shorter than a label's
descent would, as if **the label's ascent counts and its descent does not** (ROADMAP.md,
"Nine more real documents", finding 4).  This probe crosses:

* text faces with small to large descents -- Tahoma (0.207 em), Arial (0.212), Cambria
  (0.222), Calibri (0.269) -- at 11 pt;
* labels whose descent is smaller or larger than the text's, and whose ascent is too:
  SymbolMT and Wingdings bullets (0.220, 0.211), Courier New ``o`` (0.300, ascent only
  0.833), Arial Black ``•`` (0.310, ascent 1.101), all at 11 pt; and Courier New and
  SymbolMT at 20 pt -- **a label taller than the text** above and below;
* controls: the Courier New ``o`` at 11 and 20 pt as inline *runs* of the text, whose
  descent is expected to count like any text's (not SymbolMT's bullet: the PDF gives
  its private-use code point back as U+2022, so its line cannot be matched by text);
* single spacing and ``auto`` 276;

one group per page, 22 one-line paragraphs each (the pitch pinned to 1/21 px), the mark
in the text's face and size.  In three settings: no ``settings.xml`` (as the mixed-line
probe), mode 12 (``sample1``'s) and mode 15.  Measured by ``read_label_probe.py``
through ``baselines.predict``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

LINES = 22
TEXT_FACES = ("Tahoma", "Arial", "Cambria", "Calibri")
#: name -> (face, half points, lvlText).  The level index is the position here.
LABELS = {
    "symbol": ("Symbol", 22, ""),
    "wingdings": ("Wingdings", 22, ""),
    "courier": ("Courier New", 22, "o"),
    "arialblack": ("Arial Black", 22, "•"),
    "courier20": ("Courier New", 40, "o"),
    "symbol20": ("Symbol", 40, ""),
}
#: Inline controls: the same glyphs as runs of the text.
INLINE = {"inline-courier": "courier", "inline-courier20": "courier20"}
RULES = (("auto", 240), ("auto", 276))
SETTINGS = {"none": None, "12": 12, "15": 15}


def _fonts(face: str) -> dict:
    if face in ("Symbol", "Wingdings"):
        return {"ascii": face, "hAnsi": face, "hint": "default"}
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


@dataclass(frozen=True)
class Group:
    text: str
    extra: str
    rule: str
    line: int

    @property
    def name(self) -> str:
        return f"{self.extra}/{self.text.replace(' ', '')}/{self.rule}{self.line}"


GROUPS = tuple(Group(text, extra, rule, line)
               for rule, line in RULES for text in TEXT_FACES for extra in (*LABELS, *INLINE))


@dataclass(frozen=True)
class Probe:
    setting: str

    @property
    def name(self) -> str:
        return f"label-{self.setting}"


PROBES = tuple(Probe(setting) for setting in SETTINGS)


def blocks() -> list[tuple[str, str]]:
    out = []
    levels = list(LABELS)
    for group in GROUPS:
        text_rpr = {"rFonts": _fonts(group.text), "sz": 22, "szCs": 22}
        for index in range(LINES):
            runs = wml.run(f"Hxample {index}", **text_rpr)
            props = {
                "pageBreakBefore": index == 0,
                "spacing": {"before": 0, "after": 0, "line": group.line, "lineRule": group.rule},
            }
            if group.extra in LABELS:
                level = levels.index(group.extra)
                props["numPr"] = f'<w:ilvl w:val="{level}"/><w:numId w:val="1"/>'
            else:
                face, half_points, char = LABELS[INLINE[group.extra]]
                runs += wml.run(char, rFonts=_fonts(face), sz=half_points, szCs=half_points)
            out.append((group.name, wml.paragraph(runs, mark=text_rpr, **props)))
    return out


def kinds(probe: Probe) -> list[str]:
    return [kind for kind, _ in blocks()]


def build(probe: Probe) -> bytes:
    numbering = wml.numbering_part([
        ("bullet", char, {"ind": {"left": 720, "hanging": 360}},
         {"rFonts": _fonts(face), "sz": half_points, "szCs": half_points})
        for face, half_points, char in LABELS.values()
    ])
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": _fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    extra = [numbering]
    if SETTINGS[probe.setting] is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=SETTINGS[probe.setting]))
    return probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=tuple(extra))
