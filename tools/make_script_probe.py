#!/usr/bin/env python3
"""What superscripts, subscripts, ``w:position`` and run borders do to a line.

``filesamples/sample1`` (ROADMAP.md, "Nine more real documents", finding 5) has a line
holding a ``w:vertAlign`` superscript and subscript that is ~3 px taller than its text,
and a line holding a run with ``w:bdr`` that is ~2 px taller; and Word draws the
``vertAlign`` glyphs at 2/3 size, which the cascade ignores.  One document cannot say
what the rule is, so this probe isolates each variable:

* ``sup`` / ``sub`` / ``supsub``: a ``w:vertAlign`` run, alone or both, in the text's
  face and size, or at twice and half the text's size (``extra_hp``) so it is taller or
  shorter than the rest of the line;
* ``pos``: ``w:position`` (half points, raised when positive), swept;
* ``bdr``: a ``w:bdr`` run, ``w:sz`` and ``w:space`` swept, in the text's size and at
  other sizes (a border on a run that does not reach the line's extremes);
* ``plain``: nothing, the control;

in Calibri, Aptos, Times New Roman, Arial, Cambria and Courier New (different ascent /
descent ratios), at several sizes, under ``auto`` single and 1.15, ``exact`` and
``atLeast``, in three compatibility settings.  One group per page, ``LINES`` one-line
paragraphs each, so the pitch is pinned by the first and last baselines; the mark is in
the text's face and size.  Every paragraph also ends with the extra run's text in the
text's own size, so the drawn advance of the two gives the extra run's size exactly.

Scored through ``baselines.predict`` by ``read_script_probe.py`` (``probe_documents``):
the probe tests the model's code.  A tenth document, ``script-sizes``, is one line per
``w:sz`` from 2 to 96 half points in twelve faces (:data:`SIZE_FACES`), each holding a
superscript and a subscript followed by text objects of their own, so their advances
give the size Word draws them at, and their baselines the offset; Word paginates it.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

LINES = 12
SETTINGS = {"none": None, "12": 12, "15": 15}
SCRIPT = "Hxh"


def fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


@dataclass(frozen=True)
class Group:
    face: str
    half_points: int
    #: ``plain``, ``sup``, ``sub``, ``supsub``, ``pos``, ``bdr``.
    kind: str
    rule: str = "auto"
    line: int = 240
    #: The extra run's own size (``None``: the text's).
    extra_hp: int | None = None
    #: ``pos``: ``w:position`` in half points; ``bdr``: ``(w:sz, w:space)``.
    param: object = None

    @property
    def name(self) -> str:
        extra = "" if self.extra_hp is None else f"@{self.extra_hp}"
        param = "" if self.param is None else (
            f"{self.param:+d}" if isinstance(self.param, int) else "x".join(map(str, self.param)))
        return (f"{self.kind}{param}{extra}/{self.face.replace(' ', '')}{self.half_points}"
                f"/{self.rule}{self.line}")

    def extra_runs(self) -> str:
        hp = self.extra_hp or self.half_points
        rpr = {"rFonts": fonts(self.face), "sz": hp, "szCs": hp}
        if self.kind in ("sup", "sub"):
            return wml.run(SCRIPT, **rpr, vertAlign="superscript" if self.kind == "sup" else "subscript")
        if self.kind == "supsub":
            return (wml.run(SCRIPT, **rpr, vertAlign="superscript")
                    + wml.run(SCRIPT, **rpr, vertAlign="subscript"))
        if self.kind == "pos":
            return wml.run(SCRIPT, **rpr, position=self.param)
        if self.kind == "bdr":
            sz, space = self.param
            return wml.run(SCRIPT, **rpr, bdr={"val": "single", "sz": sz, "space": space, "color": "auto"})
        return ""


@dataclass(frozen=True)
class Probe:
    name: str
    setting: str
    groups: tuple[Group, ...] = ()
    #: ``(face, half points)`` of the size sweep's paragraphs, after the groups.
    sizes: tuple[tuple[str, int], ...] = ()


def paragraphs(group: Group) -> list[str]:
    text = {"rFonts": fonts(group.face), "sz": group.half_points, "szCs": group.half_points}
    out = []
    for index in range(LINES):
        runs = wml.run(f"Hxample {index} ", **text) + group.extra_runs()
        if group.kind != "plain":
            runs += wml.run(f" {SCRIPT}", **text)
        props = {
            "pageBreakBefore": index == 0,
            "spacing": {"before": 0, "after": 0, "line": group.line, "lineRule": group.rule},
        }
        out.append(wml.paragraph(runs, mark=text, **props))
    return out


def size_paragraph(face: str, half_points: int) -> str:
    """One line: ``Hx``, a superscript, ``Hx``, a subscript, ``Hx``, all at one ``w:sz``.
    Each script run is followed by a text object of its own, whose exact pen x gives the
    script run's advance, and so its size."""
    text = {"rFonts": fonts(face), "sz": half_points, "szCs": half_points}
    runs = (wml.run("Hx ", **text) + wml.run(SCRIPT, **text, vertAlign="superscript")
            + wml.run(" Hx ", **text) + wml.run(SCRIPT, **text, vertAlign="subscript")
            + wml.run(" Hx", **text))
    return wml.paragraph(runs, mark=text, spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})


def kinds(probe: Probe) -> list[str]:
    return ([group.name for group in probe.groups for _ in range(LINES)]
            + [f"size/{face.replace(' ', '')}{hp}" for face, hp in probe.sizes])


def build(probe: Probe) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    extra = []
    if SETTINGS[probe.setting] is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=SETTINGS[probe.setting]))
    body = ("".join(p for group in probe.groups for p in paragraphs(group))
            + "".join(size_paragraph(face, hp) for face, hp in probe.sizes))
    return probe_docx.package(body, styles=styles, extra_parts=tuple(extra))


# -- the sweeps ----------------------------------------------------------------------------

FACES = ("Calibri", "Aptos", "Times New Roman", "Arial", "Cambria", "Courier New")

RULES = (("auto", 240), ("auto", 276), ("exact", 360), ("atLeast", 300))

#: ``w:vertAlign``: in the text's size, and in runs twice and half its size (a script
#: taller or shorter than the rest of the line), under every rule; the controls beside.
SCRIPT_GROUPS = tuple(
    Group(face, hp, kind, rule, line, extra_hp)
    for face in FACES for hp in (22, 40) for rule, line in RULES
    for kind, extra_hp in (("plain", None), ("sup", None), ("sub", None), ("supsub", None),
                           ("sup", 2 * hp), ("sub", 2 * hp), ("sup", hp // 2), ("sub", hp // 2))
)

POSITIONS = (1, 2, 3, 6, 12, 24, 48, -1, -2, -3, -6, -12, -24, -48)

#: ``w:position`` swept under every rule; and on runs half and twice the text's size.
POSITION_GROUPS = (
    tuple(Group(face, 22, "pos", rule, line, param=pos)
          for face in FACES for rule, line in RULES for pos in POSITIONS)
    + tuple(Group(face, 22, "pos", extra_hp=extra_hp, param=pos)
            for face in FACES for extra_hp, positions in ((11, (6, 12, 24, -6, -12, -24)), (44, (6, -6)))
            for pos in positions)
)

BORDERS = ((2, 0), (4, 0), (6, 0), (8, 0), (12, 0), (18, 0), (24, 0), (48, 0), (96, 0), (27, 0),
           (4, 1), (4, 2), (4, 4), (4, 8), (4, 31), (24, 12))

#: ``w:bdr``: ``w:sz`` and ``w:space`` swept under every rule; on runs half and twice
#: the text's size (a border on a run that does not reach the line's extremes).
BORDER_GROUPS = (
    tuple(Group(face, 22, "bdr", rule, line, param=border)
          for face in FACES for rule, line in RULES for border in BORDERS)
    + tuple(Group(face, 22, "bdr", extra_hp=extra_hp, param=border)
            for face in FACES for extra_hp in (11, 44) for border in ((4, 0), (24, 0), (4, 4), (24, 12)))
)

#: Faces for the size sweep: the six above, and faces whose OS/2 script fields are far
#: from 0.65 em -- Helvetica Neue 0.204, Baskerville Old Face 0.416,
#: Impact 0.528, Galvji 1.228 -- or whose subscript offset is negative (Charter).
SIZE_FACES = FACES + ("Georgia", "Helvetica Neue", "Baskerville Old Face", "Impact", "Galvji",
                      "Charter")

PROBES = (
    *(Probe(f"script-{setting}", setting, SCRIPT_GROUPS) for setting in SETTINGS),
    *(Probe(f"position-{setting}", setting, POSITION_GROUPS) for setting in SETTINGS),
    *(Probe(f"border-{setting}", setting, BORDER_GROUPS) for setting in SETTINGS),
    Probe("script-sizes", "none", sizes=tuple((face, hp) for face in SIZE_FACES for hp in range(2, 97))),
)
