"""Which face, size and weight a character is measured in.

A run does not have *a* face.  ``w:rFonts`` has four slots and each character picks one:

* **complex script** -- if the run is ``w:cs`` or ``w:rtl``: the ``cs`` slot, with
  ``w:szCs``, ``w:bCs`` and ``w:iCs``.  **Word 16.106 decides this from those two
  properties, not from the character**: Hebrew and Arabic in a run marked neither are
  drawn at ``w:sz``, emboldened by ``w:b``, in the non-complex slot's face (which here
  lacks the glyphs, so Word falls back to Times New Roman) -- cases ``f09``, ``f13``,
  ``f24``.  ECMA-376 17.3.2.26 classifies by Unicode range.  Only Hebrew and Arabic in a
  left-to-right paragraph have been drawn.
* otherwise **by Unicode range** (ECMA-376 17.3.2.26's table, abridged): ASCII takes
  ``ascii``; East Asian ideographs, kana, Hangul and full-width forms take ``eastAsia``;
  everything else takes ``hAnsi`` -- except that with ``w:hint="eastAsia"`` the shared
  punctuation and symbol ranges take ``eastAsia`` too (measured on U+201C/D, U+00A7,
  U+2014, and U+00E9 with an East Asian language of ``zh-CN``: cases ``f11``, ``f12``).

A slot then names a face either directly or through the theme (``w:asciiTheme``...):

* ``major*`` / ``minor*`` Latin references take the theme's ``a:latin`` typeface;
* East Asian and complex-script references go through ``w:themeFontLang`` in
  ``settings.xml``: the theme's per-script font for that language's script (``Hans`` for
  ``zh-CN``, ``Hebr`` for ``he-IL``), else its ``a:ea`` / ``a:cs`` typeface.  **Without
  ``w:themeFontLang`` Word does not use the theme for these slots at all** -- it drew
  Times New Roman for a ``minorBidi`` whose theme said Courier New, and MS Mincho for a
  ``minorEastAsia`` (cases ``f15``, ``f18``, ``f23`` across three settings variants).
  That fallback is Word's, not the document's, and is reported as ``None``.

A package with no theme part resolves Latin theme references against a built-in theme
whose minor font is Calibri and major font Calibri Light (``style-no-theme``,
``style-no-styles``).  ROADMAP.md records that Phase 0's ``hello.docx`` once exported an
Aptos paragraph mark from the same bytes; re-exported now, it draws Calibri.  So what an
unstyled mark resolves to is Word's state, and callers get :data:`APPLICATION` origins
for it rather than a silent guess.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from fractions import Fraction

from ..model import Document, FontScheme, ThemeFonts
from ..parse.properties import FontRef
from .cascade import APPLICATION, THEME, Origin, Resolved

#: Used when the package has no theme part.  Latin only: measured n02/n03.
BUILTIN_FONT_SCHEME = FontScheme(
    major=ThemeFonts(latin="Calibri Light"),
    minor=ThemeFonts(latin="Calibri"),
)

#: ``w:sz`` when nothing in the hierarchy declares one: 12 pt, drawn at 50 device px in
#: every probe that states no size (with no styles part, and with one lacking
#: ``w:docDefaults``).  ECMA-376 says 10 pt.  12 pt is also the size in this machine's
#: ``Normal.dotm``, and which of the two Word is reading is not settled.
APPLICATION_HALF_POINTS = 24

#: Language (primary subtag, or full tag where the region decides) -> ISO 15924 script,
#: the key of a theme's ``a:font script=...`` list.  Hebr and Hans are measured; the rest
#: follow the tags Office's own themes use.
_SCRIPT_OF_LANGUAGE = {
    "he": "Hebr", "yi": "Hebr", "ar": "Arab", "fa": "Arab", "ur": "Arab", "ps": "Arab",
    "sd": "Arab", "ug": "Arab", "syr": "Syrc", "dv": "Thaa", "ja": "Jpan", "ko": "Hang",
    "zh-cn": "Hans", "zh-sg": "Hans", "zh-tw": "Hant", "zh-hk": "Hant", "zh-mo": "Hant",
    "zh": "Hans", "th": "Thai", "hi": "Deva", "mr": "Deva", "ne": "Deva", "sa": "Deva",
    "bn": "Beng", "as": "Beng", "ta": "Taml", "te": "Telu", "kn": "Knda", "ml": "Mlym",
    "gu": "Gujr", "pa": "Guru", "or": "Orya", "km": "Khmr", "lo": "Laoo", "my": "Mymr",
    "si": "Sinh", "bo": "Tibt", "am": "Ethi", "ti": "Ethi", "iu": "Cans", "chr": "Cher",
    "mn": "Mong", "ka": "Geor", "hy": "Armn", "vi": "Viet",
}


def script_of_language(tag: str | None) -> str | None:
    if not tag:
        return None
    tag = tag.lower()
    return _SCRIPT_OF_LANGUAGE.get(tag) or _SCRIPT_OF_LANGUAGE.get(tag.split("-")[0])


# -- which slot ------------------------------------------------------------------------

#: Ranges that are East Asian whatever the hint.
_EAST_ASIAN_RANGES = (
    (0x1100, 0x11FF), (0x2E80, 0x2FDF), (0x2FF0, 0x31FF), (0x3200, 0x4DBF),
    (0x4E00, 0x9FFF), (0xA000, 0xA4CF), (0xAC00, 0xD7AF), (0xF900, 0xFAFF),
    (0xFE30, 0xFE4F), (0xFF00, 0xFFEF), (0x20000, 0x2FFFF),
)
#: Latin-1 characters that take the East Asian slot under ``w:hint="eastAsia"``
#: (ECMA-376 17.3.2.26), and those that do so only for a Chinese East Asian language.
_LATIN1_HINTED = frozenset(
    [0xA1, 0xA4, 0xA7, 0xA8, 0xAA, 0xAD, 0xAF, 0xD7, 0xF7]
    + list(range(0xB0, 0xB5)) + list(range(0xB6, 0xBB)) + list(range(0xBC, 0xC0))
)
_LATIN1_HINTED_CHINESE = frozenset(
    [0xE0, 0xE1, 0xE8, 0xE9, 0xEA, 0xEC, 0xED, 0xF2, 0xF3, 0xF9, 0xFA, 0xFC]
)


def slot_for(char: str, *, complex_script: bool = False, hint: str | None = None,
             east_asian_language: str | None = None) -> str:
    """``ascii`` / ``hAnsi`` / ``eastAsia`` / ``cs`` for one character of a run."""
    if complex_script:
        return "cs"
    code = ord(char)
    if code < 0x80:
        return "ascii"
    if any(low <= code <= high for low, high in _EAST_ASIAN_RANGES):
        return "eastAsia"
    if hint == "eastAsia":
        if 0xA0 <= code <= 0xFF:
            if code in _LATIN1_HINTED:
                return "eastAsia"
            chinese = (east_asian_language or "").lower().startswith("zh")
            return "eastAsia" if chinese and code in _LATIN1_HINTED_CHINESE else "hAnsi"
        return "eastAsia"
    return "hAnsi"


# -- which face ------------------------------------------------------------------------


def theme_face(reference: str, document: Document) -> tuple[str | None, str]:
    """Resolve an ``ST_Theme`` value; returns ``(face or None, explanation)``."""
    scheme = document.font_scheme or BUILTIN_FONT_SCHEME
    source = "theme" if document.font_scheme else "built-in theme"
    fonts = scheme.major if reference.startswith("major") else scheme.minor
    part = "major" if reference.startswith("major") else "minor"
    if reference.endswith(("Ascii", "HAnsi")):
        return (fonts.latin or None), f"{source} {part} latin"
    language_key = "eastAsia" if reference.endswith("EastAsia") else "bidi"
    language = document.theme_font_lang.get(language_key)
    if not language:
        return None, f"{reference} without w:themeFontLang/@{language_key}: Word's own fallback"
    script = script_of_language(language)
    if script and fonts.scripts.get(script):
        return fonts.scripts[script], f"{source} {part} {script} (themeFontLang {language})"
    face = fonts.east_asian if language_key == "eastAsia" else fonts.complex_script
    slot = "ea" if language_key == "eastAsia" else "cs"
    return (face or None), f"{source} {part} {slot} (no {script} font for {language})"


def slot_face(resolved: Resolved, slot: str, document: Document) -> tuple[str | None, tuple[Origin, ...]]:
    """The face a slot names, and the origins that led to it (theme step last)."""
    key = f"rFonts.{slot}"
    reference: FontRef | None = resolved.get(key)
    chain = tuple(resolved.origins.get(key, ()))
    if reference is None:
        # Nothing names this slot anywhere: Word uses the minor theme font (measured,
        # style-theme-no-rfonts: Georgia from the theme; Calibri from the built-in one).
        default = {"ascii": "minorHAnsi", "hAnsi": "minorHAnsi",
                   "eastAsia": "minorEastAsia", "cs": "minorBidi"}[slot]
        reference = FontRef(theme=default)
        chain = (Origin(APPLICATION, detail=f"no rFonts: {default}"),)
    if reference.theme:
        face, detail = theme_face(reference.theme, document)
        return face, chain + (Origin(THEME, detail=detail),)
    return reference.face, chain


@dataclass(frozen=True)
class CharacterFormat:
    """What one character is measured in."""

    slot: str
    face: str | None
    half_points: int
    bold: bool
    italic: bool
    origins: dict
    #: ``w:vertAlign`` (``superscript`` / ``subscript``), or ``None``.  ``half_points``
    #: stays the run's ``w:sz``: that is the size the run takes part in its line's height
    #: at; the size its glyphs are drawn at is :func:`script_half_points`.
    vertical_align: str | None = None
    #: ``w:position``: how far the run is raised (lowered when negative), in half points.
    position: int = 0
    #: ``w:bdr`` as ``(width in eighths of a point, space in points)``; ``None`` when the
    #: run has no border (none stated, or ``none`` / ``nil``).
    border: tuple[int, int] | None = None

    def explain(self) -> str:
        return "\n".join(f"{key}: " + "; ".join(map(str, chain)) for key, chain in self.origins.items())


def character_format(resolved: Resolved, char: str, document: Document) -> CharacterFormat:
    """Face, size, bold and italic for ``char`` in a run whose properties are ``resolved``."""
    complex_script = bool(resolved.get("cs") or resolved.get("rtl"))
    slot = slot_for(char, complex_script=complex_script, hint=resolved.get("rFonts.hint"),
                    east_asian_language=resolved.get("lang.eastAsia"))
    face, face_origins = slot_face(resolved, slot, document)
    size_key, bold_key, italic_key = ("szCs", "bCs", "iCs") if complex_script else ("sz", "b", "i")
    if size_key in resolved:
        half_points = resolved[size_key]
        size_origin = resolved.origins[size_key]
    else:
        half_points = APPLICATION_HALF_POINTS
        size_origin = (Origin(APPLICATION, detail="no size declared anywhere"),)
    vertical_align = resolved.get("vertAlign")
    border = resolved.get("bdr")
    return CharacterFormat(
        slot=slot,
        face=face,
        half_points=half_points,
        bold=bool(resolved.get(bold_key)),
        italic=bool(resolved.get(italic_key)),
        origins={
            "face": face_origins,
            size_key: size_origin,
            bold_key: resolved.origins.get(bold_key, ()),
            italic_key: resolved.origins.get(italic_key, ()),
        },
        vertical_align=vertical_align if vertical_align in ("superscript", "subscript") else None,
        position=resolved.get("position", 0) or 0,
        border=(border[1], border[2]) if border and border[0] not in (None, "none", "nil") else None,
    )


# -- superscript and subscript ------------------------------------------------------------

#: The fraction of the size a ``w:vertAlign`` run is drawn at when the face's own value
#: (``OS/2.ySuperscriptYSize`` / ``ySubscriptYSize`` over ``unitsPerEm``) is out of
#: :data:`SCRIPT_SIZE_RANGE`: Helvetica Neue (0.204) and Galvji (1.228) are drawn at
#: 3/5 at every size from 2 to 96 half points (the feasible ratio is 0.5990-0.6011).
SCRIPT_SIZE_FALLBACK = Fraction(3, 5)

#: Where Word takes a face's own script size.  Measured inside: 0.416 (Baskerville Old
#: Face), 0.528 (Impact), 0.600 (Aptos, Charter), 0.650 (Calibri, Times New Roman, Arial,
#: Cambria, Courier New, Georgia, and sample1's embedded Ubuntu); outside: 0.204 and
#: 1.228.  **The bounds themselves are not measured**: the lower lies in (0.204, 0.416],
#: the upper in [0.650, 1.228).  1/4 and 1 are placeholders inside those intervals.
SCRIPT_SIZE_RANGE = (Fraction(1, 4), Fraction(1))


#: The least ``OS/2.ySuperscriptYOffset`` (of an em) at which Word takes a face's script
#: fields.  Measured by ``make_script_offset_probe.py`` with faces of our own: at 0, 200
#: and 400 units of 2048 (0.195 em) Word ignores the face's script size *and* offsets and
#: draws its fallback (3/5, :data:`SCRIPT_SIZE_FALLBACK`); at 600 (0.293) it takes them;
#: real faces add Cambria's 0.246, taken.  So the bound lies in (0.195, 0.246]; 1/5 is a
#: placeholder inside it.  A negative ``ySubscriptYOffset`` (Charter's, and the probe's
#: ``Dx Negative``) is refused the same way.
SCRIPT_OFFSET_MINIMUM = Fraction(1, 5)


def script_fields_taken(face) -> bool:
    """Whether Word draws a ``w:vertAlign`` run with the face's own ``OS/2`` script size and
    offsets, or with its fallback: the size in :data:`SCRIPT_SIZE_RANGE`, the superscript
    offset at least :data:`SCRIPT_OFFSET_MINIMUM` of an em, the subscript offset not
    negative.  Offsets not recorded (``None``) are taken as fine."""
    ratio = Fraction(face.superscript_size, face.units_per_em) if face.superscript_size is not None else None
    if ratio is not None and not SCRIPT_SIZE_RANGE[0] <= ratio <= SCRIPT_SIZE_RANGE[1]:
        return False
    sup = getattr(face, "superscript_offset", None)
    sub = getattr(face, "subscript_offset", None)
    if sup is not None and Fraction(sup, face.units_per_em) < SCRIPT_OFFSET_MINIMUM:
        return False
    if sub is not None and sub < 0:
        return False
    return True


def script_half_points(half_points: int, vertical_align: str | None, face) -> int | None:
    """The size, in half points, Word draws a ``w:vertAlign`` run's glyphs at.

    **Not 2/3 of ``w:sz``** (what ``filesamples/sample1`` suggested): the face's own
    ``OS/2`` script size, ``w:sz x ySuperscriptYSize / unitsPerEm`` (``ySubscriptYSize``
    for a subscript), rounded half up to a whole half point and never below 2 (1 pt).
    sample1's "2/3" is Ubuntu's 0.65 x 24 = 15.6, rounded to 16.  ECMA-376 17.3.2.42 says
    only "in a smaller size"; it gives no value.

    Measured by ``make_script_probe.py`` (``script-sizes``: every ``w:sz`` from 2 to 96 in
    twelve faces, the size read from the run's advance to 1e-3 half point): 1,133 of
    1,140 superscripts and the same subscripts.  The seven that are not: ``w:sz`` 90 in
    the six faces at 0.65 is drawn at 59 where 58.49 rounds to 58, and Baskerville Old
    Face at ``w:sz`` 6 at 4 where 2.49 gives 2.  No rounding of the ratio (to a percent,
    per mille, 1/64, 1/256) or of the size in another unit reproduces 90 without breaking
    10, 30, 50 and 70 (6.499, 19.497... all drawn rounded down).

    The run's *line height* is not this: it takes part at ``w:sz``, on the baseline
    (:func:`docx2svg.vertical.item_extent`).  ``face`` is a
    :class:`~docx2svg.vertical.FaceMetrics`; returns ``None`` when it has no script
    sizes, and ``half_points`` for a run that is neither raised nor lowered.
    """
    if vertical_align not in ("superscript", "subscript"):
        return half_points
    size = face.superscript_size if vertical_align == "superscript" else face.subscript_size
    if size is None:
        return None
    ratio = Fraction(size, face.units_per_em)
    if not SCRIPT_SIZE_RANGE[0] <= ratio <= SCRIPT_SIZE_RANGE[1] or not script_fields_taken(face):
        ratio = SCRIPT_SIZE_FALLBACK
    return max(2, math.floor(half_points * ratio + Fraction(1, 2)))


def script_raise_half_points(half_points: int, vertical_align: str | None, face) -> int:
    """How far a ``w:vertAlign`` run's baseline is moved from its line's, in whole half
    points (up for a superscript, down for a subscript; the value is not negative).

    Measured by ``make_script_offset_probe.py``: 30 faces of our own, each moving one
    ``OS/2`` or ``hhea`` field, at ``w:sz`` 8-96 (690 lines); and held against the
    real faces of ``make_script_probe.py``'s size sweep (twelve faces, ``w:sz`` 2-96):

    * a script is moved by a **whole number of half points** (every one of 1,380 drawn
      offsets of the probe is ``round(k x 25/12)`` px);
    * where Word takes the face's script fields (:func:`script_fields_taken`), by the
      face's own offset, ``w:sz x ySuperscriptYOffset / upm`` rounded -- **but no more
      than the difference of the two sizes' ascents**, ``round(A x w:sz) - round(A x
      s)`` for the drawn size ``s`` and the face's ascent with its line gap ``A``
      (descent, for a subscript).  That bound is what makes Calibri, Times New Roman,
      Arial and Courier New (offsets 0.42-0.48 em) rise by about a third of the size;
    * where Word does not take them, by that bound alone, at the fallback size;
    * **not** bounded by the line: beside a 36 pt run, or at 24 pt among 11 pt text, a
      script moves exactly as it does alone at its own size.

    Scores (the rest of ROADMAP.md, "The script offset"): the probe's superscripts
    545 / 690 exact and 661 within a half point, subscripts 445 / 690 and 670; the real
    faces' 786 / 1,140 and 1,068, and 965 / 1,140 and 1,135.  The face's own offset
    alone: 309 / 690 and 131 / 690; 459 / 1,140 and 505 / 1,140.  What rounds the
    bound is not settled.
    """
    if vertical_align not in ("superscript", "subscript") or face is None:
        return 0
    drawn = script_half_points(half_points, vertical_align, face)
    if drawn is None:
        return 0
    upm = face.units_per_em
    extent = face.ascent + face.line_gap if vertical_align == "superscript" else face.descent
    bound = max(0, _round(Fraction(extent * half_points, upm)) - _round(Fraction(extent * drawn, upm)))
    offset = face.superscript_offset if vertical_align == "superscript" else face.subscript_offset
    if not script_fields_taken(face) or offset is None:
        return bound
    if offset <= 0:
        return 0
    return min(_round(Fraction(half_points * offset, upm)), bound)


def _round(value: Fraction) -> int:
    return math.floor(value + Fraction(1, 2))
