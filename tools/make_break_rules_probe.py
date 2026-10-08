#!/usr/bin/env python3
"""Where may a line break, and what counts toward its width?  One rule per family.

The δ sweep (``make_wrap_budget_probe.py``) measured the budget with nothing but letters
and spaces.  This probe asks what else Word does at a line's end, each case built so
that one question decides it: a first word ``F`` of filler letters is chosen so that
``F`` + a space + the text under test ends exactly δ (in 1/4096 pt) short of the right
edge, and what follows it then either fits or does not.  Widths are the recorded ones
(``probe_advances.py``), so the documents regenerate without a font file.

Families:

* ``after`` / ``before`` -- **break classes.**  ``aaaa`` + X + ``bbbb`` with ``aaaaX``
  fitting and ``aaaaXb`` not (``after``), or ``aaaa`` fitting and ``aaaaX`` not
  (``before``).  Where the line ends says whether Word may break after (before) X, or
  only at the space before ``aaaa``.  Some 55 characters -- hyphens and dashes,
  punctuation, brackets, quotes, symbols, the no-break, fixed-width and zero-width
  spaces, the soft hyphen -- between letters, and a dozen between digits and mixed.
  Calibri and Times New Roman.
* ``hang`` -- **what does not count at a line's end**: ``aaaa`` fitting exactly, then X
  and more text; if X hangs past the edge like a space, the line holds it.  For the
  spaces of every width, the no-break space and the tab.
* ``shy`` -- the soft hyphen, as ``w:softHyphen`` and as U+00AD: a break at it draws a
  hyphen, which must fit (δ just above and just below the hyphen's width).
* ``nbh`` -- ``w:noBreakHyphen``.
* ``tab`` -- a tab at the line's end whose stop is past the right edge; a tab then a word
  that does not fit after it; default stops (interval 720, and 567 in a second
  document); a custom stop clearing the defaults before it; the hanging indent's
  implicit stop; right and centre stops at the right edge, the text after them fitting
  exactly or by one unit too much.
* ``br`` -- ``w:br`` and ``w:cr``: the line ends there; spaces after it; a break whose
  text before it does not fit.
* ``kern`` -- with ``w:kern`` on: a pair counted inside the line (the line fits only
  with it); the pair of the line's last letter and the space after it (Times New Roman
  kerns ``A `` and ``Y ``); a word too long for the line split between two kerned
  letters; Aptos, whose legacy ``kern`` table and OpenType ``kern`` feature disagree;
  the ``w:kern`` threshold at, below and above the size.
* ``run`` -- a run boundary inside a word at the line's end (colour, size, face, bold).
* ``long`` -- a word too long for any line, after other words on the line.
* ``format`` -- ``w:caps``, ``w:spacing`` (tracking), ``w:w`` (scale) and a superscript
  at the line's end.
* ``label`` -- a list label wider than the hanging indent.

Measured by ``read_break_rules_probe.py`` through ``tools/breaks.py``.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass, replace
from fractions import Fraction

import probe_advances as pa
import probe_docx as w
import wml

COLUMN = 9020  # twips; the page is 11900 wide with 1440 margins, as in the δ sweep
PAGE_WIDTH = 11900
PAGE_HEIGHT = 16840
MARGIN = 1440
POOL = "Hnesoiatrdlmu"
TAIL = " Hnnn Hnnn"

CLASS_FACES = ("Calibri", "Times New Roman")
CLASS_CHARS = (
    "-", "\u2010", "\u2011", "\u2012", "\u2013", "\u2014", "\u2015", "/", "\\", "?", "!", ".", ",", ";",
    ":", ")", "(", "]", "[", "}", "{", '"', "'", "“", "”", "‘", "’", "%", "$", "&",
    "+", "=", "*", "#", "@", "_", "|", "~", "<", ">", "^", "`", "…", "·", "•", "¡",
    "¿", "€", "£", "°", "\u00a0", "\u202f", "\u2007", "\u2002", "\u2003", "\u2004", "\u2005",
    "\u2006", "\u2008", "\u2009", "\u200a", "\u200b", "\u00ad",
)
DIGIT_CHARS = ("-", "–", "—", "/", ".", ",", ":", "%", "+", "$", "(", ")")
HANG_CHARS = ("\u00a0", "\u2002", "\u2003", "\u2004", "\u2005", "\u2006", "\u2008", "\u2009", "\u200a",
              "\u200b", "\u202f", "\u2007")


@dataclass(frozen=True)
class Case:
    family: str
    name: str
    face: str
    half_points: int
    #: ``((text, props), ...)``; text may be a marker: ``<tab>``, ``<br>``, ``<cr>``,
    #: ``<shy>`` (w:softHyphen), ``<nbh>`` (w:noBreakHyphen).
    runs: tuple
    ind_right: int = 0
    ind_left: int = 0
    hanging: int = 0
    tabs: tuple = ()  # ((position twips, alignment), ...)
    numbered: bool = False
    setting: str = "default"  # the document: "default", or "tab567"

    def text(self) -> str:
        return "".join(text for text, _ in self.runs)


# -- widths -------------------------------------------------------------------------------


def width(face: str, text: str, hp: int, *, bold: bool = False, kern: str | None = None) -> Fraction:
    return pa.units(face, text, hp, bold=bold, kern=kern)


@functools.lru_cache(maxsize=None)
def _fillers(face: str, bold: bool = False) -> list:
    """Coin-change table: the shortest word of :data:`POOL` letters of each width."""
    advances = pa.load()[pa.key(face, bold, False)]["advances"]
    limit = COLUMN * 1024 // 5 // 16 + 1
    coins = sorted({advances[c]: c for c in POOL}.items())
    best: list = [None] * (limit + 1)
    best[0] = ""
    for total in range(1, limit + 1):
        for value, char in coins:
            if value <= total and best[total - value] is not None and (
                    best[total] is None or len(best[total - value]) + 1 < len(best[total])):
                best[total] = best[total - value] + char
    return [None if word is None else "".join(sorted(word, key=POOL.index)) for word in best]


def place(face: str, hp: int, used: Fraction, delta: int, *, bold: bool = False,
          start: int = 0, near: int = 5000) -> tuple[str, int]:
    """A filler word ``F`` and a right indent so that ``F`` + a space, then ``used``
    units of text, end ``delta`` units short of the right edge.  ``start`` is where the
    line starts (twips).  Returns ``(F, ind_right)``."""
    table = _fillers(face, bold)
    space = width(face, " ", hp, bold=bold)
    for step in range(0, 1500):
        for col in (near + 5 * step, near - 5 * step):
            if col < 1500 or col > COLUMN - start:
                continue
            want = Fraction(col * 1024, 5) - delta - used - space
            if want <= 0 or want.denominator != 1 or want % hp:
                continue
            index = int(want) // hp
            if index < len(table) and table[index] and len(table[index]) >= 3:
                return table[index], COLUMN - start - col
    raise RuntimeError(f"cannot place {face} {hp} {used} {delta}")


# -- families -----------------------------------------------------------------------------


def _plain(face: str, hp: int) -> dict:
    return {}


def _class_cases() -> list[Case]:
    out = []
    for face in CLASS_FACES:
        hp = 23
        advances = pa.load()[pa.key(face, False, False)]["advances"]
        contexts = [("aa", "aaaa", "bbbb", CLASS_CHARS), ("11", "1234", "5678", DIGIT_CHARS),
                    ("a1", "aaaa", "5678", ("-", "/", "–")), ("1a", "1234", "bbbb", ("-", "/", "–"))]
        for context, left, right, chars in contexts:
            for char in chars:
                if char not in advances:
                    continue
                x = width(face, char, hp)
                b = width(face, right[0], hp)
                # after: left+X fits by a third of the next glyph; left+X+right[0] does not
                delta = int(b / 3) or 1
                filler, ind_right = place(face, hp, width(face, left + char, hp), delta, near=3000 + len(out) * 5 % 4000)
                out.append(Case("after", f"{context} U+{ord(char):04X} {char}", face, hp,
                                ((f"{filler} {left}{char}{right}{TAIL}", {}),), ind_right))
                # before: left fits; left+X does not (X must have width)
                if x > 0:
                    delta = int(x / 3) or 1
                    filler, ind_right = place(face, hp, width(face, left, hp), delta, near=3000 + len(out) * 5 % 4000)
                    out.append(Case("before", f"{context} U+{ord(char):04X} {char}", face, hp,
                                    ((f"{filler} {left}{char}{right}{TAIL}", {}),), ind_right))
    return out


def _hang_cases() -> list[Case]:
    out = []
    for face in CLASS_FACES:
        hp = 23
        advances = pa.load()[pa.key(face, False, False)]["advances"]
        for char in HANG_CHARS + (" ",):
            if char not in advances:
                continue
            # aaaa ends exactly at the edge; then X, then more text.
            filler, ind_right = place(face, hp, width(face, "aaaa", hp), 0)
            out.append(Case("hang", f"U+{ord(char):04X} then word", face, hp,
                            ((f"{filler} aaaa{char}bbbb{TAIL}", {}),), ind_right))
            out.append(Case("hang", f"U+{ord(char):04X} then space", face, hp,
                            ((f"{filler} aaaa{char} bbbb{TAIL}", {}),), ind_right))
        filler, ind_right = place(face, hp, width(face, "aaaa", hp), 0)
        out.append(Case("hang", "tab then word", face, hp, ((f"{filler} aaaa", {}), ("<tab>", {}), (f"bbbb{TAIL}", {})),
                        ind_right))
    return out


def _shy_cases() -> list[Case]:
    out = []
    for face in CLASS_FACES:
        hp = 23
        hyphen = width(face, "-", hp)
        for form in ("element", "U+00AD"):
            runs_for = (lambda f: ((f"{f} aaaa", {}), ("<shy>", {}), (f"bbbb{TAIL}", {}))) if form == "element" else (
                lambda f: ((f"{f} aaaa\u00adbbbb{TAIL}", {}),))
            for label, delta in (("hyphen fits", int(hyphen) + 3), ("hyphen fits exactly", int(hyphen)),
                                 ("hyphen one unit over", int(hyphen) - 1), ("hyphen does not fit", int(hyphen / 2))):
                if hyphen.denominator != 1 and "exactly" in label:
                    continue
                filler, ind_right = place(face, hp, width(face, "aaaa", hp), delta)
                out.append(Case("shy", f"{form} {label}", face, hp, runs_for(filler), ind_right))
            # a soft hyphen where the line does not end: nothing drawn, no width
            filler, ind_right = place(face, hp, width(face, "aaaabbbb", hp), 0)
            out.append(Case("shy", f"{form} mid-line, word ending exactly", face, hp, runs_for(filler), ind_right))
    return out


def _nbh_cases() -> list[Case]:
    out = []
    for face in CLASS_FACES:
        hp = 23
        hyphen = width(face, "-", hp)
        for label, used, delta in (("after: aaaa- fits", "aaaa-", int(width(face, "b", hp) / 3)),
                                   ("before: aaaa fits", "aaaa", int(hyphen / 3)),
                                   ("the whole fits exactly", "aaaa-bbbb", 0)):
            filler, ind_right = place(face, hp, width(face, used, hp), delta)
            out.append(Case("nbh", label, face, hp, ((f"{filler} aaaa", {}), ("<nbh>", {}), (f"bbbb{TAIL}", {})),
                            ind_right))
    return out


def _exact_word(face: str, hp: int, units: Fraction) -> str | None:
    """A word of :data:`POOL` letters exactly ``units`` wide, or ``None``."""
    if units <= 0 or units.denominator != 1 or units % hp:
        return None
    table = _fillers(face)
    index = int(units) // hp
    return table[index] if index < len(table) else None


def _tab_cases() -> list[Case]:
    out = []
    face, hp = "Calibri", 23
    col = 6000
    right = COLUMN - col
    # A tab at the line's end whose next stop (the default) is past the edge.
    filler, ind_right = place(face, hp, width(face, "aaaa", hp), 200, near=4970)
    out.append(Case("tab", "default stop past the edge, then a word", face, hp,
                    ((f"{filler} aaaa", {}), ("<tab>", {}), (f"bbbb{TAIL}", {})), ind_right))
    # A custom left stop within the line, and a word after it ending delta short of the edge.
    lead = "Hn aa"
    for label, delta in (("fits exactly", 0), ("one unit over", -1), ("a third of a letter over", -300)):
        for stop in range(4000, 4400, 5):
            word = _exact_word(face, hp, Fraction((col - stop) * 1024, 5) - delta)
            if word:
                break
        else:
            raise RuntimeError("no tab word")
        out.append(Case("tab", f"custom stop, the word after it {label}", face, hp,
                        ((lead, {}), ("<tab>", {}), (f"{word}{TAIL}", {})), right, tabs=((stop, "left"),)))
    # A custom left stop past the edge, and a long text after it: where does it wrap?
    out.append(Case("tab", "custom left stop past the edge", face, hp,
                    (("Hn aa", {}), ("<tab>", {}), (f"bbbb{TAIL}", {})), right, tabs=((col + 400, "left"),)))
    words = " ".join(["Hnnes"] * 30)
    out.append(Case("tab", "custom left stop past the edge, long text", face, hp,
                    (("Hn aa", {}), ("<tab>", {}), (words, {})), right, tabs=((col + 400, "left"),)))
    out.append(Case("tab", "custom left stop past the edge, long text, section-wide column", face, hp,
                    (("Hn aa", {}), ("<tab>", {}), (words, {})), 0, tabs=((COLUMN + 400, "left"),)))
    # The default stop past the edge, a second time on the same line: two tabs.
    filler, ind_right = place(face, hp, width(face, "aaaa", hp), 200, near=4970)
    out.append(Case("tab", "default stop past the edge, two tabs", face, hp,
                    ((f"{filler} aaaa", {}), ("<tab>", {}), ("<tab>", {}), (f"bbbb{TAIL}", {})), ind_right))
    # A tab mid-line whose word after it does not fit: is the break before the tab or after?
    filler, ind_right = place(face, hp, width(face, "aa", hp), 1000, near=4000)
    out.append(Case("tab", "a word after a mid-line tab does not fit", face, hp,
                    ((f"{filler} aa", {}), ("<tab>", {}), (f"bbbbbbbbbbbbbbbbbbbbbbbb{TAIL}", {})), ind_right))
    # Default stops: several tabs in a row from short words.
    for setting in ("default", "tab567"):
        out.append(Case("tab", f"default stops ({setting})", face, hp,
                        (("a", {}), ("<tab>", {}), ("bb", {}), ("<tab>", {}), ("ccc", {}), ("<tab>", {}),
                         ("dddddddddd", {}), ("<tab>", {}), ("e", {})), right, setting=setting))
    # A custom stop at 3000: do the default stops before it survive?
    out.append(Case("tab", "custom stop at 3000 after a short word", face, hp,
                    (("a", {}), ("<tab>", {}), ("b", {}), ("<tab>", {}), ("c", {})), right, tabs=((3000, "left"),)))
    # The hanging indent's implicit stop.
    out.append(Case("tab", "hanging indent stop", face, hp,
                    (("a", {}), ("<tab>", {}), ("b", {}), ("<tab>", {}), ("c", {})), right, ind_left=1440,
                    hanging=1440))
    # Right and centre stops at the right edge, the text after them as wide as the room.
    name = width(face, "Name", hp)
    for alignment in ("right", "center"):
        room = Fraction(col * 1024, 5) - name
        if alignment == "right":
            variants = [(label, _exact_word(face, hp, room - delta), col)
                        for label, delta in (("fits exactly", 0), ("one unit over", -1))]
        else:
            # Centred at s, text t wide spans s - t/2 .. s + t/2.  A stop is whole twips,
            # so t is a multiple of 2048 font units (t/2 = hp * 5 twips per 1024).
            table = _fillers(face)
            k = max(k for k in range(1, len(table) // 2048) if table[2048 * k] and 2048 * k * hp <= room)
            word = table[2048 * k]
            half = k * hp * 5  # twips
            variants = [("ends exactly at the edge", word, col - half),
                        ("ends one twip past the edge", word, col - half + 1)]
        for label, word, stop in variants:
            if word is None:
                continue
            out.append(Case("tab", f"{alignment} stop, the text after it {label}", face, hp,
                            (("Name", {}), ("<tab>", {}), (word, {}), (TAIL, {})), right,
                            tabs=((stop, alignment),)))
        for label, stop in (("stop at the edge", col), ("stop past the edge", col + 400)):
            out.append(Case("tab", f"{alignment} {label}, short text", face, hp,
                            (("Name", {}), ("<tab>", {}), ("Date", {})), right, tabs=((stop, alignment),)))
    return out


def _word_of(face: str, hp: int, units: Fraction) -> str:
    table = _fillers(face)
    target = units / hp
    index = int(target)
    while index > 0 and not table[index]:
        index -= 1
    return table[index]


def _br_cases() -> list[Case]:
    out = []
    face, hp = "Calibri", 23
    for marker in ("<br>", "<cr>"):
        out.append(Case("br", f"{marker} mid-line", face, hp,
                        ((f"aaaa bbbb", {}), (marker, {}), (f"cccc{TAIL}", {})), 3000))
        out.append(Case("br", f"{marker} then spaces", face, hp,
                        ((f"aaaa", {}), (marker, {}), (f"   cccc{TAIL}", {})), 3000))
    filler, ind_right = place(face, hp, width(face, "aaaa", hp), -5)
    out.append(Case("br", "text before the break does not fit", face, hp,
                    ((f"{filler} aaaa", {}), ("<br>", {}), (f"cccc{TAIL}", {})), ind_right))
    filler, ind_right = place(face, hp, width(face, "aaaa", hp), 0)
    out.append(Case("br", "text before the break fits exactly", face, hp,
                    ((f"{filler} aaaa", {}), ("<br>", {}), (f"cccc{TAIL}", {})), ind_right))
    out.append(Case("br", "text before the break fits, spaces after it", face, hp,
                    ((f"{filler} aaaa   ", {}), ("<br>", {}), (f"cccc{TAIL}", {})), ind_right))
    return out


def _kern_cases() -> list[Case]:
    out = []
    on = {"kern": 2}
    for face, pair in (("Calibri", "AV"), ("Times New Roman", "AV"), ("Arial", "AV")):
        hp = 23
        text = pair * 3
        plain = width(face, text, hp)
        kerned = width(face, text, hp, kern="kern")
        gap = plain - kerned  # > 0
        for label, delta in (("fits only kerned", -int(gap / 2)), ("fits kerned exactly", -int(gap))):
            if gap.denominator != 1 and "exactly" in label:
                continue
            filler, ind_right = place(face, hp, plain, delta)
            out.append(Case("kern", f"{pair}x3 {label}", face, hp, ((f"{filler} {text}{TAIL}", on),), ind_right))
        # without w:kern the same line does not fit
        filler, ind_right = place(face, hp, plain, -int(gap / 2))
        out.append(Case("kern", f"{pair}x3 no w:kern", face, hp, ((f"{filler} {text}{TAIL}", {}),), ind_right))
    # The line's last letter and the space after it (Times New Roman kerns "A " by -113),
    # and the space and the letter after it (" A", -113).  Each placed so that the line
    # fits only if the pair is charged (delta = pair / 2 < 0), and once exactly at the edge.
    face, hp = "Times New Roman", 23
    kerns = pa.load()[pa.key(face, False, False)]["kern"]
    for last in ("A", "Y"):
        pair = Fraction(kerns.get(last + " ", 0) * hp)
        for label, delta in (("exactly at the edge", 0), ("fits only if the pair is charged", int(pair / 2))):
            text = "bbb" + last
            filler, ind_right = place(face, hp, width(face, text, hp), delta)
            out.append(Case("kern", f"'{last} ' at the line's end, {label} (pair {pair} units)", face, hp,
                            ((f"{filler} {text} Hnnn{TAIL}", on),), ind_right))
        pair = Fraction(kerns.get(" " + last, 0) * hp)
        # The filler's own last letter kerns with nothing; the pair is " " + last.
        for label, delta in (("exactly at the edge", 0), ("fits only if the pair is charged", int(pair / 2))):
            filler, ind_right = place(face, hp, width(face, last + "bb", hp), delta)
            out.append(Case("kern", f"' {last}' inside the line, {label} (pair {pair} units)", face, hp,
                            ((f"{filler} {last}bb Hnnn{TAIL}", on),), ind_right))
    # A word too long for the line, split between two kerned letters.
    face, hp = "Calibri", 23
    long_word = "AV" * 60
    out.append(Case("kern", "long AVAV word split", face, hp, ((f"{long_word}{TAIL}", on),), 5000))
    out.append(Case("kern", "long AVAV word split, no w:kern", face, hp, ((f"{long_word}{TAIL}", {}),), 5000))
    # Aptos: pairs in the OpenType feature and not in the legacy table.
    face, hp = "Aptos", 23
    text = "BoBoBoBoBo"
    legacy = width(face, text, hp, kern="kern")
    gpos = width(face, text, hp, kern="gpos")
    filler, ind_right = place(face, hp, legacy, -int((legacy - gpos) / 2))
    out.append(Case("kern", f"Aptos Bo: fits with the OpenType pairs only ({legacy - gpos} units)", face, hp,
                    ((f"{filler} {text}{TAIL}", on),), ind_right))
    filler, ind_right = place(face, hp, legacy, 0)
    out.append(Case("kern", "Aptos Bo: fits with the legacy pairs exactly", face, hp,
                    ((f"{filler} {text}{TAIL}", on),), ind_right))
    # The threshold: w:kern in half points, against the size.
    face = "Calibri"
    for hp, threshold in ((23, 23), (23, 24), (23, 22)):
        text = "AV" * 3
        plain = width(face, text, hp)
        gap = plain - width(face, text, hp, kern="kern")
        filler, ind_right = place(face, hp, plain, -int(gap / 2))
        out.append(Case("kern", f"threshold {threshold} at size {hp}", face, hp,
                        ((f"{filler} {text}{TAIL}", {"kern": threshold}),), ind_right))
    return out


def _run_cases() -> list[Case]:
    out = []
    face, hp = "Calibri", 23
    for label, props in (("colour", {"color": "C00000"}), ("size", {"sz": 29, "szCs": 29}),
                         ("face", {"rFonts": {"ascii": "Arial", "hAnsi": "Arial", "eastAsia": "Arial", "cs": "Arial"}}),
                         ("bold", {"b": True})):
        filler, ind_right = place(face, hp, width(face, "aaaa", hp), int(width(face, "b", hp) / 3))
        out.append(Case("run", f"run boundary inside a word ({label})", face, hp,
                        ((f"{filler} aaaa", {}), (f"bbbb", props), (TAIL, {})), ind_right))
    filler, ind_right = place(face, hp, width(face, "aaaa", hp), int(width(face, "b", hp) / 3))
    out.append(Case("run", "run boundary after a space", face, hp,
                    ((f"{filler} ", {}), ("aaaabbbb", {"color": "C00000"}), (TAIL, {})), ind_right))
    return out


def _long_cases() -> list[Case]:
    out = []
    face, hp = "Calibri", 23
    long_word = "abcdefghij" * 12
    out.append(Case("long", "a long word after two words", face, hp, ((f"Hnnn Hnnn {long_word}{TAIL}", {}),), 5000))
    out.append(Case("long", "a long word first", face, hp, ((f"{long_word}{TAIL}", {}),), 5000))
    out.append(Case("long", "a long word after a hyphen", face, hp, ((f"Hnnn Hnnn-{long_word}{TAIL}", {}),), 5000))
    return out


def _format_cases() -> list[Case]:
    out = []
    face, hp = "Calibri", 23
    b = int(width(face, "B", hp) / 3)
    filler, ind_right = place(face, hp, width(face, "AAAA", hp), b)
    out.append(Case("format", "caps: AAAA fits, B does not", face, hp,
                    ((f"{filler} ", {}), ("aaaab", {"caps": True}), (TAIL, {})), ind_right))
    track = 40  # twips per glyph
    extra = 5 * (track * 1024 // 5)
    filler, ind_right = place(face, hp, width(face, "aaaa", hp) + 4 * (track * 1024 // 5), int(width(face, "b", hp) / 3))
    out.append(Case("format", "tracking 40 twips: aaaa fits, b does not", face, hp,
                    ((f"{filler} ", {}), ("aaaab", {"spacing": track}), (TAIL, {})), ind_right))
    return out


def _label_cases() -> list[Case]:
    face, hp = "Calibri", 23
    return [Case("label", "label wider than the hanging indent", face, hp,
                 (("Item text that follows the label", {}),), 3000, ind_left=360, hanging=360, numbered=True)]


def _tag(index: int) -> str:
    """A word unique to each case, so that no two paragraphs read alike once whitespace
    is removed (the reader matches lines to paragraphs by text, and U+2002 is
    whitespace)."""
    letters = ""
    index += 26 * 26
    while index:
        index, digit = divmod(index, 26)
        letters = "abcdefghijklmnopqrstuvwxyz"[digit] + letters
    return "Q" + letters


@functools.lru_cache(maxsize=1)
def cases() -> tuple[Case, ...]:
    found = (_class_cases() + _hang_cases() + _shy_cases() + _nbh_cases() + _tab_cases() + _br_cases()
             + _kern_cases() + _run_cases() + _long_cases() + _format_cases() + _label_cases())
    return tuple(replace(case, runs=case.runs + ((f" {_tag(index)}", {}),)) for index, case in enumerate(found))


# -- the documents ------------------------------------------------------------------------


def _rpr(face: str, hp: int, props: dict) -> str:
    merged = {"rFonts": {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}, "sz": hp, "szCs": hp}
    merged.update(props)
    return wml.rpr(**merged)


def _run(text: str, face: str, hp: int, props: dict) -> str:
    content = {"<tab>": "<w:tab/>", "<br>": "<w:br/>", "<cr>": "<w:cr/>", "<shy>": "<w:softHyphen/>",
               "<nbh>": "<w:noBreakHyphen/>"}.get(text)
    if content is None:
        content = f'<w:t xml:space="preserve">{w.escape(text)}</w:t>'
    return f"<w:r>{_rpr(face, hp, props)}{content}</w:r>"


def _paragraph(case: Case) -> str:
    runs = "".join(_run(text, case.face, case.half_points, props) for text, props in case.runs)
    ppr = {"spacing": {"before": 0, "after": 120, "line": 240, "lineRule": "auto"}}
    if case.numbered:
        ppr["numPr"] = "<w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/>"
    if case.tabs:
        stops = "".join(f'<w:tab w:val="{alignment}" w:pos="{round(position)}"/>'
                        for position, alignment in case.tabs)
        ppr["tabs"] = stops
    ind = {"left": case.ind_left, "right": case.ind_right}
    if case.hanging:
        ind["hanging"] = case.hanging
    ppr["ind"] = ind
    mark = {"rFonts": {"ascii": case.face, "hAnsi": case.face, "eastAsia": case.face, "cs": case.face},
            "sz": case.half_points, "szCs": case.half_points}
    return wml.paragraph(runs, mark=mark, **ppr)


def section() -> str:
    return (
        "<w:sectPr>"
        f'<w:pgSz w:w="{PAGE_WIDTH}" w:h="{PAGE_HEIGHT}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )


SETTINGS = {"default": None, "tab567": 567}


def documents() -> dict[str, list[Case]]:
    out: dict[str, list[Case]] = {}
    for case in cases():
        out.setdefault(f"break-rules-{case.setting}", []).append(case)
    return out


def build(name: str) -> bytes:
    group = documents()[name]
    setting = group[0].setting
    extra = [wml.numbering_part([("decimal", "%1.)))))", {"ind": {"left": 360, "hanging": 360}}, {})])]
    if SETTINGS[setting] is not None:
        part = wml.settings_part({"val": "en-GB"})
        xml = part[3].replace("<w:themeFontLang", f'<w:defaultTabStop w:val="{SETTINGS[setting]}"/><w:themeFontLang')
        extra.append((*part[:3], xml))
    body = "".join(_paragraph(case) for case in group)
    return w.package(body, final_section=section(), extra_parts=tuple(extra))


if __name__ == "__main__":
    for name, group in documents().items():
        print(name, len(group), "paragraphs", len(build(name)), "bytes")
