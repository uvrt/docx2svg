"""Where a word may break: Liang's algorithm over openly licensed patterns.

Word hyphenates with its own proofing tools, one hyphenator per language, and those are
Microsoft's: nothing of them is here.  What is here is Frank Liang's method (*Word
Hy-phen-a-tion by Com-put-er*, 1983), written from its published description, over the
patterns TeX distributions ship in ``hyph-utf8`` -- each file's source and licence are in
``patterns/SOURCES.txt`` beside it.  ``make_hyphen_probe.py`` observes where Word breaks
its probe words, and ``read_hyphen_probe.py`` scores these patterns against that (ROADMAP.md,
3.9); nothing Word answered is written back into the patterns.

The method: a pattern is a string of letters with a digit between any two of them (and at
either end) -- ``.ach4``, ``1tion`` -- where ``.`` marks a word's edge.  Every pattern that
occurs in ``.word.`` raises the value between the letters it covers to its digit, and the
word may break where the largest value is odd.  A list of exceptions spells some words out
with their hyphens instead.

Before the patterns a language may have word lists (``WORD_LISTS``): every word they hold
breaks where they say, and only a word none of them holds goes to the patterns.  English has
the Moby Hyphenator list (Grady Ward, public domain; ROADMAP.md, 3.9.1), and before it
the probe's own English words that Word breaks otherwise, as the probe observed Word
breaking them (``en-observed-exceptions.txt``: Word's behaviour, not its data).

Standard library only.
"""

from __future__ import annotations

import functools
import gzip
from pathlib import Path

PATTERNS = Path(__file__).resolve().parent / "patterns"

#: ``w:lang`` -> the pattern set Word's hyphenator for that language is modelled by,
#: looked up by the whole tag and then by its primary subtag.  Measured
#: (``make_hyphen_probe.py``, ``lang``): Word breaks en-GB words where it breaks en-US
#: ones, and de-AT and de-CH words where it breaks de-DE ones; it does not hyphenate nl-BE
#: at all on the machine measured (it seems to have no Belgian Dutch hyphenator installed),
#: nor ``x-none``; and French (``french``) in fr-FR, fr-CH and fr-CA alike, but not in
#: fr-BE.  A language without patterns here is not hyphenated.  English is Knuth's
#: ``hyphen.tex`` with the TUGboat exception log (``en-us-knuth``): of the open English
#: sets it agrees with Word's breaks most often (ROADMAP.md, 3.9.1), and breaks the words
#: the Moby list does not hold (``WORD_LISTS``); hyph-utf8's ``en-us`` is kept for the
#: probe's paragraphs hyphenated by hand.
PATTERN_SETS = {"en": "en-us-knuth", "nl": "nl", "nl-be": None, "de": "de-1996", "fr": "fr", "fr-be": None}

#: Pattern set -> the word lists consulted before it, in order (``patterns/<name>``: one
#: word per line, ``WORD_MARK`` at each point, and after a space the word as it breaks in
#: compatibility mode 15 where that differs; ``.gz`` compressed).  English: the words of
#: the probe that Word breaks otherwise than the rest would, as it was observed breaking
#: them (``read_hyphen_probe.py --exceptions``); then the Moby Hyphenator list, which
#: agrees with Word's breaks more often than any pattern set (ROADMAP.md, 3.9.1); the
#: patterns break only the words neither holds.
WORD_LISTS = {"en-us-knuth": ("en-observed-exceptions.txt", "en-moby.txt.gz")}
#: A word list's break mark: not a hyphen, which a word's own spelling may hold.
WORD_MARK = "="

#: Word's shortest fragments, before and after a break (``points``, every language): two
#: letters before it in every mode; two after it below compatibility mode 15 and three in
#: mode 15, which callers pass (``linebreak.auto_hyphens``).
LEFT_MIN = 2
RIGHT_MIN = 2


class Patterns:
    """One language's patterns and exceptions."""

    def __init__(self, patterns: list[str], exceptions: list[str]) -> None:
        self.table: dict[str, tuple[int, ...]] = {}
        self.longest = 0
        for pattern in patterns:
            letters = "".join(c for c in pattern if not c.isdigit())
            values = [0] * (len(letters) + 1)
            at = 0
            for c in pattern:
                if c.isdigit():
                    values[at] = int(c)
                else:
                    at += 1
            self.table[letters] = tuple(values)
            self.longest = max(self.longest, len(letters))
        self.exceptions = {item.replace("-", ""): tuple(_positions(item)) for item in exceptions}

    def values(self, word: str) -> list[int]:
        """The value between every two letters of ``word`` (index ``i``: before letter
        ``i``), lower case."""
        text = f".{word}."
        values = [0] * (len(text) + 1)
        for i in range(len(text)):
            for j in range(i + 1, min(len(text), i + self.longest) + 1):
                found = self.table.get(text[i:j])
                if found is not None:
                    for k, value in enumerate(found):
                        if value > values[i + k]:
                            values[i + k] = value
        return values[1:-1]

    def points(self, word: str) -> tuple[int, ...]:
        """Every ``i`` (1 .. len - 1) after which ``word`` may break, before minimums."""
        lower = word.lower()
        if lower in self.exceptions:
            return self.exceptions[lower]
        values = self.values(lower)
        return tuple(i for i in range(1, len(word)) if values[i] % 2)


def _positions(spelled: str, mark: str = "-") -> list[int]:
    out, at = [], 0
    for c in spelled:
        if c == mark:
            out.append(at)
        else:
            at += 1
    return out


def _lines(name: str) -> list[str]:
    path = PATTERNS / name
    if not path.exists():
        return []
    return [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


@functools.lru_cache(maxsize=None)
def load(name: str) -> Patterns:
    """The pattern set ``name`` (``en-us``, ``nl``...)."""
    return Patterns(_lines(f"hyph-{name}.pat.txt"), _lines(f"hyph-{name}.hyp.txt"))


class WordList:
    """Words with their break points, looked up by their case-folded spelling."""

    def __init__(self, entries: list[str]) -> None:
        self.words: dict[str, list[str]] = {}
        for entry in entries:
            self.words.setdefault(entry.split(" ")[0].replace(WORD_MARK, "").casefold(), []).append(entry)

    def points(self, word: str, mode15: bool = False) -> tuple[int, ...] | None:
        """Where the list breaks ``word``, or ``None`` if it does not hold it.  A spelling
        the list holds as typed is taken first (``Polish``, ``polish``), then the lower-case
        one (a capital at a sentence's start, or a word in capitals), then the first."""
        found = self.words.get(word.casefold())
        if found is None:
            return None
        spellings = {entry.split(" ")[0].replace(WORD_MARK, ""): entry for entry in reversed(found)}
        entry = (spellings.get(word) or spellings.get(word.lower()) or found[0]).split(" ")
        return tuple(_positions(entry[-1] if mode15 else entry[0], WORD_MARK))


@functools.lru_cache(maxsize=None)
def word_list(name: str) -> WordList:
    """The word list ``patterns/<name>``, read once, when a word first needs it."""
    path = PATTERNS / name
    if not path.exists():
        return WordList([])
    data = path.read_bytes()
    if name.endswith(".gz"):
        data = gzip.decompress(data)
    return WordList([line for line in data.decode("utf-8").splitlines() if line])


def word_points(word: str, name: str, mode15: bool = False, lists: tuple[str, ...] | None = None) -> tuple[int, ...]:
    """Every ``i`` after which ``word`` may break by pattern set ``name`` and the word lists
    before it (``lists``, by default ``WORD_LISTS``'), before minimums."""
    for listed in WORD_LISTS.get(name, ()) if lists is None else lists:
        found = word_list(listed).points(word, mode15)
        if found is not None:
            return found
    return load(name).points(word)


def pattern_set(lang: str | None) -> str | None:
    """The pattern set for a ``w:lang`` value, or ``None`` (no hyphenation)."""
    if not lang:
        return None
    tag = lang.lower()
    if tag in PATTERN_SETS:
        return PATTERN_SETS[tag]
    return PATTERN_SETS.get(tag.split("-")[0])


def points(word: str, lang: str | None, *, left_min: int = LEFT_MIN, right_min: int = RIGHT_MIN,
           mode15: bool = False) -> tuple[int, ...]:
    """Where ``word`` (letters only) may break in language ``lang``: every ``i`` such that
    ``word[:i]`` and ``word[i:]`` may end and start a line, at least ``left_min`` and
    ``right_min`` letters long; ``mode15``, as Word breaks it in compatibility mode 15."""
    name = pattern_set(lang)
    if name is None:
        return ()
    return tuple(i for i in word_points(word, name, mode15) if left_min <= i <= len(word) - right_min)
