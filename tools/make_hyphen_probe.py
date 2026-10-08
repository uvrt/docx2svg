#!/usr/bin/env python3
"""Word's automatic hyphenation (``w:autoHyphenation``): where, when and how it hyphenates.

Phase 3 measured the line breaker with hyphenation off; ``w:softHyphen`` and
``w:noBreakHyphen`` were measured there (3.3, ``shy`` and ``nbh``).  This probe turns
``w:autoHyphenation`` on and asks the questions that decide a line:

* **where** a word may break -- Word's own proofing hyphenator, chosen by the run's
  ``w:lang``, whose break points are observed here and never copied anywhere but the
  recording, which scores the openly licensed patterns of ``docx2svg.hyphenate``;
* **when** Word hyphenates at all -- the hyphenation zone (``w:hyphenationZone``) in each
  alignment and compatibility mode, the minimum word and fragment;
* **the hyphen** -- which glyph, how wide, where, and whether it must fit;
* **limits and exclusions** -- ``w:consecutiveHyphenLimit``, ``w:doNotHyphenateCaps``,
  ``w:suppressAutoHyphens``, ``w:noProof``, and words holding a soft, a hard or a
  non-breaking hyphen;
* **languages** -- en-US, nl-NL and de-DE (and their neighbours en-GB, nl-BE, de-AT,
  de-CH), compounds and spelling changes at a break.

Every document comes in three settings -- no compatibility mode, mode 14 and mode 15 --
since Word 2013's layout (mode 15) hyphenates by other rules.  Calibri unless said, laid
out from ``ooxml-common``'s advance tables, so the probe is rebuilt byte for byte with no
font file.  Every paragraph has ``w:keepLines`` and 18 pt after it, so its lines are found
in Word's PDF by the gaps between them (``read_hyphen_probe.py``).

The families
------------

``points`` (document ``points-*``)
    Every word of :data:`WORDS` at 24 pt after ``Hn``, once for every ``k`` from 1 to its
    length less one, in a column that ends where the word's first ``k`` letters and a
    hyphen end (rounded up to 5 twips, less than a letter): Word breaks the word at its
    last break point at or before ``k``, if it breaks it.  The zone is 0 (and at 24 pt
    mode 15's own threshold is passed by any fragment), so nothing but the points and the
    fragment minimums decides.
``zone`` (documents ``zone-*``)
    ``W1 infrastructure Hnnn Hnnn`` with ``W1`` composed so that the word starts exactly
    ``Z + δ`` units (1/4096 pt) before the right edge, δ in :data:`DELTAS`, for the zones
    the document states, in four alignments; and in mode 15, where the zone is not read,
    the gap at every whole twip from 336 to 364.
``fit``, ``shy``, ``compound``, ``caps``, ``suppress``, ``noproof``, ``lang``
    (``rules-*``; ``caps`` again in ``caps-*``, which sets ``w:doNotHyphenateCaps``)
    ``k``-sweeps of words that hold a soft, hard or non-breaking hyphen, in capitals, in a
    suppressed paragraph or a ``w:noProof`` run, and in other languages; and ``fit``, the
    hyphen's budget to the unit.
``text``, ``limit`` (``text-*``, ``limit1-*``, ``limit2-*``)
    Paragraphs of running text in the three languages, in narrow columns, left-aligned and
    justified: every line scored.  The ``limit`` documents set
    ``w:consecutiveHyphenLimit`` 1 and 2, and hold paragraphs hyphenated by hand with soft
    hyphens too.
``points`` again (``words-*``, ``french-*``)
    The sweep over :data:`MORE_WORDS`, 311 more English words, to tell open English
    pattern sets apart; and over :data:`FRENCH_WORDS` -- French words, accented letters and
    apostrophes, in fr-FR and again in fr-BE, fr-CH and fr-CA -- with French running text.
``bottom`` (``bottom-*``)
    A hyphenated paragraph run over a page's end after each of its lines, in the three
    settings and in mode 15 with each of [MS-DOCX]'s two compatSettings for a hyphen at
    the end of a page: every line scored, page by page.
"""

from __future__ import annotations

import functools
import math
import sys
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import probe_docx as w  # noqa: E402

from docx2svg.measure import TableAdvances  # noqa: E402

PAGE_WIDTH = 11900
PAGE_HEIGHT = 16840
MARGIN = 1440
COLUMN = PAGE_WIDTH - 2 * MARGIN  # 9020 twips
#: Units (1/4096 pt) in 5 twips.
FIVE = 1024
SHY = "­"  # a w:softHyphen in a case's text
NBH = "‑"  # a w:noBreakHyphen
SETTINGS = {"none": None, "14": 14, "15": 15}
POOL = "Hnesoiatrdlmu"
AFTER = 360  # twips after every paragraph

#: The probe words.  (``read_hyphen_probe.py`` refuses to record a line that ends in the
#: letters "i", "c" and a hyphen, a string this repository keeps out of its files.)
WORDS = {
    "en-US": (
        "infrastructure", "international", "representatives", "responsibility", "administrative",
        "telecommunications", "considerable", "demonstrates", "notwithstanding", "extraordinary",
        "organizations", "unprecedented", "hyphenation", "paragraph", "documentation", "measurement",
        "independent", "information", "development", "environment", "government", "understanding",
        "approximately", "everything", "generally", "establishment", "temperature", "afternoon",
        "probably", "table", "paper", "water", "basis", "over", "after", "river", "people", "number",
        "little", "simple", "second", "between", "company", "problem", "nevertheless", "relationship",
        "uncomfortable", "manufacturers",
    ),
    "nl-NL": (
        "ontwikkelingssamenwerking", "gemeentelijke", "verantwoordelijkheden", "beleidsmedewerkers",
        "arbeidsongeschiktheidsverzekeringen", "organisaties", "veroorzaakte", "onverwachte",
        "verschillende", "ziekenhuis", "tafel", "water", "appel", "boterham", "kinderen", "regering",
        "ontwikkeling", "gebruiker", "bedrijfsleven", "overheid", "wetenschappelijk", "vriendelijk",
        "geschiedenis", "belangrijk", "eigenlijk", "natuurlijk", "mogelijkheden", "ondernemers",
        "zeeën", "coördinatie", "beïnvloeden", "ideeën", "omaatje", "autootje", "menuutje",
        "cadeautje", "huisarts", "voetbalwedstrijd", "verkeersveiligheid", "ambtenaar", "deze",
        "ander", "kopje", "stoel", "boeken", "paarden", "studeren", "aardappelen",
    ),
    "de-DE": (
        "Bundesverfassungsgericht", "Arbeitsunfähigkeitsbescheinigung", "Verwaltungsschwierigkeiten",
        "Versicherungsgesellschaften", "unvorhergesehene", "Entscheidung", "verursachte",
        "Donaudampfschifffahrt", "Zucker", "backen", "Schifffahrt", "Brennnessel", "Straße", "Größe",
        "Mädchen", "Häuser", "Wissenschaft", "Wirtschaft", "Bundesrepublik",
        "Geschwindigkeitsbegrenzung", "Krankenhaus", "Fußballweltmeisterschaft", "Tisch", "Wasser",
        "Apfel", "Fenster", "Kinder", "Regierung", "Entwicklung", "Benutzer", "Möglichkeiten",
        "Unternehmen", "wahrscheinlich", "Verantwortung", "Zusammenarbeit", "Universität", "Sprache",
        "Beispiel", "Erdbeere", "Urinstinkt", "Staubecken", "Wachstube", "Abend", "Ofen", "Lehrerin",
        "Zeitungsartikel",
    ),
}


#: More English words for the ``points`` sweep (documents ``words-*``): ordinary dictionary
#: words chosen for their spread of lengths and endings -- none holding the letters "i" and
#: "c" side by side -- to tell openly licensed English pattern sets apart on more than the
#: 48 words of :data:`WORDS` (ROADMAP.md, 3.9.1).  Nothing chose them by Word's answer.
MORE_WORDS = {
    "en-US": (
        "abandonment", "absolutely", "academy", "accommodation", "accordingly", "accountability",
        "achievement", "acknowledgement", "adventure", "advertisement", "alternative",
        "ambassador", "anniversary", "announcement", "appearance", "appreciation", "appropriate",
        "architecture", "argument", "arrangement", "assessment", "assistance", "atmosphere",
        "attention", "authority", "availability", "background", "beautiful", "beginning",
        "believe", "benefit", "biography", "boundary", "brother", "butterfly", "calendar",
        "capability", "catastrophe", "celebration", "challenge", "character", "chocolate",
        "circumstances", "collaboration", "comfortable", "commitment", "committee", "community",
        "comparison", "competition", "complexity", "comprehensive", "concentration", "conference",
        "confidence", "consequence", "conservation", "constitution", "construction",
        "contemporary", "continuously", "contribution", "conversation", "cooperation",
        "corporation", "correspondence", "customer", "dangerous", "daughter", "decision",
        "definitely", "delivery", "departure", "description", "desperately", "destination",
        "determination", "difference", "disappointment", "discovery", "distribution", "education",
        "effectiveness", "elephant", "elimination", "emergency", "employment", "encouragement",
        "engineering", "entertainment", "enthusiasm", "equipment", "especially", "evaluation",
        "eventually", "examination", "excellent", "expectation", "experience", "experiment",
        "explanation", "extremely", "familiar", "fascinating", "favorite", "financial",
        "flexibility", "foundation", "frequently", "friendship", "furniture", "generation",
        "geography", "gradually", "grandmother", "happiness", "hospital", "household",
        "imagination", "immediately", "importance", "impossible", "improvement", "incredible",
        "individual", "industrial", "inevitable", "influence", "ingredient", "inheritance",
        "innovation", "inspiration", "installation", "institution", "instrument", "intelligence",
        "interesting", "interpretation", "introduction", "investigation", "invitation", "island",
        "jealousy", "journalist", "knowledge", "laboratory", "landscape", "language", "leadership",
        "legislation", "literature", "location", "machinery", "magazine", "maintenance",
        "management", "marvelous", "mechanism", "membership", "memory", "millionaire", "minister",
        "misunderstanding", "moreover", "mountain", "necessary", "negotiation", "neighborhood",
        "newspaper", "nonetheless", "nowadays", "occasionally", "opportunity", "opposition",
        "orchestra", "ordinary", "original", "otherwise", "ourselves", "partnership", "passenger",
        "perception", "performance", "permission", "personality", "perspective", "philosophy",
        "photograph", "population", "possibility", "potato", "preparation", "presentation",
        "president", "pressure", "previously", "principle", "priority", "procedure", "profession",
        "program", "pronunciation", "properly", "proportion", "prosperity", "protection",
        "psychology", "punishment", "quantity", "question", "recognition", "recommendation",
        "reduction", "reference", "refrigerator", "regulation", "relatively", "remarkable",
        "replacement", "reputation", "requirement", "reservation", "resolution", "restaurant",
        "revolution", "satisfaction", "secretary", "separately", "situation", "strawberry",
        "structure", "substantial", "successful", "suggestion", "supermarket", "surprisingly",
        "surrounding", "temporary", "territory", "theater", "thoroughly", "throughout", "together",
        "tomorrow", "tournament", "tradition", "transformation", "transportation", "tremendous",
        "umbrella", "understandable", "unemployment", "university", "unfortunately", "vegetable",
        "vocabulary", "volunteer", "wonderful", "yesterday", "also", "into", "only", "even",
        "open", "upon", "ever", "many", "very", "lady", "city", "baby", "hotel", "money", "women",
        "never", "under", "other", "about", "began", "happy", "seven", "eleven", "garden",
        "window", "yellow", "summer", "winter", "mother", "father", "sister", "letter", "matter",
        "better", "butter", "rather", "button", "pretty", "ready", "heavy", "sorry", "lemon",
        "honey", "robot", "final", "metal", "total", "label", "level", "novel", "model", "panel",
    ),
}

#: French (documents ``french-*``): the ``points`` sweep in fr-FR -- ordinary words, accented
#: letters, and words with an apostrophe, typed both as ``'`` and as ``’`` -- and a few of
#: them again in fr-BE, fr-CH and fr-CA, to see whether those are hyphenated at all.
FRENCH_WORDS = {
    "fr-FR": (
        "gouvernement", "développement", "responsabilité", "international", "extraordinaire",
        "connaissance", "malheureusement", "parlementaire", "environnement", "administration",
        "consommation", "fonctionnement", "représentation", "amélioration", "recommandation",
        "entreprise", "problème", "exemple", "personne", "nouveau", "ordinateur", "université",
        "constitution", "transport", "montagne", "campagne", "semaine", "ensemble", "quelque",
        "toujours", "beaucoup", "pendant", "travail", "maison", "table", "enfant", "école",
        "élève", "réalité", "événement", "préférence", "téléphone", "bibliothèque", "fenêtre",
        "château", "hôpital", "naïveté", "ambiguïté", "cœur", "œuvre", "garçon", "français",
        "leçon", "déjà", "être", "aussi", "avec", "elle", "idée", "aimer", "rose", "porte", "ami",
        "élu", "constitutionnellement", "l'homme", "l’homme", "aujourd'hui", "aujourd’hui",
        "jusqu'à", "jusqu’à", "l'organisation", "l’organisation", "d'administration",
        "qu'aujourd'hui", "presqu'île", "s'inscrire", "l'environnement",
    ),
    "fr-BE": ("gouvernement", "développement", "connaissance", "l'organisation", "bibliothèque"),
    "fr-CH": ("gouvernement", "développement", "connaissance", "l'organisation", "bibliothèque"),
    "fr-CA": ("gouvernement", "développement", "connaissance", "l'organisation", "bibliothèque"),
}


@functools.lru_cache(maxsize=None)
def _advances() -> TableAdvances:
    return TableAdvances()


def advance(face: str, char: str) -> int:
    """Font units (2048 per em) of ``char``: a soft hyphen draws nothing mid-line, a
    non-breaking hyphen is a hyphen (Phase 3.3)."""
    if char == SHY:
        return 0
    if char == NBH:
        char = "-"
    found = _advances().advance(face, False, False, char)
    if found is None or found[1] != 2048:
        raise KeyError(f"no advance for {char!r} in {face}")
    return found[0]


def units(face: str, text: str, hp: int, *, caps: bool = False) -> int:
    """``text``'s width in 1/4096 pt at ``hp`` half points (every advance is exact)."""
    return sum(advance(face, c.upper() if caps else c) for c in text) * hp


@dataclass(frozen=True)
class Case:
    family: str
    #: ``((text, run properties), ...)``; text may hold :data:`SHY` and :data:`NBH`.
    runs: tuple
    hp: int = 48
    face: str = "Calibri"
    jc: str = "left"
    ind_right: int = 0
    lang: str = "en-US"
    #: Extra paragraph properties: ``suppressAutoHyphens``, ``pStyle``.
    ppr: tuple = ()
    #: What the case asks, for the reader: the word under test, ``k``, δ...
    meta: tuple = field(default=(), compare=False)

    @property
    def text(self) -> str:
        return "".join(text for text, _ in self.runs)

    def info(self) -> dict:
        return dict(self.meta)


def _fonts(face: str) -> str:
    return f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="{face}" w:cs="{face}"/>'


def _rpr(face: str, hp: int, lang: str, props: dict) -> str:
    out = _fonts(face)
    for name in ("caps", "smallCaps", "noProof"):
        if props.get(name):
            out += f"<w:{name}/>"
    out += f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>'
    language = props.get("lang", lang)
    if language:
        out += f'<w:lang w:val="{language}"/>'
    return f"<w:rPr>{out}</w:rPr>"


def _run_xml(text: str, rpr: str) -> str:
    out = ""
    chunk = ""
    for char in text:
        if char in (SHY, NBH):
            if chunk:
                out += f'<w:t xml:space="preserve">{w.escape(chunk)}</w:t>'
                chunk = ""
            out += "<w:softHyphen/>" if char == SHY else "<w:noBreakHyphen/>"
        else:
            chunk += char
    if chunk:
        out += f'<w:t xml:space="preserve">{w.escape(chunk)}</w:t>'
    return f"<w:r>{rpr}{out}</w:r>"


def paragraph(case: Case) -> str:
    runs = "".join(_run_xml(text, _rpr(case.face, case.hp, case.lang, props)) for text, props in case.runs)
    ppr = ""
    props = dict(case.ppr)
    if "pStyle" in props:
        ppr += f'<w:pStyle w:val="{props["pStyle"]}"/>'
    ppr += "<w:keepLines/>"
    if "suppressAutoHyphens" in props:
        value = props["suppressAutoHyphens"]
        ppr += "<w:suppressAutoHyphens/>" if value else '<w:suppressAutoHyphens w:val="0"/>'
    ppr += f'<w:spacing w:before="0" w:after="{AFTER}" w:line="240" w:lineRule="auto"/>'
    ppr += f'<w:ind w:left="0" w:right="{case.ind_right}"/>'
    ppr += f'<w:jc w:val="{case.jc}"/>'
    ppr += _rpr(case.face, case.hp, case.lang, {})
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def right_for(width: int) -> int:
    """``w:ind/@w:right`` for a line ``width`` units wide, rounded up to 5 twips."""
    col = -(-width // FIVE) * 5
    if col > COLUMN:
        raise ValueError("wider than the column")
    return COLUMN - col


# -- k-sweeps -----------------------------------------------------------------------------

W1 = "Hn"
TAIL = " Hn"


def sweep(family: str, word_runs: tuple, *, hp: int = 48, lang: str = "en-US", caps: bool = False,
          ppr: tuple = (), face: str = "Calibri", ks=None, extra: tuple = ()) -> list[Case]:
    """``Hn <word> Hn`` once for every ``k``: a column that ends where ``Hn``, the space,
    the word's first ``k`` characters and a hyphen end, rounded up to 5 twips.  The word
    is ``word_runs`` -- ``((text, run properties), ...)``, its characters counted across
    them."""
    word = "".join(text for text, _ in word_runs)
    lead = units(face, W1 + " ", hp)
    hyphen = advance(face, "-") * hp
    out = []
    for k in (ks or range(1, len(word))):
        prefix = units(face, word[:k], hp, caps=caps)
        right = right_for(lead + prefix + hyphen)
        runs = ((W1 + " ", {}),) + tuple(word_runs) + ((TAIL, {}),)
        out.append(Case(family, runs, hp=hp, face=face, ind_right=right, lang=lang, ppr=ppr,
                        meta=(("word", word), ("k", k)) + extra))
    return out


def points_cases(words_by_lang: dict = WORDS) -> list[Case]:
    out = []
    for lang, words in words_by_lang.items():
        for word in words:
            out += sweep("points", ((word, {}),), lang=lang, extra=(("lang", lang),))
    return out


# -- exact compositions -------------------------------------------------------------------


@functools.lru_cache(maxsize=None)
def _sums(face: str) -> list:
    """For every width (font units) up to a column at 7 pt, the word of :data:`POOL`
    letters with the fewest characters that is exactly that wide, or ``None``."""
    limit = COLUMN * FIVE // 5 // 14 + 1
    coins = sorted({advance(face, c): c for c in POOL}.items())
    best: list = [None] * (limit + 1)
    best[0] = ""
    for total in range(1, limit + 1):
        for width, char in coins:
            if width <= total and best[total - width] is not None and (
                    best[total] is None or len(best[total - width]) + 1 < len(best[total])):
                best[total] = best[total - width] + char
    return [None if word is None else "H" + "".join(sorted(word, key=POOL.index)) for word in best]


def compose(face: str, hp: int, gap: int, after: str, *, near: int = 4000) -> tuple[str, int] | None:
    """``(W1, ind_right)``: a first word whose end is exactly ``gap`` units before the
    right edge, with ``after`` (the space and what follows it on the line) not counted,
    in a column near ``near`` twips.  ``W1`` starts with ``H``."""
    sums = _sums(face)
    h = advance(face, "H") * hp
    near -= near % 5
    for step in range(0, 1200):
        for col in (near + 5 * step, near - 5 * step):
            if not 600 <= col <= COLUMN:
                continue
            want = col * FIVE // 5 - gap - h
            if want > 0 and want % hp == 0 and want // hp < len(sums) and sums[want // hp] is not None:
                word = "H" + sums[want // hp][1:]
                if len(word) >= 3:
                    return word, COLUMN - col
    return None


#: δ about each zone, in units (1/4096 pt): the word's start is ``Z + δ`` before the edge.
DELTAS = (-1024, -205, -41, -3, -1, 0, 1, 3, 41, 205, 1024)
#: The zones the ``zone`` documents state (``None``: unstated), per document.
ZONE_DOCUMENTS = {
    "none": {"zone-none-unset": None, "zone-none-180": 180, "zone-none-360": 360, "zone-none-720": 720,
             "zone-none-0": 0},
    "14": {"zone-14-unset": None, "zone-14-180": 180, "zone-14-360": 360, "zone-14-720": 720, "zone-14-0": 0},
    "15": {"zone-15-unset": None, "zone-15-720": 720},
}
#: What an unstated zone is taken as when composing the δ sweep: 0.75 cm, the metric
#: ``Normal.dotm``'s (the exploration found Word's threshold there, not at 360).
UNSET_ZONE = 425
ZONE_WORD = "infrastructure"
ZONE_TAIL = " Hnnn Hnnn"


def twip_units(twips: int) -> int:
    return math.floor(twips * 4096 / 20 + 0.5)


def compose_near(face: str, hp: int, gap: int, *, near: int, side: int = 0) -> tuple[str, int, int]:
    """:func:`compose` for the composable gap nearest ``gap`` (at 16 pt only multiples of
    32 units are): ``(W1, ind_right, gap)``.  ``side`` keeps it on the same side of
    ``gap - side_origin``: 1 at or above ``gap``, -1 at or below, 0 either."""
    for step in range(0, 4 * hp):
        for candidate in ((gap + step, gap - step) if side == 0 else (gap + side * step,)):
            found = compose(face, hp, candidate, "", near=near)
            if found is not None:
                return found[0], found[1], candidate
    raise RuntimeError(f"no composition for {face} {hp} {gap}")


def zone_cases(setting: str, zone: int | None) -> list[Case]:
    out = []
    mode15 = setting == "15"
    for jc in ("left", "both", "right", "center"):
        for face, hp in (("Calibri", 22), ("Calibri", 32), ("Times New Roman", 22)):
            space = advance(face, " ") * hp
            if mode15:
                # The zone is not read: the gap (from the first word's end) at whole twips.
                gaps = range(336, 365) if jc == "left" or (face, hp) == ("Calibri", 22) else range(336, 365, 4)
                for near, twips in enumerate(gaps):
                    first, right, gap = compose_near(face, hp, twip_units(twips), near=2400 + 37 * near)
                    out.append(Case("zone", ((first + " " + ZONE_WORD + ZONE_TAIL, {}),), hp=hp, face=face, jc=jc,
                                    ind_right=right, meta=(("gap", gap), ("space", space))))
                continue
            # With no zone, the word's first fragment and its hyphen must fit: δ about that.
            z = twip_units(UNSET_ZONE if zone is None else zone) if zone != 0 else units(face, "in-", hp)
            for near, delta in enumerate(DELTAS):
                side = 1 if delta > 0 else -1 if delta < 0 else 0
                first, right, gap = compose_near(face, hp, z + space + delta, near=2400 + 211 * near, side=side)
                if side == 0 and gap != z + space:
                    continue  # no case exactly at the zone at this size
                out.append(Case("zone", ((first + " " + ZONE_WORD + ZONE_TAIL, {}),), hp=hp, face=face, jc=jc,
                                ind_right=right, meta=(("gap", gap), ("space", space), ("delta", gap - z - space))))
    return out


# -- the rules document -------------------------------------------------------------------


def fit_cases() -> list[Case]:
    """The hyphen's budget, to the unit: ``W1 infrastructure`` where the line through
    ``infra`` and a hyphen ends δ units from the edge (Word breaks there: ``points``)."""
    out = []
    for jc in ("left", "both"):
        for face, hp in (("Calibri", 22), ("Times New Roman", 22), ("Calibri", 29)):
            hyphen = advance(face, "-") * hp
            for near, delta in enumerate((-205, -41, -3, -1, 0, 1, 3, 41, 205)):
                # The line ends ``-delta`` past the edge: the gap from W1's end is the
                # space, "infra", the hyphen, less delta.
                base = units(face, " infra", hp) + hyphen
                side = 1 if delta < 0 else -1 if delta > 0 else 0
                first, right, gap = compose_near(face, hp, base - delta, near=3000 + 173 * near, side=side)
                if delta == 0 and gap != base:
                    continue
                out.append(Case("fit", ((first + " " + ZONE_WORD + ZONE_TAIL, {}),), hp=hp, face=face, jc=jc,
                                ind_right=right, meta=(("delta", base - gap),)))
    return out


SHY_WORDS = (
    "infra" + SHY + "structure",  # at a point of Word's
    "infr" + SHY + "astructure",  # not at one
    "infrastructu" + SHY + "re",  # late
    "con" + SHY + "siderable",
)
COMPOUNDS = (
    "telecommunications-infrastructure",
    "telecommunications" + NBH + "infrastructure",
    "well-established",
    "Arbeits-Unfähigkeit",
)


def rules_cases(setting: str) -> list[Case]:
    out = fit_cases()
    for word in SHY_WORDS:
        out += sweep("shy", ((word, {}),), hp=32)
    out += sweep("shy", (("infr" + SHY + "astructure", {}),), hp=32, ppr=(("suppressAutoHyphens", True),),
                 extra=(("suppressed", True),))
    for word in COMPOUNDS:
        out += sweep("compound", ((word, {}),), hp=32, lang="de-DE" if word[0] == "A" else "en-US")
    out += caps_cases()
    # Suppression: direct, through a paragraph style, and turned off again over the style.
    out += sweep("suppress", (("infrastructure", {}),), hp=32, ppr=(("suppressAutoHyphens", True),))
    out += sweep("suppress", (("infrastructure", {}),), hp=32, ppr=(("pStyle", "NoHyphens"),))
    out += sweep("suppress", (("infrastructure", {}),), hp=32,
                 ppr=(("pStyle", "NoHyphens"), ("suppressAutoHyphens", False)))
    out += sweep("noproof", (("infrastructure", {"noProof": True}),), hp=32)
    out += sweep("noproof", (("infra", {}), ("structure", {"noProof": True})), hp=32)
    # Languages: one string in each, and a language's neighbours.
    for word in ("Infrastruktur", "Demonstration", "parlementaire", "Wasserstoff", "information"):
        for lang in ("en-US", "nl-NL", "de-DE"):
            out += sweep("lang", ((word, {}),), hp=32, lang=lang, extra=(("lang", lang),))
    for word, langs in (("project", ("en-US", "en-GB")), ("knowledge", ("en-US", "en-GB")),
                        ("progress", ("en-US", "en-GB")), ("democracy", ("en-US", "en-GB")),
                        ("verantwoordelijkheid", ("nl-NL", "nl-BE")),
                        ("Verantwortung", ("de-DE", "de-AT", "de-CH")),
                        ("Straße", ("de-DE", "de-CH")),
                        ("infrastructure", ("fr-FR", "x-none")),
                        ("infrastructure", ("",))):
        for lang in langs:
            out += sweep("lang", ((word, {}),), hp=32, lang=lang, extra=(("lang", lang),))
    # A word whose language changes inside it.
    out += sweep("lang", (("Infra", {"lang": "de-DE"}), ("struktur", {"lang": "en-US"})), hp=32, lang="de-DE",
                 extra=(("lang", "de-DE+en-US"),))
    return out


def caps_cases() -> list[Case]:
    out = []
    for word, props, caps in (("infrastructure", {}, False), ("Infrastructure", {}, False),
                              ("INFRASTRUCTURE", {}, False), ("infrastructure", {"caps": True}, True),
                              ("InfraStructure", {}, False), ("DOCUMENTATION", {}, False),
                              ("Documentation", {}, False)):
        out += sweep("caps", ((word, props),), hp=32, caps=caps,
                     extra=(("form", word + (" (caps)" if caps else "")),))
    return out


# -- running text ------------------------------------------------------------------------

TEXT = {
    "en-US": ("The extraordinary development of telecommunications infrastructure demonstrates "
              "considerable responsibility among representatives of independent organizations, "
              "notwithstanding unprecedented administrative difficulties. Nevertheless the government "
              "understood that approximately everything depended on measurement, documentation and "
              "uncomfortable relationships between manufacturers."),
    "nl-NL": ("De ontwikkelingssamenwerking tussen verschillende gemeentelijke organisaties veroorzaakte "
              "onverwachte verantwoordelijkheden voor beleidsmedewerkers en arbeidsongeschiktheidsverzekeringen. "
              "Natuurlijk wilden de ondernemers eigenlijk een vriendelijk ziekenhuis, waar de geschiedenis van de "
              "verkeersveiligheid belangrijk bleef."),
    "de-DE": ("Die Bundesverfassungsgerichtsentscheidung über Arbeitsunfähigkeitsbescheinigungen verursachte "
              "unvorhergesehene Verwaltungsschwierigkeiten bei Versicherungsgesellschaften. Wahrscheinlich "
              "wollte die Regierung die Zusammenarbeit der Universitäten mit den Unternehmen fördern, deren "
              "Verantwortung für die Entwicklung der Wirtschaft wuchs."),
}
TEXT_COLUMNS = (2880, 3600, 4320)
#: French running text (documents ``french-*``), scored line by line as :data:`TEXT` is.
FRENCH_TEXT = ("Le développement extraordinaire des transports démontre aujourd'hui la responsabilité "
               "considérable des représentants d'organisations indépendantes, malgré des problèmes "
               "administratifs sans précédent. Néanmoins le gouvernement comprenait que l'environnement de "
               "l'entreprise dépendait presque entièrement de la consommation, de la documentation et des "
               "relations inconfortables entre les producteurs.")


def text_cases(texts: dict = TEXT) -> list[Case]:
    out = []
    for lang, text in texts.items():
        for col in TEXT_COLUMNS:
            for jc in ("left", "both"):
                out.append(Case("text", ((text, {}),), hp=22, jc=jc, ind_right=COLUMN - col, lang=lang,
                                meta=(("lang", lang), ("column", col))))
    return out


#: The pattern sets :func:`by_hand` hyphenates with, per language.
BY_HAND = {"en-US": "en-us", "de-DE": "de-1996"}


def by_hand(text: str, lang: str) -> str:
    """``text`` with a soft hyphen at every point Liang's patterns find (``docx2svg.hyphenate``),
    a paragraph hyphenated by hand, as some documents are: by the pattern sets the probe
    was first made with (:data:`BY_HAND`), so that it stays the document Word was shown."""
    from docx2svg import hyphenate

    patterns = hyphenate.load(BY_HAND[lang])
    out = []
    for token in text.split(" "):
        letters = "".join(c for c in token if c.isalpha())
        points = tuple(i for i in patterns.points(letters) if 2 <= i <= len(letters) - 2) \
            if letters == token.strip(".,") else ()
        pieces, last = [], 0
        for point in points:
            pieces.append(token[last:point])
            last = point
        pieces.append(token[last:])
        out.append(SHY.join(pieces))
    return " ".join(out)


def limit_cases() -> list[Case]:
    out = []
    for lang in ("en-US", "de-DE"):
        for col in (2880, 3600):
            for jc in ("left", "both"):
                out.append(Case("limit", ((TEXT[lang], {}),), hp=22, jc=jc, ind_right=COLUMN - col, lang=lang,
                                meta=(("lang", lang), ("column", col))))
        out.append(Case("limit", ((by_hand(TEXT[lang], lang), {}),), hp=22, ind_right=COLUMN - 2880, lang=lang,
                        ppr=(("suppressAutoHyphens", True),), meta=(("lang", lang), ("by hand", True))))
        out.append(Case("limit", ((by_hand(TEXT[lang], lang), {}),), hp=22, ind_right=COLUMN - 2880, lang=lang,
                        meta=(("lang", lang), ("by hand", True), ("auto", True))))
    return out


# -- the page's last line -----------------------------------------------------------------

#: The ``bottom`` documents: a hyphenated paragraph that runs over a page's end, the page
#: break put after its line ``j`` for every ``j`` by an empty paragraph of exact height before
#: it (``w:pageBreakBefore``), to see whether Word lets a page end in an automatic hyphen.
#: [MS-DOCX] 2.3.7 and 2.3.8 say how it does not: ``useWord2013TrackBottomHyphenation`` (the
#: hyphenated word's whole line goes to the next page when it is absent or true, only the word
#: when false) and ``allowHyphenationAtTrackBottom`` (true: a page may end in a hyphen).
BOTTOM_LINES = range(1, 17)
#: The line pitch of 11 pt Calibri, single (Phase 2): 2,500 / 2,048 em, in twips.
BOTTOM_PITCH = 11 * 20 * 2500 / 2048
BODY_HEIGHT = PAGE_HEIGHT - 2 * MARGIN
#: ``name suffix -> (mode, compatSettings)`` of the ``bottom`` documents.
BOTTOM_DOCUMENTS = {
    "none": (None, ()), "14": (14, ()), "15": (15, ()),
    "15-word2013-0": (15, (("useWord2013TrackBottomHyphenation", 0),)),
    "15-allow-1": (15, (("allowHyphenationAtTrackBottom", 1),)),
}


def bottom_cases() -> list[Case]:
    out = []
    for jc in ("left", "both"):
        for j in BOTTOM_LINES:
            out.append(Case("bottom", ((TEXT["en-US"], {}),), hp=22, jc=jc, ind_right=COLUMN - 2880,
                            meta=(("lines", j), ("jc", jc))))
    return out


def bottom_paragraphs(case: Case) -> str:
    """An empty paragraph starting a page, ``BODY_HEIGHT`` less ``j + 1/2`` lines high, then
    the case's paragraph -- no ``w:keepLines``, no widow control -- with nothing after it."""
    spacer = round(BODY_HEIGHT - (case.info()["lines"] + 0.5) * BOTTOM_PITCH)
    runs = "".join(_run_xml(text, _rpr(case.face, case.hp, case.lang, props)) for text, props in case.runs)
    rpr = _rpr(case.face, case.hp, case.lang, {})
    return (f'<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="{spacer}" '
            f'w:lineRule="exact"/>{rpr}</w:pPr></w:p>'
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:ind w:left="0" w:right="{case.ind_right}"/><w:jc w:val="{case.jc}"/>{rpr}</w:pPr>{runs}</w:p>')


# -- documents ----------------------------------------------------------------------------

SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
SETTINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"


def settings_part(mode: int | None, *, zone: int | None = None, limit: int | None = None,
                  caps: bool = False, compat: tuple = ()) -> tuple[str, str, str, str]:
    """``settings.xml`` with ``w:autoHyphenation`` and what else is given, in schema order."""
    inner = "<w:autoHyphenation/>"
    if limit is not None:
        inner += f'<w:consecutiveHyphenLimit w:val="{limit}"/>'
    if zone is not None:
        inner += f'<w:hyphenationZone w:val="{zone}"/>'
    if caps:
        inner += "<w:doNotHyphenateCaps/>"
    settings = ((("compatibilityMode", mode),) if mode is not None else ()) + tuple(compat)
    if settings:
        inner += "<w:compat>" + "".join(
            f'<w:compatSetting w:name="{name}" w:uri="http://schemas.microsoft.com/office/word" w:val="{value}"/>'
            for name, value in settings) + "</w:compat>"
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<w:settings xmlns:w="{w.W_NS}">{inner}</w:settings>')
    return ("word/settings.xml", SETTINGS_CT, SETTINGS_REL, xml)


def styles() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{w.W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>{_fonts("Calibri")}'
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>'
        "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>"
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        "</w:pPr></w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="NoHyphens"><w:name w:val="No Hyphens"/>'
        "<w:pPr><w:suppressAutoHyphens/></w:pPr></w:style>"
        "</w:styles>"
    )


def section() -> str:
    return (
        "<w:sectPr>"
        f'<w:pgSz w:w="{PAGE_WIDTH}" w:h="{PAGE_HEIGHT}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )


def names() -> list[str]:
    """Every document's name, without building any (:func:`documents` does)."""
    out = []
    for setting in SETTINGS:
        out += [f"points-{setting}", *ZONE_DOCUMENTS[setting], f"rules-{setting}", f"caps-{setting}",
                f"text-{setting}", f"limit1-{setting}", f"limit2-{setting}", f"limit0-{setting}",
                f"words-{setting}", f"french-{setting}"]
    return out + [f"bottom-{suffix}" for suffix in BOTTOM_DOCUMENTS]


@functools.lru_cache(maxsize=None)
def documents() -> dict[str, tuple[dict, tuple[Case, ...]]]:
    """Every document: name -> (its settings, its cases)."""
    out: dict = {}
    for setting, mode in SETTINGS.items():
        out[f"points-{setting}"] = ({"mode": mode, "zone": 0}, tuple(points_cases()))
        for name, zone in ZONE_DOCUMENTS[setting].items():
            out[name] = ({"mode": mode, "zone": zone}, tuple(zone_cases(setting, zone)))
        out[f"rules-{setting}"] = ({"mode": mode, "zone": 0}, tuple(rules_cases(setting)))
        out[f"caps-{setting}"] = ({"mode": mode, "zone": 0, "caps": True}, tuple(caps_cases()))
        out[f"text-{setting}"] = ({"mode": mode}, tuple(text_cases()))
        for limit in (1, 2):
            out[f"limit{limit}-{setting}"] = ({"mode": mode, "zone": 0, "limit": limit}, tuple(limit_cases()))
        out[f"limit0-{setting}"] = ({"mode": mode, "zone": 0}, tuple(limit_cases()))
        out[f"words-{setting}"] = ({"mode": mode, "zone": 0}, tuple(points_cases(MORE_WORDS)))
        out[f"french-{setting}"] = ({"mode": mode, "zone": 0},
                                    tuple(points_cases(FRENCH_WORDS) + text_cases({"fr-FR": FRENCH_TEXT})))
    for suffix, (mode, compat) in BOTTOM_DOCUMENTS.items():
        out[f"bottom-{suffix}"] = ({"mode": mode, "compat": compat}, tuple(bottom_cases()))
    return out


def build(name: str) -> bytes:
    options, cases = documents()[name]
    body = "".join(bottom_paragraphs(case) if case.family == "bottom" else paragraph(case) for case in cases)
    return w.package(body, final_section=section(), styles=styles(), extra_parts=(settings_part(**options),))


if __name__ == "__main__":
    for name, (options, cases) in documents().items():
        print(f"{name:24} {len(cases):5} paragraphs {len(build(name)):8} bytes")
