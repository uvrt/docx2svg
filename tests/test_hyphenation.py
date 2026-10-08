"""Automatic hyphenation (``w:autoHyphenation``): ``tools/make_hyphen_probe.py`` against
Word, recorded in ``tests/fixtures/hyphen-observations.json`` (ROADMAP.md, Phase 3.9).

Every probe document is laid out by the library from the faces' recorded numbers and its
lines held to Word's: the first line of every swept case, every line of running text.
Where they differ, the patterns break a word elsewhere than Word's own hyphenator does
-- the rules (the zone, the fit, the minimums, the limit, the exclusions) agree in every
case; the counts below pin both.  The patterns' agreement with Word's break points is
pinned per language and mode, and every hyphen Word drew at a line's end in running text
is held to where the library draws it."""

from __future__ import annotations

import functools
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_hyphen_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
FACES = json.loads(reader.FACES.read_text(encoding="utf-8"))

#: Paragraphs whose recorded lines the library's lines are, per document.  Every miss is
#: a word the patterns break elsewhere than Word (``telecommunications``, ``responsibility``,
#: ``measurement``, ``connaissance``...), or a word whose language changes inside it.
EXPECTED = {
    "points-14": (1372, 1394), "points-15": (1365, 1394), "points-none": (1372, 1394),
    "french-14": (771, 798), "french-15": (772, 798), "french-none": (771, 798),
    "rules-14": (684, 734), "rules-15": (687, 734), "rules-none": (684, 734),
}
#: Families in which every case agrees, in every document.
EXACT_FAMILIES = ("zone", "fit", "caps", "suppress", "noproof")
#: ``(words whose points are Word's, words, points in common, ours, Word's)`` per language.
#: English is every word since the probe's words that Word breaks otherwise are the
#: exceptions (``en-observed-exceptions.txt``); without them, ``WITHOUT_EXCEPTIONS``.
PATTERNS = {
    "french-14": {"fr-FR": (67, 78, 122, 132, 126), "fr-BE": (5, 5, 0, 0, 0), "fr-CH": (4, 5, 14, 14, 15),
                  "fr-CA": (4, 5, 14, 14, 15)},
    "french-15": {"fr-FR": (69, 78, 121, 125, 126), "fr-BE": (5, 5, 0, 0, 0), "fr-CH": (4, 5, 14, 14, 15),
                  "fr-CA": (4, 5, 14, 14, 15)},
    "french-none": {"fr-FR": (67, 78, 122, 132, 126), "fr-BE": (5, 5, 0, 0, 0), "fr-CH": (4, 5, 14, 14, 15),
                    "fr-CA": (4, 5, 14, 14, 15)},
    "points-14": {"en-US": (48, 48, 117, 117, 117), "nl-NL": (46, 48, 119, 120, 120),
                  "de-DE": (42, 46, 110, 112, 114)},
    "points-15": {"en-US": (48, 48, 113, 113, 113), "nl-NL": (46, 48, 111, 112, 112),
                  "de-DE": (42, 46, 104, 106, 108)},
    "points-none": {"en-US": (48, 48, 117, 117, 117), "nl-NL": (46, 48, 119, 120, 120),
                    "de-DE": (42, 46, 110, 112, 114)},
    "words-14": {"en-US": (311, 311, 663, 663, 663)},
    "words-15": {"en-US": (311, 311, 609, 609, 609)},
    "words-none": {"en-US": (311, 311, 663, 663, 663)},
}
#: English as the Moby list and the patterns break it, the exceptions left out: the only
#: score that says anything of words the probe has not seen (ROADMAP.md, 3.9.1).
WITHOUT_EXCEPTIONS = {
    "points-14": (40, 48, 114, 122, 117), "points-15": (43, 48, 110, 115, 113), "points-none": (40, 48, 114, 122, 117),
    "words-14": (277, 311, 663, 698, 663), "words-15": (299, 311, 609, 621, 609), "words-none": (277, 311, 663, 698, 663),
}
#: Word's line-end hyphens in running text the library draws at the same place, and all
#: of them; the rest are on lines whose words the patterns break elsewhere.
HYPHENS = {"text-none": (98, 98), "text-14": (98, 98), "text-15": (105, 105)}


@functools.cache
def _ours(name: str) -> list[list[str]]:
    return reader.model(name, reader.probe.build(name), render_record.RecordedFonts(FACES))


def _verdicts(name: str) -> list[bool]:
    return reader.score(name, DATA["documents"][name], _ours(name))


@pytest.mark.parametrize("name", reader.probe.names())
def test_lines_break_where_word_breaks_them(name):
    verdicts = _verdicts(name)
    cases = reader.probe.documents()[name][1]
    assert len(verdicts) == len(cases) == len(DATA["documents"][name])
    for family in EXACT_FAMILIES:
        assert all(v for v, case in zip(verdicts, cases) if case.family == family), family
    assert (sum(verdicts), len(verdicts)) == EXPECTED.get(name, (len(cases), len(cases)))


@pytest.mark.parametrize("name", sorted(PATTERNS))
def test_patterns_against_words_points(name):
    firsts = DATA["documents"][name]
    assert reader.pattern_agreement(name, firsts) == PATTERNS[name]


@pytest.mark.parametrize("name", sorted(WITHOUT_EXCEPTIONS))
def test_english_without_the_observed_exceptions(name, monkeypatch):
    from docx2svg import hyphenate

    lists = tuple(n for n in hyphenate.WORD_LISTS["en-us-knuth"] if n != reader.EXCEPTIONS.name)
    monkeypatch.setitem(hyphenate.WORD_LISTS, "en-us-knuth", lists)
    assert reader.pattern_agreement(name, DATA["documents"][name])["en-US"] == WITHOUT_EXCEPTIONS[name]


def test_the_observed_exceptions_are_the_recordings():
    """``en-observed-exceptions.txt`` is what ``read_hyphen_probe.py --exceptions`` makes of
    the recording: the probe's English words that the Moby list and the patterns break
    otherwise than Word, with Word's points -- 42 words, 8 with points of their own in
    mode 15 (``bene-fit``, ``ben-e-fit``)."""
    lines = reader.EXCEPTIONS.read_text(encoding="utf-8").splitlines()
    assert lines == reader.observed_exceptions(DATA["documents"])
    assert len(lines) == 42 and sum(" " in line for line in lines) == 8
    assert not reader.FORBIDDEN.search("\n".join(lines)) and "-" not in "".join(lines)


@pytest.mark.parametrize("name", sorted(HYPHENS))
def test_line_end_hyphens_are_drawn_where_word_draws_them(name):
    ours = reader.model_hyphen_glyphs(reader.probe.build(name), render_record.RecordedFonts(FACES))
    assert reader.hyphens_placed(DATA["positions"][name], ours) == HYPHENS[name]


def test_the_recording_holds_no_forbidden_string():
    assert not reader.FORBIDDEN.search(json.dumps(DATA, ensure_ascii=False))


# -- the pieces -------------------------------------------------------------------------------


def test_liang_breaks_words_as_the_patterns_say():
    from docx2svg import hyphenate

    def spelt(word, lang, **kw):
        out, last = [], 0
        for point in hyphenate.points(word, lang, **kw):
            out.append(word[last:point])
            last = point
        return "-".join(out + [word[last:]])

    # English: the Moby list first, as Word's dictionary breaks; Knuth's patterns after it.
    assert spelt("independent", "en-US") == "in-de-pend-ent"  # Moby's (the patterns: in-de-pen-dent)
    assert spelt("associate", "en-GB") == "as-so-ci-ate"  # Moby's (hyphen.tex's own: as-so-ciate)
    assert spelt("circumstances", "en-US") == "cir-cum-stanc-es"  # Moby lists it in a compound only
    assert spelt("Polish", "en-US") == "Po-lish" and spelt("polish", "en-US") == "pol-ish"
    assert spelt("Education", "en-US") == "Educa-tion" and spelt("education", "en-US") == "ed-u-ca-tion"
    assert spelt("EDUCATION", "en-US") == "ED-U-CA-TION"  # no spelling in capitals: the lower-case one
    assert spelt("demonstrates", "en-US") == "demon-strates"  # an observed exception
    assert spelt("benefit", "en-US") == "bene-fit"  # an observed exception, and in mode 15:
    assert spelt("benefit", "en-US", right_min=3, mode15=True) == "ben-e-fit"
    assert spelt("documents", "en-US") == "doc-u-ments"  # Moby's
    assert spelt("pharmacology", "en-US") == "phar-ma-col-o-gy"
    assert hyphenate.load("en-us-knuth").points("hyphenation") == (2, 6, 7)  # a TUGboat exception
    assert hyphenate.load("en-us-knuth").points("associate") == (2, 4)  # one of hyphen.tex's own
    assert hyphenate.load("en-us").points("hyphenation") == (2, 6)  # hyph-utf8's, kept
    assert spelt("Bundesrepublik", "de-CH") == "Bun-des-re-pu-blik"
    assert spelt("verantwoordelijkheden", "nl-NL") == "ver-ant-woor-de-lijk-he-den"
    assert spelt("gemeentelijke", "nl-NL", right_min=3) == "ge-meen-te-lijke"
    assert hyphenate.points("verantwoordelijkheden", "nl-BE") == ()
    assert spelt("développement", "fr-CA") == "dé-ve-lop-pe-ment"
    assert spelt("l'organisation", "fr-FR") == "l'or-ga-ni-sa-tion"  # never at the apostrophe
    assert spelt("responsabilité", "fr-CH", right_min=3) == "res-pon-sa-bi-lité"
    assert hyphenate.points("gouvernement", "fr-BE") == ()
    assert hyphenate.points("infrastructure", None) == ()


def test_the_moby_list_is_read_lazily_and_holds_only_marked_words():
    """``patterns/en-moby.txt.gz``: one word per line, ``=`` at its points (never a hyphen),
    167,826 words converted from Moby's 187,175 entries (``SOURCES.txt``)."""
    import gzip

    from docx2svg import hyphenate

    text = gzip.decompress((hyphenate.PATTERNS / "en-moby.txt.gz").read_bytes()).decode("utf-8")
    lines = text.splitlines()
    assert len(lines) == 167826
    assert hyphenate.WORD_MARK == "=" and "-" not in text and not reader.FORBIDDEN.search(text)
    assert all(line.replace("=", "").isalpha() for line in lines)
    assert hyphenate.word_list("en-moby.txt.gz") is hyphenate.word_list("en-moby.txt.gz")


def _document(settings: str, paragraph_props: str = "", text: str = "infrastructure", lang: str = "en-US") -> bytes:
    import probe_docx as w

    body = (f"<w:p><w:pPr>{paragraph_props}</w:pPr><w:r><w:rPr><w:lang w:val=\"{lang}\"/></w:rPr>"
            f"<w:t>{text}</w:t></w:r></w:p>")
    part = ("word/settings.xml", reader.probe.SETTINGS_CT, reader.probe.SETTINGS_REL,
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="{w.W_NS}">'
            f"{settings}</w:settings>")
    return w.package(body, extra_parts=(part,))


def test_settings_are_read():
    from docx2svg import parse_package

    assert parse_package(_document("")).hyphenation is None
    assert parse_package(_document('<w:autoHyphenation w:val="0"/>')).hyphenation is None
    found = parse_package(_document('<w:autoHyphenation/><w:consecutiveHyphenLimit w:val="2"/>'
                                    '<w:hyphenationZone w:val="720"/><w:doNotHyphenateCaps/>')).hyphenation
    assert (found.zone, found.limit, found.no_caps) == (720, 2, True)


def test_no_hyphens_without_the_setting_or_in_a_suppressed_paragraph():
    from docx2svg import linebreak, parse_package
    from docx2svg.measure import TableAdvances

    def autos(data):
        document = parse_package(data)
        items = linebreak.pieces(document, document.paragraphs[0], TableAdvances())
        return [piece.source for piece in items if piece.auto]

    assert autos(_document("")) == []
    assert autos(_document("<w:autoHyphenation/>")) == [2, 5, 10]
    assert autos(_document("<w:autoHyphenation/>", "<w:suppressAutoHyphens/>")) == []
    assert autos(_document("<w:autoHyphenation/><w:doNotHyphenateCaps/>", text="INFRASTRUCTURE")) == []
    assert autos(_document("<w:autoHyphenation/>", text="INFRASTRUCTURE")) == [2, 5, 10]


def test_french_apostrophes():
    """The typewriter apostrophe is part of a French word; the typographic one joins its
    parts as a hyphen does: below mode 15 only the first part is hyphenated, in mode 15
    each part (``french``)."""
    from docx2svg import linebreak, parse_package
    from docx2svg.measure import TableAdvances

    def autos(text, compat=""):
        document = parse_package(_document("<w:autoHyphenation/>" + compat, text=text, lang="fr-FR"))
        items = linebreak.pieces(document, document.paragraphs[0], TableAdvances())
        return [piece.source for piece in items if piece.auto]

    mode15 = ('<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>')
    assert autos("l'organisation") == [4, 6, 8, 10]
    assert autos("l\u2019organisation") == []
    assert autos("l\u2019organisation", mode15) == [4, 6, 8, 10]
    assert autos("aujourd\u2019hui") == [2]


def test_what_moves_when_a_page_would_end_in_an_automatic_hyphen():
    """[MS-DOCX] 2.3.7 and 2.3.8, as the ``bottom`` documents measure them."""
    from docx2svg.model import Document
    from docx2svg.paginate import hyphen_at_page_end

    def moves(mode, **settings):
        return hyphen_at_page_end(Document(compatibility_mode=mode, compat_settings=settings))

    assert moves(None) is None and moves(14) is None
    assert moves(15) == "line"
    assert moves(15, useWord2013TrackBottomHyphenation="1") == "line"
    assert moves(15, useWord2013TrackBottomHyphenation="0") == "word"
    assert moves(15, useWord2013TrackBottomHyphenation="0", allowHyphenationAtTrackBottom="1") is None
