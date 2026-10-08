"""What ``w:beforeAutospacing`` / ``w:afterAutospacing`` do to a paragraph's spacing, without Word.

``tests/fixtures/autospacing-observations.json`` holds the lines Word drew for every
document of ``tools/make_autospacing_probe.py`` (recorded by
``tools/read_autospacing_probe.py --record``); the documents are regenerated from the
generator.  Every baseline goes through ``baselines.predict``, so this holds the cascade
(``docx2svg.resolve``: 14 pt whatever is stated) and ``docx2svg.vertical.autospace_kept``
(nothing between two list items or before the document's first paragraph) to Word.
ROADMAP.md, "Autospacing -- measured", has the tables.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import baselines
import make_autospacing_probe as probe
import probe_documents
import read_autospacing_probe as reader

from docx2svg import parse_package
from docx2svg.resolve import resolve_paragraph

DATA = json.loads((Path(__file__).parent / "fixtures" / "autospacing-observations.json").read_text())
SCORED = reader.offline(DATA)

#: (exact, scored) per case, in each of the three settings that use autospacing.
EXPECTED = {
    "document start": (2, 2), "stated": (240, 240), "neighbour": (35, 35), "consecutive": (5, 5),
    "list": (5, 5), "list edges": (5, 5), "list stated": (5, 5), "collapse": (8, 8),
    "contextual": (5, 5), "section start": (3, 3), "manual break": (2, 2), "pageBreakBefore": (2, 2),
}
#: Under ``w:doNotUseHTMLParagraphAutoSpacing`` the stated values hold, and the lines the
#: model misses are the ones where two paragraphs' spacings **add** (a finding not
#: adopted; see the last test).
EXPECTED_NO_HTML = EXPECTED | {"consecutive": (2, 5), "list": (2, 5), "list stated": (2, 5),
                               "collapse": (2, 8)}


def test_the_recording_is_of_these_documents():
    assert set(DATA["documents"]) == {p.name for p in probe.PROBES}
    for p, results in SCORED:
        assert all(r.status in ("exact", "miss") for r in results), p.name


@pytest.mark.parametrize("setting", sorted(probe.SETTINGS))
def test_autospacing_by_case(setting):
    table = reader.tabulate(SCORED)
    want = EXPECTED_NO_HTML if setting == "15-noHTML" else EXPECTED
    assert {case: tuple(table[(setting, case)]) for case in want} == want


def test_the_cascade_resolves_an_autospaced_side_to_14_pt_and_says_why():
    document = parse_package(probe.build(probe.Probe("15")))
    paragraph = next(p for p in document.paragraphs if p.text == "Before 1000-40-auto240")
    resolved = resolve_paragraph(document, paragraph)
    assert resolved["spacing.before"] == 280
    assert "14 pt" in resolved.origin("spacing.before").detail
    no_html = parse_package(probe.build(probe.Probe("15-noHTML")))
    paragraph = next(p for p in no_html.paragraphs if p.text == "Before 1000-40-auto240")
    assert resolve_paragraph(no_html, paragraph)["spacing.before"] == 1000


def _total(scored, settings=("none", "14", "15")):
    table = reader.tabulate(scored)
    return sum(v[0] for (s, _), v in table.items() if s in settings)


def test_the_stated_value_and_the_neighbour_free_rule_are_refuted(monkeypatch):
    """Of the three settings' 951 lines, reading the stated spacing (the model before the
    probe) fits 275; 14 pt everywhere, between list items and before the document's
    first paragraph too, fits 927; the model fits all of them."""
    import docx2svg.resolve.cascade as cascade

    assert _total(SCORED) == 951
    monkeypatch.setattr(baselines, "autospace_kept", lambda **_: True)
    assert _total(reader.offline(DATA)) == 927
    monkeypatch.setattr(cascade, "_autospacing", lambda *_: None)
    assert _total(reader.offline(DATA)) == 275


def test_do_not_use_html_paragraph_auto_spacing_makes_all_spacing_add():
    """Not adopted, recorded: under this compat option every paragraph's space after and
    the next one's space before **add** -- plain paragraphs too -- and every line of the
    document is exact with ``collapse="sum"`` (302 / 317 collapsing); in the other
    settings adding misses 33."""
    data = DATA["documents"]
    metrics = probe_documents.recorded_metrics(DATA["faces"])
    for setting, want in (("15-noHTML", (317, 317)), ("15", (284, 317))):
        p = probe.Probe(setting)
        docx = probe.build(p)
        results = baselines.predict(parse_package(docx), probe_documents.expand(data[p.name]), metrics,
                                    package=docx, collapse="sum")
        assert baselines.summary(results)[:2] == want
