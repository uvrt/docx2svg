"""Which kern pairs Word charges at a line's edge: ``tools/make_wrap_kern_probe.py``,
recorded by ``tools/read_wrap_kern_probe.py`` in ``tests/fixtures/wrap-kern-observations.json``.

pptx-agent's wrap-boundary probe set PowerPoint's rule (a static face's legacy ``kern``
table, never ``GPOS``); this is the same question asked of Word, with the same words
(Pass, Fail, Total, AVAWAY, Review), faces (Aptos, Calibri, Arial), sizes (12, 18, 24 pt)
and boundary widths, kerning on (``w:kern``) and off.  The rule is ooxml-common's
``DrawingRules.kerning`` for Word, and what docx2svg's own line breaking charges
(``fonts.Face.kern``, the legacy table).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ooxml_common.drawingml.rules import WORD

DATA = json.loads((Path(__file__).parent / "fixtures" / "wrap-kern-observations.json").read_text(encoding="utf-8"))
CASES = [case for case in DATA["cases"] if case["tag"] != "width"]


def _width(case: dict, model: str) -> float:
    """The word's width in points: plain advances (``U``), the legacy table's pairs
    (``L``) or ``GPOS``'s (``G``)."""
    units = DATA["units"][f"{case['face']}|{case['word']}"]
    pairs = {"U": 0, "L": sum(units["legacy"]), "G": sum(units["gpos"])}[model]
    return (sum(units["advances"]) + pairs) * case["size"] / units["units_per_em"]


def _wrong(model_on: str, model_off: str) -> list[dict]:
    """Cases whose verdict -- one line when the word fits the text width -- the models
    get wrong: ``model_on`` with ``w:kern``, ``model_off`` without."""
    out = []
    for case in CASES:
        width = _width(case, model_on if case["kern"] else model_off)
        if (len(case["lines"]) == 1) != (width <= case["twips"] / 20):
            out.append(case)
    return out


def test_the_cases_cover_both_switches_and_every_word():
    assert len(CASES) == 648
    assert {case["kern"] for case in CASES} == {2, None}
    assert {(case["face"], case["size"], case["word"]) for case in CASES} == {
        (face, size, word) for face in ("Aptos", "Calibri", "Arial") for size in (12, 18, 24)
        for word in ("Pass", "Fail", "Total", "AVAWAY", "Review")}


def test_word_kerns_with_the_legacy_table_where_w_kern_asks_and_not_at_all_otherwise():
    """648 of 648 verdicts: the legacy table's pairs with ``w:kern``, none without."""
    assert _wrong("L", "U") == []


def test_the_gpos_pairs_are_refuted():
    """Aptos's ``ss`` (-33/2048 em) is in ``GPOS`` alone: "Pass" fits a line ``GPOS``
    says it does not need -- 7 of the 324 kerned cases."""
    wrong = _wrong("G", "U")
    assert len(wrong) == 7
    assert {case["face"] for case in wrong} == {"Aptos"}
    assert {case["word"] for case in wrong} == {"Pass"}
    # Nor does Word kern by default: the kerned widths get the unkerned cases wrong.
    assert len(_wrong("L", "L")) > 100


def test_each_width_word_kept_on_a_line_is_the_model_s_to_a_thousandth_of_a_point():
    for case in CASES:
        if case["width"] is not None:
            assert case["width"] == pytest.approx(_width(case, "L" if case["kern"] else "U"), abs=1e-3), case


def test_a_variable_face_is_not_kerned_from_gpos():
    """Noto Sans JP and STIX Two Text, kerned: PowerPoint charges their ``GPOS`` pairs,
    Word charges none (they have no legacy table)."""
    widths = [case for case in DATA["cases"] if case["tag"] == "width"]
    assert {case["face"] for case in widths} == {"Noto Sans JP", "STIX Two Text"}
    for case in widths:
        assert case["width"] == pytest.approx(_width(case, "U"), abs=2e-3)
        assert abs(case["width"] - _width(case, "G")) > 0.5


def test_the_shared_rule_is_word_s():
    assert (WORD.kerning.static, WORD.kerning.variable) == ("legacy", "none")
