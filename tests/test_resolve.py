"""The cascade against what Word drew, and its rules one at a time.

``tests/fixtures/style-observations.json`` holds every glyph Word 16.106 drew for the
probes of ``tools/make_style_probe.py`` (face, drawn size, pen x), recorded by
``tools/read_style_probe.py --record``.  The probes are rebuilt here from the committed
generator -- byte for byte what was exported -- parsed, resolved, and compared glyph by
glyph, so the cascade is held to Word without Word.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parent.parent / "tools"
sys.path.insert(0, str(TOOLS))

import make_style_probe as probe  # noqa: E402
import read_style_probe as reader  # noqa: E402

from docx2svg import parse_package  # noqa: E402
from docx2svg.model import Document, FontScheme, ThemeFonts  # noqa: E402
from docx2svg.resolve import character_format, resolve_paragraph, resolve_run, slot_for, theme_face  # noqa: E402

OBSERVATIONS = json.loads((Path(__file__).parent / "fixtures" / "style-observations.json").read_text(encoding="utf-8"))
PROBES = {p.name: p for p in probe.probes()}


@pytest.mark.parametrize("name", sorted(PROBES))
def test_every_case_resolves_to_what_word_drew(name):
    glyphs, misses, _, report = reader.check_probe(PROBES[name], OBSERVATIONS["probes"][name])
    assert misses == 0, "\n".join(report)
    assert glyphs > 0


def test_the_probes_cover_what_the_brief_asked():
    total = sum(reader.check_probe(p, OBSERVATIONS["probes"][n])[0] for n, p in PROBES.items())
    assert total == 1770
    assert sum(len(p.cases) for p in PROBES.values()) == 172


def _case(name: str, label: str):
    document = parse_package(PROBES[name].build())
    paragraph = next(p for p in reader.paragraphs(document.body) if p.text.startswith(label))
    return document, paragraph


def _bold(name: str, label: str) -> bool:
    document, paragraph = _case(name, label)
    return bool(resolve_run(document, paragraph, paragraph.runs[-1]).get("b"))


def test_toggles_xor_across_levels_but_override_within_a_chain():
    assert _bold("style-toggles", "t01") and _bold("style-toggles", "t02")
    assert not _bold("style-toggles", "t03")  # paragraph bold + character bold
    assert _bold("style-toggles", "t08")  # bold based on bold: override, not XOR
    assert not _bold("style-toggles", "t09")  # b=0 based on bold
    assert _bold("style-toggles", "t11")  # character b=0 flips nothing
    assert _bold("style-toggles", "t29")  # table + paragraph + character
    assert not _bold("style-toggles", "t05") and _bold("style-toggles", "t04")  # direct


def test_toggles_flip_relative_to_the_document_default():
    assert _bold("style-defaults-bold", "g07")  # docDefaults, paragraph and character all bold
    assert not _bold("style-defaults-bold", "g10")  # a style saying false turns it off
    assert not _bold("style-normal-bold", "g07")  # Normal is a style, so it XORs


def test_the_default_character_style_contributes_nothing():
    assert not _bold("style-charstyle-bold", "g01")
    assert not _bold("style-charstyle-bold", "g09")  # not even through basedOn


def test_every_value_says_where_it_came_from():
    document, paragraph = _case("style-toggles", "t03")
    resolved = resolve_run(document, paragraph, paragraph.runs[-1])
    text = resolved.explain(["b", "sz"])
    assert "paragraph style 'PB' [True, flips]" in text
    assert "character style 'CB' [True, flips]" in text
    assert "sz = 22  <- docDefaults" in text
    document, paragraph = _case("style-toggles", "t10")
    origin = resolve_run(document, paragraph, paragraph.runs[-1]).origins["b"][-1]
    assert (origin.style_id, origin.declared_in) == ("PBkid", "PB")


def test_direct_numbering_outranks_the_paragraph_style_indent():
    document, paragraph = _case("style-precedence", "p02")
    assert resolve_paragraph(document, paragraph).get("ind.left") == 720
    document, paragraph = _case("style-precedence", "p03")
    resolved = resolve_paragraph(document, paragraph)
    assert (resolved.get("ind.left"), resolved.get("ind.hanging")) == (2160, 360)


def test_theme_fonts_go_through_theme_font_lang_for_east_asian_and_complex_script():
    scheme = FontScheme(
        major=ThemeFonts("Trebuchet MS", "", "Verdana", {"Hebr": "Courier New"}),
        minor=ThemeFonts("Georgia", "", "Courier New", {"Hebr": "Arial", "Hans": "DengXian"}),
    )
    with_lang = Document(font_scheme=scheme, theme_font_lang={"eastAsia": "zh-CN", "bidi": "he-IL"})
    assert theme_face("minorHAnsi", with_lang)[0] == "Georgia"
    assert theme_face("minorBidi", with_lang)[0] == "Arial"
    assert theme_face("majorBidi", with_lang)[0] == "Courier New"
    assert theme_face("minorEastAsia", with_lang)[0] == "DengXian"
    unlisted = Document(font_scheme=scheme, theme_font_lang={"bidi": "ar-SA"})
    assert theme_face("minorBidi", unlisted)[0] == "Courier New"  # a:cs
    assert theme_face("minorBidi", Document(font_scheme=scheme))[0] is None  # Word's fallback
    assert theme_face("majorHAnsi", Document())[0] == "Calibri Light"  # built-in theme


def test_complex_script_is_decided_by_the_run_not_the_character():
    document, paragraph = _case("style-fonts-langs", "f09")  # Hebrew, no w:cs / w:rtl
    resolved = resolve_run(document, paragraph, paragraph.runs[-1])
    fmt = character_format(resolved, "ש", document)
    assert (fmt.slot, fmt.half_points) == ("hAnsi", 22)
    document, paragraph = _case("style-fonts-langs", "f20")  # Hebrew with w:rtl
    resolved = resolve_run(document, paragraph, paragraph.runs[-1])
    fmt = character_format(resolved, "ש", document)
    assert (fmt.slot, fmt.face, fmt.half_points) == ("cs", "Arial", 40)


def test_slots_by_character():
    assert slot_for("a") == "ascii"
    assert slot_for("é") == "hAnsi"
    assert slot_for("é", hint="eastAsia", east_asian_language="zh-CN") == "eastAsia"
    assert slot_for("é", hint="eastAsia", east_asian_language="ja-JP") == "hAnsi"
    assert slot_for("§", hint="eastAsia") == "eastAsia"
    assert slot_for("“", hint="eastAsia") == "eastAsia"
    assert slot_for("中") == "eastAsia"
    assert slot_for("a", complex_script=True) == "cs"


def test_table_conditional_formats_apply_by_cell_and_override_each_other():
    """sample-simple.docx: the first row and the first column are bold and in the major
    font through ``w:tblStylePr``; the corner cell, in both, is drawn bold (not XOR)."""
    from docx2svg.model import Table

    path = Path(__file__).parent / "fixtures" / "samplelib" / "sample-simple.docx"
    document = parse_package(path.read_bytes())
    table = next(block for block in document.body if isinstance(block, Table))
    corner = table.rows[0][0][0]
    body = table.rows[1][1][0]
    assert corner.table_conditions == ("wholeTable", "firstCol", "firstRow", "nwCell")
    resolved = resolve_run(document, corner, corner.runs[0])
    assert resolved.get("b") is True
    assert resolved.origins["b"][0].level == "table style"
    assert character_format(resolved, "P", document).face == "Calibri"
    plain = resolve_run(document, body, body.runs[0])
    assert not plain.get("b") and character_format(plain, "3", document).face == "Cambria"
