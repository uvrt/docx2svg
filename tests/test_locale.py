"""The conversion's locale: ``ConvertOptions(locale=...)``.

Word takes a few things from the machine it runs on rather than from the document: the
character a decimal tab stop aligns on is the *system's* decimal separator, not the
document language's (measured, ``tools/make_render_probe.py``, ``tabs``), and a chart's
numbers are printed in its separators.  docx2svg never reads them from its own machine:
the caller says which locale a conversion is in, and without one it is en-US.

The faces' numbers come from the render probe's recording, so these run on any machine.
"""

from __future__ import annotations

import json
import pickle
import sys
from fractions import Fraction
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_render_probe  # noqa: E402
import probe_docx  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402
import wml  # noqa: E402

from docx2svg import ConvertOptions, convert_docx_to_layout, convert_docx_to_svg  # noqa: E402
from docx2svg.vertical import PX_PER_PT  # noqa: E402
from ooxml_common.locales import EN_US, Locale  # noqa: E402

DATA = json.loads(read_render.OBSERVATIONS.read_text(encoding="utf-8"))

#: A caller-built locale whose decimal separator is a comma, as the measuring machine's is.
COMMA = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy", language="nl-NL",
               list_separator=";")

#: The decimal stop of the documents below, twips from the column's left edge.
STOP = 4500
TEXTS = ("12.345", "1234.5", "1234,5", "7")


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


def _document() -> bytes:
    """One paragraph per text, each after a decimal stop, in an ``en-GB`` document."""
    runs = lambda text: wml.run("Lead") + wml.run("\t") + wml.run(text)  # noqa: E731
    body = "".join(make_render_probe._p(runs(text), tabs=f'<w:tab w:val="decimal" w:pos="{STOP}"/>')
                   for text in TEXTS)
    styles = wml.styles_part([wml.style("paragraph", "Normal", default=True)],
                             run_defaults={"rFonts": make_render_probe.fonts("Calibri"), "sz": 22, "szCs": 22})
    return probe_docx.package(body, styles=styles, extra_parts=(wml.settings_part({"val": "en-GB"}),))


def _after_tab(fonts, locale: Locale | None) -> dict[str, tuple[list[str], list[Fraction], Fraction, Fraction]]:
    """Each text's glyphs after the tab: its characters, their pen x and the span's end,
    and the stop -- device px from the page's left edge."""
    options = render_record.options(fonts) if locale is None else render_record.options(fonts, locale=locale)
    layout = convert_docx_to_layout(_document(), options)
    page = layout.pages[0]
    stop = page.column[0] + Fraction(STOP, 20) * PX_PER_PT
    out = {}
    for line in page.lines:
        spans = [span for span in line.spans if span.kind == "text"]
        text = "".join("".join(span.chars) for span in spans)
        assert text.startswith("Lead"), text
        chars = [c for span in spans for c in span.chars][4:]
        xs = [x for span in spans for x in span.xs][4:]
        out[text[4:]] = (chars, xs, spans[-1].end, stop)
    assert sorted(out) == sorted(TEXTS)
    return out


def test_the_default_is_en_us_and_never_the_machines():
    assert ConvertOptions().locale is EN_US
    assert EN_US.decimal_separator == "."


def test_the_options_carry_the_locale_into_a_worker_process():
    options = pickle.loads(pickle.dumps(ConvertOptions(locale=COMMA)))
    assert options.locale == COMMA


def test_a_decimal_stop_aligns_on_the_locales_separator(fonts):
    for locale, separator in ((None, "."), (EN_US, "."), (COMMA, ",")):
        for text, (chars, xs, end, stop) in _after_tab(fonts, locale).items():
            if separator in text:
                # The separator's pen x is the stop.
                assert xs[chars.index(separator)] == stop, (locale, text)
            else:
                # No separator: drawn as after a right stop, ending at it.
                assert end == stop, (locale, text)


def test_a_comma_locale_draws_what_word_drew_on_the_measuring_machine(fonts):
    """``12.345`` and ``1234.5`` after a decimal stop in an ``en-GB`` document end at the
    stop on a machine whose separator is a comma (``make_render_probe.py``, ``tabs``: the
    render probe's recording is held to this in tests/test_render.py)."""
    drawn = _after_tab(fonts, COMMA)
    assert drawn["12.345"][2] == drawn["12.345"][3]
    assert drawn["1234.5"][2] == drawn["1234.5"][3]


def test_the_locale_changes_nothing_but_the_decimal_stops(fonts):
    """A document with no decimal stop and no chart draws the same in every locale."""
    data = make_render_probe.build("none")
    default = convert_docx_to_svg(data, render_record.options(fonts, pages=[1, 2]))
    assert convert_docx_to_svg(data, render_record.options(fonts, pages=[1, 2], locale=COMMA)) == default


def test_a_chart_is_built_in_the_locale(monkeypatch):
    """The chart layout prints a chart's numbers in the locale's separators
    (``ChartBuilder(locale=...)``, ooxml-common's own tests)."""
    import make_chart_text_probe as chart_probe
    import read_chart_text_probe as chart_reader

    from docx2svg import chart

    seen = []

    class Builder(chart.chart_layout.ChartBuilder):
        def __init__(self, *args, **kwargs):
            seen.append(kwargs.get("locale"))
            super().__init__(*args, **kwargs)

    monkeypatch.setattr(chart.chart_layout, "ChartBuilder", Builder)
    monkeypatch.setattr(chart_probe, "CASES", chart_probe.CASES[:1])
    recorded = json.loads(chart_reader.OBSERVATIONS.read_text(encoding="utf-8"))
    chart_fonts = render_record.RecordedFonts(recorded["faces"])
    data = chart_probe.build("15")
    convert_docx_to_svg(data, render_record.options(chart_fonts))
    convert_docx_to_svg(data, render_record.options(chart_fonts, locale=COMMA))
    assert seen == [EN_US, COMMA]
