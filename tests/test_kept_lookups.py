"""What a layout finds once and keeps (round 39): numbers written in integers, a glyph's x
made at once, faces and kerning pairs per font object, properties per parsed document.
Each is the same value the work it replaces gives; made-up documents only."""

from __future__ import annotations

import io
import pickle
import random
import zipfile
from fractions import Fraction


import docx2svg
from docx2svg import parse_package
from docx2svg.fonts import InstalledFonts
from docx2svg.layout import LAYOUT_UNIT_PX, px_at, units_px
from docx2svg.resolve import cascade, resolve_paragraph, resolve_run
from docx2svg.svg import number


def _number_before(value) -> str:
    """``number`` as it was: through a Fraction ten thousand times larger."""
    scaled = Fraction(value) * 10000
    whole = (scaled.numerator * 2 + scaled.denominator) // (2 * scaled.denominator)
    sign = "-" if whole < 0 else ""
    whole = abs(whole)
    integer, fraction = divmod(whole, 10000)
    if not fraction:
        return f"{sign}{integer}"
    return f"{sign}{integer}.{fraction:04d}".rstrip("0")


def test_numbers_are_written_as_before():
    rng = random.Random(39)
    values: list = [0, 1, -1, True, False, 0.5, -0.5, 0.00005, -0.00005, Fraction(1, 20000), Fraction(-1, 20000),
                    Fraction(3, 2), Fraction(-3, 2), 12.25, -7.000049999]
    for _ in range(20000):
        k = rng.random()
        if k < 0.3:
            values.append(rng.randint(-10 ** 7, 10 ** 7))
        elif k < 0.8:
            values.append(Fraction(rng.randint(-10 ** 9, 10 ** 9),
                                   rng.choice([1, 2, 3, 24576, 49152, rng.randint(1, 10 ** 6)])))
        else:
            values.append(rng.uniform(-1e4, 1e4))
    assert [number(v) for v in values] == [_number_before(v) for v in values]


def test_a_glyph_s_x_is_the_sum_it_replaces():
    rng = random.Random(7)
    lefts = [0, 300, -12, Fraction(301, 3), Fraction(-5, 7), LAYOUT_UNIT_PX * 99, 2.5]
    units = [0, 1, 4096, -4096, 123456, Fraction(7, 3), Fraction(-1, 1000)]
    for _ in range(2000):
        lefts.append(Fraction(rng.randint(-10 ** 6, 10 ** 6), rng.randint(1, 5000)))
        units.append(rng.randint(-10 ** 7, 10 ** 7))
    for left in lefts:
        for unit in units[:: max(1, len(units) // 40)]:
            made = px_at(left, unit)
            expected = left + units_px(unit)
            assert made == expected and type(made) is type(expected)
            if type(made) is Fraction:
                assert (made.numerator, made.denominator) == (expected.numerator, expected.denominator)


def test_faces_and_kerning_pairs_are_found_once_and_never_pickled():
    fonts = InstalledFonts()
    for family in ("Calibri", "Times New Roman", "No Such Face 39"):
        for bold in (False, True):
            found = fonts.face(family, bold, False)
            assert found is fonts.face(family, bold, False)
            again = InstalledFonts()._find(InstalledFonts()._index(), family, bold, False, recorded=True)
            assert (found is None) == (again is None)
            if found is not None:
                assert (found.source, getattr(found, "offset", None)) == \
                    (again.source, getattr(again, "offset", None))
    face = fonts.face("Calibri")
    for left, right in (("A", "V"), ("T", "o"), ("a", "b")):
        expected = None if face is None else face.kern(left, right)
        assert fonts.kern("Calibri", False, False, left, right) == expected
        assert fonts.kern("Calibri", False, False, left, right) == expected
    assert "No Such Face 39" in fonts.missing or "No Such Face 39" in fonts.substitutions
    state = fonts.__getstate__()
    assert "_faces" not in state and "_kerns" not in state
    copy = pickle.loads(pickle.dumps(fonts))
    assert copy.substitutions == fonts.substitutions and copy.missing == fonts.missing
    assert (copy.face("Calibri") is None) == (face is None)


def test_kerning_pairs_held_are_bounded(monkeypatch):
    from docx2svg import fonts as fonts_module

    monkeypatch.setattr(fonts_module, "KERN_PAIRS", 5)
    fonts = InstalledFonts()
    for k in range(40):
        fonts.kern("Calibri", False, False, chr(65 + k % 26), chr(97 + k))
    assert len(fonts.__dict__["_kerns"][1]) <= 5


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx(body: str, styles: str = "") -> bytes:
    files = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships'
        '+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.'
        'org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "word/document.xml": f'<?xml version="1.0"?><w:document xmlns:w="{_W}"><w:body>{body}</w:body></w:document>',
        "word/styles.xml": f'<?xml version="1.0"?><w:styles xmlns:w="{_W}"><w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" '
        'w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240"/></w:pPr>'
        f'<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong">'
        f'<w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>{styles}</w:styles>',
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in files.items():
            archive.writestr(name, text)
    return buffer.getvalue()


def _paragraph(style: str = "", run: str = "", text: str = "Text") -> str:
    p_pr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{p_pr}<w:r>{f'<w:rPr>{run}</w:rPr>' if run else ''}<w:t>{text}</w:t></w:r></w:p>"


BODY = "".join([
    _paragraph(), _paragraph("Heading1"), _paragraph(run='<w:b/>'), _paragraph(run='<w:sz w:val="22"/>'),
    _paragraph(run='<w:rStyle w:val="Strong"/>'), _paragraph("Heading1", '<w:rStyle w:val="Strong"/>'),
    _paragraph(), _paragraph("Heading1"), _paragraph(run='<w:b/>'), _paragraph(run='<w:b w:val="0"/>'),
])


def test_properties_are_resolved_once_per_input_and_are_those_resolved_alone():
    document = parse_package(_docx(BODY))
    paragraphs = document.paragraphs
    resolved = [(resolve_paragraph(document, p), resolve_run(document, p, p.runs[0])) for p in paragraphs]
    alone = [(cascade._resolve_paragraph(document, p),
              cascade._resolve_run(document, p, p.runs[0].properties.style_id if p.runs[0].properties else None,
                                   p.runs[0].properties.declared if p.runs[0].properties else {}, p.runs[0].toc_link))
             for p in paragraphs]
    for (paragraph, run), (paragraph_alone, run_alone) in zip(resolved, alone):
        assert paragraph.values == paragraph_alone.values and paragraph.origins == paragraph_alone.origins
        assert run.values == run_alone.values and run.origins == run_alone.origins
    # The same inputs: the same object; other ones: another.
    assert resolved[0][0] is resolved[6][0] and resolved[1][0] is resolved[7][0]
    assert resolved[2][1] is resolved[8][1] and resolved[2][1] is not resolved[9][1]
    assert resolved[0][1] is not resolved[3][1]           # a size stated is direct formatting
    assert resolved[4][1] is not resolved[5][1]           # the paragraph's style is an input of a run
    # Parsed again, nothing is kept from before.
    again = parse_package(_docx(BODY))
    assert resolve_paragraph(again, again.paragraphs[0]) is not resolved[0][0]


def test_direct_formatting_that_cannot_be_hashed_is_resolved_every_time():
    document = parse_package(_docx(BODY))
    paragraph = document.paragraphs[2]
    run = paragraph.runs[0]
    run.properties.declared["unhashable.test"] = ["a list"]
    first, second = resolve_run(document, paragraph, run), resolve_run(document, paragraph, run)
    assert first is not second and first.values == second.values
    assert first.get("unhashable.test") == ["a list"]


def test_a_layout_is_the_same_with_what_it_keeps_and_without(monkeypatch):
    data = _docx(BODY * 6)
    kept = docx2svg.convert_docx_to_svg(data)
    monkeypatch.setattr(cascade, "_declared_key", lambda declared: None)
    monkeypatch.setattr(InstalledFonts, "face", lambda self, family, bold=False, italic=False:
                        self._find(self._index(), family, bold, italic, recorded=True))
    assert docx2svg.convert_docx_to_svg(data) == kept
