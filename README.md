# docx2svg

[![CI](https://github.com/uvrt/docx2svg/actions/workflows/ci.yml/badge.svg)](https://github.com/uvrt/docx2svg/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Render Word (`.docx`) documents to **SVG**, page by page, in pure Python -- laid out as
Word lays them out, measured glyph by glyph against Word's own PDF export.

## Install

Not on PyPI yet. Python 3.10+; standard library only at runtime, plus
[`ooxml-common`](https://github.com/uvrt/ooxml-common), which is too.

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main"
pip install "docx2svg @ git+https://github.com/uvrt/docx2svg@main"
pip install "docx2svg[png] @ git+https://github.com/uvrt/docx2svg@main"   # PNG output (resvg-py)
```

## Example

```python
from docx2svg import convert_docx_to_svg, convert_docx_to_png, ConvertOptions

options = ConvertOptions()
svgs = convert_docx_to_svg("report.docx", options)   # one SVG string per page
pngs = convert_docx_to_png("report.docx")            # needs `pip install docx2svg[png]`
for warning in options.warnings:                     # everything not drawn faithfully
    print(warning)
```

```bash
docx2svg report.docx -o out/ -f both --strict
```

Every drawn element carries `data-docx-path` back to the XML it came from, and
`convert_docx_to_layout` returns the pages, lines and spans themselves:
[docs/api.md](docs/api.md).

## What it draws

- **Drawn:** text in its resolved face, size, weight, slant and colour at Word's pen
  position; underline, strikethrough, highlight, shading, tab leaders, list labels, line
  numbers, paragraph borders; inline pictures; tables; text columns; footnotes and
  endnotes; headers and footers with `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `PAGEREF`, `REF`
  and `SEQ` computed; floating drawings, shapes and text boxes; charts and SmartArt;
  content controls; tracked changes as Word's final (*No Markup*) view.
- **Not drawn yet:** a chart's 3-D scene (drawn flat), SmartArt with no cached drawing,
  a picture fill, a header's drawing the body wraps around, and some floating and note
  cases in multi-column sections. Each is reported with a stable warning code, never
  dropped in silence.
- **Fonts** are named in the SVG, not embedded. The layout reads the faces Word uses
  where they are installed (Office's cloud-font cache too).

Full list and per-area status: [docs/supported.md](docs/supported.md).

## Status

Pre-alpha (0.1.0). Line breaking, pagination, columns, revisions and output are
measured against Word; see the status table in [docs/supported.md](docs/supported.md)
and the full record in [`ROADMAP.md`](ROADMAP.md), which says what has been measured,
what has been refuted and where the difficulty is.

## Documentation

- [docs/supported.md](docs/supported.md) -- what is drawn, what is not, status by area
- [docs/api.md](docs/api.md) -- `convert_docx`, element paths, page info, list counters
- [docs/measurement.md](docs/measurement.md) -- the Word oracle, the measured loop, and why
  this is a sibling of pptx2svg
- [docs/development.md](docs/development.md) -- running the tests, the fidelity harness,
  constraints (standard library only; no Microsoft font file in the repository)
- [ROADMAP.md](ROADMAP.md) -- the measurements, phase by phase
- [CHANGELOG.md](CHANGELOG.md), [CONTRIBUTING.md](CONTRIBUTING.md)

## Family

- [pptx2svg](https://github.com/uvrt/pptx2svg) -- renders PowerPoint (`.pptx`) slides to SVG and PNG.
- [docx2svg](https://github.com/uvrt/docx2svg) (this repo) -- renders Word (`.docx`) documents to SVG, page by page.
- [ooxml-common](https://github.com/uvrt/ooxml-common) -- the format-neutral reading, DrawingML, fonts and text metrics both renderers share.
- [ooxml-edit](https://github.com/uvrt/ooxml-edit) -- lossless, undoable editing of OOXML packages, shared by both agent layers.
- [pptx-agent](https://github.com/uvrt/pptx-agent) -- an AI-editable PowerPoint layer: inspect, edit, re-render.
- [docx-agent](https://github.com/uvrt/docx-agent) -- an AI-editable Word layer: inspect, edit (optionally as tracked changes), re-render.

## License

MIT, see [LICENSE](LICENSE). The third-party fixtures under `tests/fixtures/samplelib/`
and `tests/fixtures/wordto/` keep their own (permissive) terms, recorded in their
`PROVENANCE.md`. No font file is included.
