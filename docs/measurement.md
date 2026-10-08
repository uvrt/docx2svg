# How it is measured

Moved from the README. [`ROADMAP.md`](../ROADMAP.md) is the full record.

## The oracle came first

What came first, and still gates everything, is what the sibling
[`pptx2svg`](https://github.com/uvrt/pptx2svg) also established before writing a line of
output code:

- **An oracle.** Microsoft Word, driven by AppleScript, exporting a document to PDF whose
  text is *vector* — so every glyph's position is readable as a number rather than as a
  pixel. `tools/word_export_pdf.applescript`. A document with revisions or comments is
  exported in Word's *No Markup* view, since Word's default prints them in balloons
  (ROADMAP.md, "Revisions — measured").
- **A fixture that is ours.** `tests/fixtures/layout-sweep.docx` is written by
  `tools/make_layout_sweep.py` and contains no third-party content. Word opens it and
  exports it, which is the only acceptance test that means anything for generated OOXML.
- **Enough of a reader to close the loop.** `src/docx2svg/` parses sections, paragraphs,
  indents, tab stops, styles, the theme's fonts, numbering and table structure — only
  what a test compares against Word's own output — and `docx2svg.resolve` resolves the
  style cascade, saying for every value which level it came from.
- **A measurement.** Below.

[`ROADMAP.md`](../ROADMAP.md) is the substantial document here. It says where the difficulty
actually is, what has been measured, and what has been refuted.

## The loop, and what it currently measures

```
tools/make_layout_sweep.py   ->  layout-sweep.docx      (ours, generated, committed)
tools/word_export_pdf.applescript -> layout-sweep.pdf   (Word decides the layout)
src/docx2svg/parse           ->  what the file asked for
tools/read_layout_sweep.py   ->  the residual between the two
```

```bash
python tools/make_layout_sweep.py
python tools/read_layout_sweep.py --export      # runs Word, then measures
python -m pytest -q                             # oracle tests skip without Word
```

Today that prints:

| Quantity | Authored | Measured out of Word's PDF | Residual |
| --- | --- | --- | --- |
| Indent ladder, 5 steps | 0 / 18 / 36 / 54 / 72 pt | same | **8e-6 pt** |
| Tab stops, 3 | 72 / 144 / 216 pt | same | **8e-6 pt** |
| A4 page box | 595.3 × 841.9 pt | 595.2 × 841.92 pt | see below |
| Baselines on the 1/300-inch grid | — | 70 of 70 | < 0.01 px |

8e-6 pt is float32 noise in pdfium's coordinates, on a 600 pt page. A twip — the smallest
thing this format can express — is 0.05 pt, six thousand times larger. So these are exact.

The residuals are exact because every one of them is a **difference** between two glyph
ink boxes. That is the whole design of the fixture: a difference cancels the glyph's left
side bearing, the face's vertical metrics and Word's page-box rounding, so the comparison
needs no font file — which matters, because this project may not ship one.

## Two things measured on day one that were not the plan

**Word's exported page box is not the page size.** A4's `w:pgSz w:w="11906"` is 595.3 pt,
and the PDF says 595.2. Both exported extents are whole 1/300-inch device pixels — 2480
and 3508, which are exactly A4 at 300 dpi. The fixture's second section exists to falsify
that: it asks for 10000 × 13000 twips, which is *exactly* 500 × 650 pt, so a page box that
were simply the authored size would read 500 × 650. It reads **499.92 × 649.92**. The law
survived the test built to kill it. `docx2svg.units.device_page_extent_pt`.

**Word snaps every baseline to that same grid.** Seventy single-line paragraphs, seventy
baselines, every one a whole device pixel. The gaps are 56 px with three of 55 — pure
accumulated rounding of an advance of 55.9489 px, which is Calibri's
`(ascender − descender + lineGap) / unitsPerEm × 11 pt` exactly. The ratio is settled; the
accumulator is not, and ROADMAP.md says which nine baselines the obvious model gets wrong.

## Why a sibling of pptx2svg rather than a feature of it

DOCX and PPTX share their container (OPC), their vector graphics (DrawingML, which arrives
inside `w:drawing`), and all of their text measurement. They do not share a layout model,
and that is the whole of the difference: **PPTX positions boxes absolutely; DOCX flows text
and paginates.** A slide says where its text goes. A document does not say where its lines
break or where its pages end — Word computes both, and a break in the wrong place
invalidates everything after it.

What the two share is not assumed here. It was measured from `pptx2svg`'s import graph
(ROADMAP.md, "Phase 1"), and the shareable half has been extracted, with its history,
into [`ooxml-common`](https://github.com/uvrt/ooxml-common): the OPC container, units,
DrawingML's reader and renderers, the font modules and the measured text metrics. This
project measures text with that package, not with a PowerPoint renderer, and draws its
DrawingML shapes -- fills, gradients, patterns, outlines, arrowheads, effects, WordArt's
text fill -- with that package's renderers, under the rules where Word was measured to
draw differently from PowerPoint (ROADMAP.md, "Floating drawings -- measured", F.12), and
its charts and SmartArt with that package's chart layout and scene renderers, under the
chart and text rules measured on Word (F.20, F.21).

