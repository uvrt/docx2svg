# Changelog

docx2svg has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) records every measurement in detail.

## Unreleased

- **Fonts without Office.** An Office face that is neither installed nor embedded is laid
  out and drawn with its open metric compatible substitute, if that is installed: Carlito
  for Calibri, and Liberation Sans, Serif and Mono for Arial, Times New Roman and Courier
  New. Each pair was measured equal in advances and line metrics against Word's copy.
  Every substitution is reported (`font-substituted`), and none is made when the real face
  is present. An absent Symbol or Wingdings is laid out from the integers recorded from
  Word's copies (`docx2svg.recorded`, `tools/record_symbol_faces.py`), and its bullets are
  drawn as Unicode equivalents. `ConvertOptions.substitute_fonts` (`--no-substitute-fonts`)
  turns substitution off, and `ConvertOptions.font_substitutes` names approximate
  substitutes. The README has a new "Fonts" section.
- **Coverage.** `ConvertOptions.coverage` / `Layout.coverage` (`docx2svg.coverage`) say
  how much was laid out: the pages, the blocks laid out and skipped, the first stop with
  its reason and path, header, footer and text-box stops, and the faces substituted or
  missing. The CLI prints its summary.
- CI installs Carlito, Caladea and Liberation on Linux and macOS, so the layout code runs
  there past the first paragraph. `tests/test_substitutes.py` checks Carlito's layout of
  `layout-sweep.docx` glyph for glyph against Word's recorded Calibri numbers.

- Chart labels as Word draws them (ooxml-common 0.7): tick, category and legend text in
  its `c:txPr` colour (`tx1` at 65%, as Office writes it), data labels in their own number
  format (`€41.2m`), and a radar's category labels 4% of the radius off their vertex.

## 0.1.0 -- 2026-10-08 (initial public release)

Developed 2026-09-24 to 2026-10-07, phase by phase against Microsoft Word's PDF export:

- **Phase 0:** the Word oracle (AppleScript export to vector PDF, measured with
  pypdfium2), a generated fixture that is ours, and a loop that closes.
- **Phase 1:** the format-neutral half shared with pptx2svg extracted into
  [`ooxml-common`](https://github.com/uvrt/ooxml-common).
- **Phases 2-4:** the vertical model, line breaking to an exact inclusive budget,
  hyphenation (English, Dutch, German, French, from openly licensed patterns), and
  pagination with keeps, footnotes, endnotes, tables, text columns and floating drawings.
- **Phase 5:** SVG and PNG output, every glyph checked against Word's PDF; headers,
  footers and computed fields; charts and SmartArt; tracked changes in Word's final view.
- **API:** `convert_docx` (layout and SVGs from one layout), `data-docx-path` element
  paths with `docx2svg.paths.locate`, `Page.info`, list counters without a document.
