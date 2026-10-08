# Changelog

docx2svg has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) records every measurement in detail.

## Unreleased

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
