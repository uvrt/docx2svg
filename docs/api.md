# API


```python
from docx2svg import ConvertOptions, convert_docx, convert_docx_to_layout, convert_docx_to_svg
from docx2svg.paths import locate, path_part, resolve_path

svgs = convert_docx_to_svg("report.docx")       # one SVG string per page
layout = convert_docx_to_layout("report.docx")  # docx2svg.layout.Layout: pages, lines, spans, rules, floats

result = convert_docx("report.docx", ConvertOptions(pages=[2, 3]))
result.layout          # every page laid out, as convert_docx_to_layout returns it
result.svgs            # the selected pages' SVGs, byte for byte convert_docx_to_svg's
result.page_numbers    # [2, 3]: the page each SVG is of
```

- **Coverage.** Every conversion sets `options.coverage` (and `layout.coverage`), a
  `docx2svg.coverage.Coverage` that sums up the warnings: `complete`; `pages` and
  `estimated_pages` (Word's count from `docProps/app.xml` when the layout stopped, else
  `None`; `estimate_source` says which); `blocks`, `blocks_laid_out` and
  `blocks_skipped`; `stop` (the body's first stop: `code`, `reason`, `message`, `page`
  and the element `path`); `story_stops`; `substituted_fonts` (`family`, `substitute`,
  `metric_compatible`) and `missing_fonts`; `approximations` (each like `stop`: a place
  laid out by an unmeasured rule rather than stopped, `layout-approximate:<reason>`); and
  `status`: `complete`, `approximate` (complete, with approximations) or `partial`.
  `as_dict()` gives it as JSON-ready data and `summary()` as one line. Faces absent from
  the machine are laid out with their open metric compatible substitutes, and each
  substitution is reported (`font-substituted`).
  `ConvertOptions(substitute_fonts=False)` turns this off, and `font_substitutes={...}`
  names more substitutes, which are reported as approximate. See the README, "Fonts".

  ```python
  options = ConvertOptions()
  layout = convert_docx_to_layout("report.docx", options)
  if not options.coverage.complete:
      print(options.coverage.summary())   # partial: 1 of ~12 page(s), 4 of 230 block(s) laid out; ...
  ```
- **`convert_docx(source, options)`** lays the document out once and returns a
  `Conversion`: the layout and the selected pages' SVGs. `convert_docx_to_layout` and
  `convert_docx_to_svg` each lay it out on their own, as they always have.
- **Element paths.** Every element the SVG draws from the document carries
  `data-docx-path`, and its layout object carries the same string as `path`: a `Line`'s
  is its paragraph's (`w:body/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1]`), a `Span`'s its run's, a
  `Float`'s its `wp:anchor`'s, a `Stop`'s the obstacle's. `resolve_path(root, path)` finds
  the element a path names in a part's root element, parsed by the standard library or
  by lxml; `path_part(package, path, line.part)` says which part (the main part, a
  header's or footer's -- `Line.part` -- or the notes part for `w:footnote[@w:id=N]/...`),
  and `locate(docx, path, part=None)` does both, returning `(part, element)`. Steps
  count same-named siblings by local name, as XPath does, except rows and cells, which are
  counted through row- and cell-level content controls; the rules are in
  `docx2svg.paths`, which the parser counts with too.

  ```python
  for line in result.layout.pages[0].lines:
      part, paragraph = locate("report.docx", line.path, line.part)
  ```
- **`Line.column`** is the 0-based text column a line is in, in a section of several
  columns (a footnote's line the column it stands under), and `None` in a section of one
  column and for a header's, footer's or text box's line.
- **`Page.info`** is a `docx2svg.PageInfo` (public): the number Word prints on the page
  (`info.number`, before formatting, restarts and blank pages counted -- `Page.number` is
  the page's 0-based place), the 0-based index of its section in `Document.sections`,
  whether that section starts on it, the header and footer kind it shows (`first`,
  `even`, `default`) and whether it is a blank page Word inserts. With
  `docx2svg.fields.field_text` it gives what a `PAGE` or `PAGEREF` to the page shows:

  ```python
  from docx2svg.fields import field_text

  page = layout.pages[4]
  fmt = parsed.sections[page.info.section].page_number_format
  field_text(" PAGEREF _Toc1 \\h ", "PAGEREF", page=page.info.number, pages=None,
             section_pages=None, section_format=fmt)          # "v" in a lowerRoman section
  ```
- **List numbers without a document**: `docx2svg.linebreak.ListCounters().item(numbering,
  numId, ilvl)` counts the next item of an instance at a level as Word counts it and
  returns its label and number (`ListItem("2.1.", 1)`), from the numbering part alone
  (`docx2svg.parse.styles.parse_numbering`). One counter per story, asked in document
  order (ROADMAP.md, "Numbering and lists — measured").
- **A note separator** is drawn as a group of its own, `data-docx-path="w:footnotes/separator"`
  (`continuationSeparator`; `w:endnotes/...`) holding one line group with no text, its
  line a `<rect data-docx-kind="separator">` whose path is that group's run
  (`.../w:r[1]`). Where the document names no separator of its own, Word's is drawn, so
  the group has no `data-docx-id` and its path resolves to no element; where the
  settings name the notes part's own, the paths are that note's
  (`w:footnote[@w:id=-1]/w:p[1]`). It is not a paragraph of the text.

