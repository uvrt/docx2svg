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
- **Laying out once, writing pages later**: `lay_out_docx(source, options)` returns
  `(layout, data, fonts)` -- every page laid out, the document's bytes and the faces it was
  measured with, all three picklable, so the layout can be kept or handed to another
  process -- and `page_svgs(layout, data, options, fonts=fonts)` writes the pages
  `options.pages` selects, byte for byte `convert_docx_to_svg`'s (`device_size=True`: at
  300 dpi when no size is asked, as `convert_docx_to_png` rasterises). `user_font_dirs(options)`
  is the application's font folders the rasteriser is handed after the faces the SVG
  names. A render tool drawing a few pages per call lays a version out once this way.
- <a id="relayout"></a>**Laying out again from a previous layout** (round 19, R1). After a
  small edit, `lay_out_docx(new_bytes, options, previous=layout)` lays out only what the edit
  reaches -- the page holding the first changed block (or the one before, where its end may
  depend on it) until the pages line up with `layout`'s again -- and keeps the rest of
  `layout`'s pages, renumbered, moved on and with their headers and footers drawn again where
  that is all that changed for them. The result is what a layout of `new_bytes` alone is,
  every page's SVG byte for byte, and says what to draw again:

  ```python
  layout0, data0, fonts0 = lay_out_docx(data0)
  layout1, data1, fonts1 = lay_out_docx(data1, previous=layout0)
  layout1.changed_pages        # [17]: only the page the edit is on; [17, 18, 19] if it pushes text on
  layout1.relaid               # None; or why the whole document was laid out again ("styles", ...)
  layout1.relaid_cause         # None; or {"reason": "fields-open", "at": "w:body/w:p[4]",
                               #   "para_id": "1A2B3C4D", "field": "TOC \\o \"1-3\" \\h"}
  svgs = page_svgs(layout1, data1, ConvertOptions(pages=layout1.changed_pages), fonts=fonts1)
  ```

  `changed_pages` is the pages whose drawing differs from `previous`'s page of the same
  number (a page whose only change is its elements' `data-docx-path` is not one) and the
  pages `previous` did not have. The changed blocks are found by comparing the bytes;
  `changed_blocks=[...]` (paths or paraIds) is accepted as a hint only. The whole document is
  laid out again -- the unchanged blocks' measurements still reused where the styles and
  settings are the same -- where an edit cannot be scoped: `docx2svg.relayout.FALLBACKS`
  lists each reason `relaid` gives, and `relaid_cause` says what caused it: the body element
  (`at`, its path, and `para_id`, its `w14:paraId`), the field an edited element opens and
  another closes (`field`), or the part that changed (`part`: `word/styles.xml`), to explain
  a slow refresh. A body with page-number fields -- a table of contents (its lines'
  `PAGEREF`s), `PAGE`, `NUMPAGES` -- or `SEQ` and `REF` is laid out from the edit too
  (round 20): a full layout lays such a body out in passes until the fields' values settle,
  and each pass is laid out from the same pass of `previous`, only where a block or a field's
  value differs (the edit; a contents line whose number changed; the captions after an
  inserted one), each stretch until the pages line up again -- a page break after it
  typically, the pages after the break then `previous`'s, moved on where pages were added or
  removed before it. `changed_pages` then holds the contents page where a number on it
  changed. `previous` must be a layout this function made, with
  the same options; one of another docx2svg, of other options, or made with `keep=False`
  (which keeps nothing for later: `layout.state` is `None`) is laid out whole. The state a
  layout keeps (`layout.state`) is picklable with it, so a worker process can hand it back.
- <a id="warm"></a>**A worker's one-time work, done at its start**: a process's first
  layout pays for what later ones reuse -- the modules, the font folders' index (which
  faces each installed file holds), and each face it uses read from its file.
  `warm(source=None, options=None, *, scan=None)` does that work in advance, for the font
  folders `options` names: call it when a render worker starts. With `source` (a
  document, typically the template a worker's documents share) it also lays that document
  out and draws its first page, which loads the faces and hyphenation patterns it uses;
  its output is discarded and `options`' warnings and coverage are left untouched. It
  returns a `docx2svg.fonts.FontScan`, what was read of the folders' files, picklable:
  save it and hand it to the next worker's `warm(scan=...)`, which then reads only the
  files whose modification time or size changed (and ignores a scan of another docx2svg
  version). Output is the same bytes with and without it (`tests/test_warm.py`).

  ```python
  import pickle
  scan = docx2svg.warm(source="template.docx", options=ConvertOptions(font_dirs=["/srv/fonts"]))
  pickle.dump(scan, open("font-scan.pickle", "wb"))       # next worker: warm(scan=pickle.load(...))
  ```
- <a id="locale"></a>**Locale.** Word takes a few things from the machine it runs on
  rather than from the document: a decimal tab stop aligns on the *system's* decimal
  separator, not the document language's (measured: on a machine with a decimal comma,
  `1234.5` after a decimal stop in an `en-GB` document is drawn as after a right stop), and
  a chart's numbers are printed in its decimal and thousands separators. A renderer that
  read them from *its* machine would draw the same document differently on every server,
  so docx2svg never does: the caller says which locale a conversion is in
  (`ConvertOptions.locale`, an `ooxml_common.locales.Locale`), and without one it is
  `EN_US` -- a decimal stop aligns on `.`, a chart prints `1,240.0`.

  ```python
  from ooxml_common.locales import Locale

  nl = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
              language="nl-NL", list_separator=";")
  svgs = convert_docx_to_svg("report.docx", ConvertOptions(locale=nl))
  ```

  `Locale` is frozen and picklable, so the options travel into worker processes
  unchanged. Only the separators are read so far: `DATE` and `TIME` fields and ordinal or
  cardinal number text are drawn as Word cached them, and a chart's own "Chart Title" is
  not drawn, each with a warning, whatever the locale.
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
- <a id="page-blocks"></a>**Pages and blocks, both ways** (round 19, R3). `Page.blocks` is
  the keys of the body's blocks a page draws something of -- every paragraph, table, row
  and cell, as the path of its element in the main part (`w:body/w:tbl[2]`,
  `w:body/w:tbl[2]/w:tr[3]`, `w:body/w:tbl[2]/w:tr[3]/w:tc[1]`, its paragraph's
  `.../w:p[1]`), once each, in the order the page draws them. `layout.blocks_on(page)`
  is the top-level ones (paragraphs and tables) on a 1-based page; `layout.pages_of(key)`
  the 1-based pages a block is drawn on -- every one, for a paragraph, table or row split
  across pages, or a header row drawn again -- taking a block's path, a path inside a
  paragraph (a run's: its paragraph's), or a paragraph's `w14:paraId`. A key is derived
  from the XML alone (`docx2svg.paths.block_keys(path)` gives a drawn object's), so a
  caller that edits the document maps its own ids onto it; headers, footers and notes are
  not the body's and are not listed.

  ```python
  layout.blocks_on(3)                  # ['w:body/w:p[26]', ..., 'w:body/w:tbl[3]', ...]
  layout.pages_of("w:body/w:tbl[3]")   # [3, 4]: the table goes on to page 4
  layout.pages_of("1A2B3C4D")          # by w14:paraId
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

## Used by docx-agent

docx-agent imports these names from docx2svg, in its `src`, `tests` and `tools`, beyond the
functions above. They are docx2svg's surface across repositories: docx-agent's CI installs
docx2svg from `main`, so renaming or removing one breaks it at its next run.
`tests/test_public_api.py` lists the same names and fails when one is gone, or when this
list and that one disagree. To rename one, keep the old name as an alias for one release,
with a comment saying who it is kept for, and move docx-agent to the new name.

- `docx2svg` -- docx-agent: `ConvertOptions`, `available_backends`, `convert_docx`, `convert_docx_to_layout`, `convert_docx_to_png`, `convert_docx_to_svg`, `drawn_faces`, `lay_out_docx`, `page_svgs`, `user_font_dirs`, `warm`
- `docx2svg.coverage` -- docx-agent: `Coverage`
- `docx2svg.fields` -- docx-agent: `Uncomputable`, `field_text`, `sequence_step`, `sequence_switches`
- `docx2svg.fonts` -- docx-agent: `InstalledFonts`, `cloud_font_dirs`, `default_font_dirs`, `installed_index`, `rasteriser_files`
- `docx2svg.linebreak` -- docx-agent: `ListCounters`
- `docx2svg.model` -- docx-agent: `Paragraph`, `ParagraphProperties`, `Run`, `RunProperties`, `Table`
- `docx2svg.paginate` -- docx-agent: `CARRY_ON`
- `docx2svg.parse.document` -- docx-agent: `cell_conditions`, `parse_package`
- `docx2svg.parse.properties` -- docx-agent: `read_paragraph_properties`, `read_run_properties`
- `docx2svg.parse.styles` -- docx-agent: `Numbering`, `parse_numbering`, `parse_styles`
- `docx2svg.paths` -- docx-agent: `block_keys`, `path_part`, `resolve_path`
- `docx2svg.png` -- docx-agent: `svg_to_png`
- `docx2svg.resolve` -- docx-agent: `character_format`, `resolve_paragraph`, `resolve_run`
- `docx2svg.table` -- docx-agent: `edge_winner`, `outer_part`, `resolve_table`, `table_style_id`
