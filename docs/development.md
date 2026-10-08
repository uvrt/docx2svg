# Development

## Running the tests

`python -m pytest` runs everything that needs nothing but the repository (about five
minutes; tests that need Word, its faces or `scratch/` skip where those are absent).
`--run-slow` adds the raster fidelity scores (`tools/fidelity.py`), which need Word's
exports and faces and take several minutes more.  The fidelity instrument rasterises Word's
page with the same engine as ours (resvg) after converting it to SVG with PyMuPDF
(`tools/pdf_svg.py`), a development-only tool behind the `fidelity` extra; install it into
a project virtual environment and run the suite from there:

```
python3 -m venv --system-site-packages .venv
.venv/bin/pip install pymupdf           # or: .venv/bin/pip install -e .[dev,fidelity]
.venv/bin/python -m pytest --run-slow
```

PyMuPDF is AGPL-3.0, acceptable for a local tool that is never distributed with the
library; nothing under `src/` imports it, and without it the tests that need it skip.
Converted pages carry the glyph outlines Word embedded (Microsoft's fonts), so they are
cached only beside Word's exports, outside the repository, and never committed.

`tools/fidelity.py` and `tools/pdf_svg.py --validate` use every logical core (fewer if
memory is short): the harness scores a page per process and the validation takes a page
per process, and both reassemble the results in order, so every score, baseline and
printed line is the serial run's. `--jobs 1` is the serial path. Neither launches Word:
exports are looked up in the calling process before any worker starts, and a worker
refuses to export.

The harness scores the default instrument only -- Word's page through the converter and
resvg (`--truth svg`), our glyphs at the device size (`--glyph-size device`) -- unless
asked for more: `--truth pdfium` / `both` and `--glyph-size exact` / `both`. The baselines
keep the recorded pdfium and exact scores beside the default ones; a default `--record`
carries them over untouched, and `--record --truth both --glyph-size both` re-records
them. In the suite, `pytest --pdfium` and `pytest --exact` do the same.

It caches both sides' rasters beside Word's exports, in the oracle directory's
`svg/rasters/` (never in the repository), keyed by every input that moves their pixels --
the SVG and PDF bytes, the *contents* of every font file resvg is handed, the converter,
resvg and Pillow versions and the options -- and holds itself to them: each run re-draws a
few pages, rotating through the corpus by date, and compares them with the cache byte for
byte. Any difference discards the whole cache and stops the run. `--verify-cache` re-draws
every page and compares; `--no-cache` bypasses the cache. SSIM is computed over the
content's bounding box only, which gives the full page's score bit for bit; the same
sampled pages are scored both ways each run and held equal.

A bare `pytest` runs serially, as CI does. To use every core, pass `-n` (pytest-xdist, in
the `dev` extra):

```
.venv/bin/python -m pytest -n auto
```

It collects, passes and skips exactly the tests the serial run does. Here `-n` means
`--dist loadgroup` (`tests/conftest.py`): `load` for every test but those marked
`@pytest.mark.word` -- the tests of Word's export of the layout sweep, and any test that
requests the `oracle_pdf` fixture, which exports through Word when the export is not
cached. Those all run on one worker, one at a time, since Word is one instance per
machine; under any other `--dist` they fail rather than race.

Across processes -- and with docx-agent, which drives the same Word -- every launch,
export, re-save and recovery holds an advisory `flock` on `word-oracle.lock` in the Office
group container (`tools/oracle.py`, `oracle.word()`). The lock is waited for, up to ten
minutes (`DOCX2SVG_WORD_LOCK_TIMEOUT`, seconds), and a Word found running once it is held
is someone else's: it is never quit, and the export refuses (`WordBusy`; a test skips).
Only the Word this process started is quit, and only the `~$` files in its own oracle
directory are deleted.

## Constraints

**Standard-library only at runtime** (and `ooxml-common`, which is too). `pyproject.toml` is the
statement of it. Reading the oracle's PDF needs `pypdfium2` and lives behind the `oracle`
extra; measuring real font files needs `fonttools` and lives behind `measure`. Neither is
imported from `src/`. PNG output needs a rasteriser, the `png` extra (`resvg-py`), imported
by `docx2svg.png` only when a PNG is asked for.

**No Microsoft font file enters this repository**, in any form — not a file, not a subset,
not an embedded blob inside a test artefact, not a Word-exported PDF (which carries
subsets, and is why the ground-truth exports live in the Office group container
(`~/Library/Group Containers/UBF8T346G9.Office/docx2svg-oracle/`) rather than in the tree). Measurements taken from those faces are facts and are recorded here as numbers.
The files stay on the machine that measured them.
`src/docx2svg/recorded.py` is such a record: Symbol's and Wingdings' metrics and
advances, written by `tools/record_symbol_faces.py` from Word's copies. No open font file
enters the repository either. CI installs Carlito, Caladea and Liberation from the
system's package manager. To lay out locally as a machine without Office does, point
`DOCX2SVG_SUBSTITUTE_FONTS` at a folder holding those faces: `tests/test_substitutes.py`
then lays out with nothing else visible.

