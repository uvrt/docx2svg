# Four third-party documents that are **not** in this repository

This directory holds only this note. The documents it describes may not be
redistributed, so they are not committed, and neither are Word's lines of them (which
are their text). They live in the gitignored `scratch/filesamples/`, and the tests that
use them (`tests/test_filesamples.py`) skip wherever they are absent -- the same rule as
`pptx2svg`'s `real-college-template`.

- **Source:** <https://filesamples.com/formats/docx>
- **Downloaded:** 2026-09-24
- **Licence: none.** The page (checked on 2026-09-24) states no licence, terms of use
  or permission for the files; its only statement of rights is the footer
  *"© 2026 FileSamples"*. Without a grant they are used here only as local measurement
  inputs, never copied into git or into any published artifact.

| File | Size | SHA-256 | Written by | `compatibilityMode` |
| --- | --- | --- | --- | --- |
| `sample1.docx` | 1,311,881 | `269329fc7ae54b3f289b3ac52efde387edc2e566ef9a48d637e841022c7e0eab` | Microsoft Office Word 12 (2007) | none (empty `w:compat`: Word 2007 mode) |
| `sample2.docx` | 120,515 | `40258ea32f3175c3a91cab65dd6a7eddeb94180edd186ad882ca51b8b0dfa7b4` | Microsoft Macintosh Word 15 | 14 |
| `sample3.docx` | 34,375 | `58211ac149c4b05110cb95959f0f5f529333af5aea060d5a5baf096f531848b8` | Microsoft Office Word 14 | 14 |
| `sample4.docx` | 14,169,117 | `7c120af503a7b2c72c756b11dd42dc6d666e4756b94b14f0b7b65c181123434e` | Microsoft Office Word 15 | 15 |

**`sample1.docx` embeds fonts**: Ubuntu (four styles), Ubuntu Mono and Tahoma, as
obfuscated `word/fonts/*.odttf` parts. That is one more reason the file stays out of
git. `tools/face_metrics.embedded` reads the four vertical-metric integers from them in
memory; nothing else leaves the file.

To reproduce: download the four files into `scratch/filesamples/`, check the hashes,
and run `python tools/read_baselines.py --record-scratch` on a machine with Word.
