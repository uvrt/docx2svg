#!/usr/bin/env python3
"""Which of ``ooxml-common``'s tables are the faces Word lays out with?

``ooxml-common``'s advance and kern tables were generated for PowerPoint, some of them
from metric clones (Carlito for Calibri, Arimo for Arial, Tinos for Times New Roman,
Cousine for Courier New).  Before ``docx2svg.measure.TableAdvances`` trusts one, every
advance in it -- upright and bold -- is compared here with the installed face Word lays
out with (``face_advances``: the macOS copy first, as Word does for Times New Roman), and
every kern pair among the table's characters with that face's legacy ``kern`` table,
which is what Word applies (ROADMAP.md, Phase 2 and "Phase 3 -- measured").

Usage::

    python tools/check_advance_tables.py
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import face_advances  # noqa: E402

#: Word's face -> the table ooxml-common's fontmap gives it (``metrics_for``).
CANDIDATES = {
    "Calibri": "Carlito", "Calibri Light": "Carlito", "Arial": "Arimo", "Times New Roman": "Tinos",
    "Courier New": "Cousine", "Cambria": "Cambria", "Aptos": "Aptos",
}


def main() -> int:
    from ooxml_common.text.metrics import METRICS

    for office, name in CANDIDATES.items():
        table = METRICS[name]
        for bold in (False, True):
            widths = table.bold_widths if bold else table.widths
            found = face_advances.advances(office, bold, False)
            label = f"{office:16} {'bold' if bold else 'upright':8} ({name})"
            if found is None or not widths:
                print(f"{label}: not installed or no table")
                continue
            differ = [(c, w, found.advance(c)) for c, w in widths.items()
                      if found.advance(c) is not None and found.advance(c) != w]
            kern_differ = 0
            if table.kerning is not None:
                chars = list(widths)
                kern_differ = sum(1 for a in chars for b in chars
                                  if table.kerning.adjustment(a, b, bold) != found.kern(a, b))
            print(f"{label}: {len(widths) - len(differ)} / {len(widths)} advances agree"
                  + (f", differ: {differ[:4]}" if differ else "")
                  + f"; kern pairs differing from the legacy table: {kern_differ}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
