#!/usr/bin/env python3
"""Score SmartArt against Word: ``make_smartart_probe.py``'s diagrams as Word's PDF draws
them and as docx2svg draws the drawing cached for them (``docx2svg.diagram``).

Word's side is read as ``read_chart_probe.py`` reads a chart's -- every span of text that
is not the probe's own (Georgia), every filled shape and every stroked segment -- and so is
the model's, off the diagram's SVG fragment; :func:`read_chart_probe.compare` pairs them.

The documents carry the drawing Word caches for their data (``make_smartart_probe.CACHES``).
``--caches`` has Word save each document again (``tools/oracle.py``'s ``resave``) and
reads the cache it wrote into that table, kept with the recording; every run checks that
the cache the documents carry is still the one Word writes.

``--record`` writes the caches, Word's side and the scores to
``tests/fixtures/smartart-observations.json``; ``tests/test_smartart.py`` holds the model to
them offline.

Usage::

    python tools/read_smartart_probe.py [-v] [--caches] [--record]
"""

from __future__ import annotations

import argparse
import io
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_chart_probe as chart  # noqa: E402
import read_render  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "smartart-observations.json"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DSP = "{http://schemas.microsoft.com/office/drawing/2008/diagram}"


def documents() -> list[tuple[str, bytes]]:
    import make_smartart_probe

    return [(f"smartart-{s}", make_smartart_probe.build(s)) for s in make_smartart_probe.SETTINGS]


def _colour(element) -> str | None:
    """A fill's colour as the probe's table writes it: ``accent1``, ``accent1@60000``."""
    if element is None or element.find(f"{A}noFill") is not None:
        return None
    scheme = element.find(f"{A}solidFill/{A}schemeClr")
    if scheme is None:
        return "?"
    tint = scheme.find(f"{A}tint")
    return scheme.get("val") + (f"@{tint.get('val')}" if tint is not None else "")


def _fill(sp_pr) -> str | None:
    for node in sp_pr:
        tag = node.tag.replace(A, "")
        if tag == "noFill":
            return None
        if tag == "solidFill":
            holder = ElementTree.Element("x")
            holder.append(node)
            return _colour(holder)
    return None


def cache_of(xml: bytes) -> list:
    """A cached drawing's shapes in :data:`make_smartart_probe.CACHES`'s terms."""
    root = ElementTree.fromstring(xml)
    out = []
    for sp in root.iter(f"{DSP}sp"):
        sp_pr = sp.find(f"{DSP}spPr")
        xfrm = sp_pr.find(f"{A}xfrm")
        off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
        line = sp_pr.find(f"{A}ln")
        line_colour = _colour(line) if line is not None else None
        body = sp.find(f"{DSP}txBody")
        paragraphs, insets, anchor, box = [], [0, 0, 0, 0], "ctr", [0, 0, 0, 0]
        if body is not None:
            body_pr = body.find(f"{A}bodyPr")
            insets = [int(body_pr.get(k, 0)) for k in ("lIns", "tIns", "rIns", "bIns")]
            anchor = body_pr.get("anchor", "t")
            for p in body.findall(f"{A}p"):
                p_pr = p.find(f"{A}pPr")
                runs = p.findall(f"{A}r")
                if not runs:
                    continue
                bullet = p_pr.find(f"{A}buChar")
                spacing = p_pr.find(f"{A}lnSpc/{A}spcPct")
                after = p_pr.find(f"{A}spcAft/{A}spcPct")
                paragraphs.append(["".join(r.find(f"{A}t").text or "" for r in runs),
                                   int(runs[0].find(f"{A}rPr").get("sz")), p_pr.get("algn", "l"),
                                   bullet.get("char") if bullet is not None else None, int(p_pr.get("marL", 0)),
                                   int(p_pr.get("indent", 0)),
                                   int(spacing.get("val")) if spacing is not None else 100000,
                                   int(after.get("val")) if after is not None else 0])
            tx = sp.find(f"{DSP}txXfrm")
            box = [int(tx.find(f"{A}off").get(k)) for k in ("x", "y")] + [int(tx.find(f"{A}ext").get(k))
                                                                            for k in ("cx", "cy")]
        out.append([sp_pr.find(f"{A}prstGeom").get("prst"),
                    [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))],
                    int(xfrm.get("rot", 0)), _fill(sp_pr),
                    [line_colour, int(line.get("w", 0))] if line_colour else None,
                    box, insets, anchor, paragraphs])
    return out


def word_caches(data: bytes, name: str) -> tuple[dict[int, list], dict[int, list]]:
    """The caches Word writes for ``data``'s diagrams, saving it again, and the effect
    extents it gives their frames."""
    import make_smartart_probe
    import oracle

    saved = zipfile.ZipFile(io.BytesIO(oracle.resave(data, name=name)))
    caches = {k: cache_of(saved.read(f"word/diagrams/drawing{k + 1}.xml"))
              for k in range(len(make_smartart_probe.CASES))}
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    root = ElementTree.fromstring(saved.read("word/document.xml"))
    effects = {}
    for frame in list(root.iter(f"{wp}inline")) + list(root.iter(f"{wp}anchor")):
        number = int(frame.find(f"{wp}docPr").get("id")) - 1
        extent = frame.find(f"{wp}effectExtent")
        effects[number] = [int(extent.get(side, 0)) for side in ("l", "t", "r", "b")]
    return caches, effects


def word_side(pdf: Path) -> list[dict]:
    import make_smartart_probe
    import pymupdf

    with pymupdf.open(str(pdf)) as document:
        return [chart.word_page(document[index]) for index in range(len(make_smartart_probe.CASES))]


def score(data: bytes, fonts, word: list[dict]) -> dict:
    import render_record
    from docx2svg import _render

    _documents, layout, _data, _fonts = _render(data, render_record.options(fonts))
    model = chart.model_side(layout, "diagram")
    totals: dict = {}
    cases, problems = [], {}
    for index, (word_row, model_row) in enumerate(zip(word, model)):
        found, wrong = chart.compare(word_row, model_row)
        cases.append(found)
        if wrong:
            problems[index] = wrong
        for key, (agree, compared) in found.items():
            totals.setdefault(key, [0, 0])
            totals[key][0] += agree
            totals[key][1] += compared
    return {"scores": totals, "cases": cases, "problems": problems}


def main(argv: list[str] | None = None) -> int:
    import importlib

    import make_smartart_probe
    import oracle
    import render_record

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--caches", action="store_true")
    parser.add_argument("--record", action="store_true")
    options = parser.parse_args(argv)
    recorded = json.loads(OBSERVATIONS.read_text(encoding="utf-8")) if OBSERVATIONS.exists() else {}
    if options.caches:
        caches, effects = word_caches(make_smartart_probe.build("15"), "smartart-15")
        recorded["caches"] = {str(k): v for k, v in caches.items()}
        recorded["effects"] = {str(k): v for k, v in effects.items()}
        render_record.dump(OBSERVATIONS, recorded)
        importlib.reload(make_smartart_probe)
        print("caches read from Word's copy into", OBSERVATIONS)
    recorded["faces"], recorded["documents"] = {}, {}
    for name, data in documents():
        written, effects = word_caches(data, name)
        if written != make_smartart_probe.CACHES:
            print(name, "Word writes a different cache from the one the document carries; run --caches")
        if effects != make_smartart_probe.EFFECTS:
            # Word sets each frame's effect extent from its layout, and not alike in every
            # setting (no settings part: 2 pt on the right, mode 14 and 15: 3 pt); only
            # what follows a diagram on its line moves with it, and nothing does here.
            print(name, "Word gives the frames other effect extents:", effects)
        pdf = oracle.export(data, name=name)
        word = word_side(pdf)
        fonts = render_record.RecordingFonts(data, recorded["faces"])
        result = score(data, fonts, word)
        print(name, " ".join(f"{k} {a}/{c}" for k, (a, c) in result["scores"].items()))
        for index, found in enumerate(result["cases"]):
            case = make_smartart_probe.CASES[index]
            print(f"  {index:2d} {case.note:40s} " + " ".join(f"{k} {a}/{c}" for k, (a, c) in found.items()))
            if options.verbose:
                for problem in result["problems"].get(index, []):
                    print("       ", *problem)
        recorded["documents"][name] = {"word": word, "scores": result["scores"]}
    if options.record:
        render_record.dump(OBSERVATIONS, recorded)
        print("recorded", OBSERVATIONS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
