#!/usr/bin/env python3
"""Measure ``make_script_probe.py`` with the model: baselines exact, by kind and rule;
and the size and offset Word gives every ``w:vertAlign`` run.

Every document goes through ``baselines.predict`` (``probe_documents``), so the line
height scored is the model's own.  Quartz draws a raised or lowered run on a baseline of
its own; here every line whose text is only the probe's extra run (``Hxh``) is folded
back into the nearest line that holds a paragraph's text, and its distance from that
baseline is kept as the run's **offset**.  The run's **size** is its advance: the pen x
of the text object after it, less its own, over the advance of ``Hxh`` in font units
(which the reader takes from the installed face; only the size is recorded).

``--record`` writes ``tests/fixtures/script-observations.json`` for
``tests/test_script.py``: per document the baselines of each page (the text is the
generator's), the offsets and sizes, and the four metric integers of every face.

Usage::

    python tools/read_script_probe.py [--record]
"""

from __future__ import annotations

import functools
import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import baselines  # noqa: E402
import make_script_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "script-observations.json"


# -- what Word drew --------------------------------------------------------------------


def paragraph_texts(p: probe.Probe) -> list[str]:
    """The text of every paragraph of ``p``, in order, as the drawn lines carry it."""
    out = []
    for group in p.groups:
        extra = {"plain": "", "supsub": probe.SCRIPT * 2}.get(group.kind, probe.SCRIPT)
        tail = "" if group.kind == "plain" else f" {probe.SCRIPT}"
        out += [f"Hxample {index} {extra}{tail}" for index in range(probe.LINES)]
    out += [f"Hx {probe.SCRIPT} Hx {probe.SCRIPT} Hx" for _ in p.sizes]
    return out


def fold(lines):
    """``quartz_pdf.lines`` with every line of the extra run alone folded into its host.

    On a page every paragraph holds the same number of extra runs, each at the same
    offset from its own baseline, so the ``k``-th of them from the top belongs to the
    ``k // n``-th paragraph from the top.  Not the nearest: a run twice the text's size,
    or raised by 24 pt, under ``exact`` is drawn nearer the paragraph above.  Returns
    ``(lines, offsets)``: ``offsets[i]`` lists ``(drawn y - host y, the run)`` for every
    run folded into the ``i``-th line.
    """
    import quartz_pdf

    hosts = [line for line in lines if line.text.strip() != probe.SCRIPT]
    runs: dict[int, list] = defaultdict(list)
    for page in sorted({line.page for line in lines}):
        mine = sorted((i for i, host in enumerate(hosts) if host.page == page), key=lambda i: hosts[i].y)
        extra = sorted((line for line in lines if line.page == page and line.text.strip() == probe.SCRIPT),
                       key=lambda line: line.y)
        if not extra:
            continue
        per, rest = divmod(len(extra), len(mine))
        for k, line in enumerate(extra):
            # The size sweep's smallest scripts are drawn on their host's baseline, so
            # its pages hold uneven numbers; its offsets are small, and it takes the
            # nearest paragraph.
            i = mine[k // per] if per and not rest else min(mine, key=lambda i: abs(hosts[i].y - line.y))
            runs[i] += [(line.y - hosts[i].y, run) for run in line.runs]
    folded = [quartz_pdf.Line(host.page, host.y, sorted(host.runs + [r for _, r in runs[i]], key=lambda r: r.x))
              for i, host in enumerate(hosts)]
    return folded, [runs[i] for i in range(len(hosts))]


@functools.lru_cache(maxsize=None)
def advance_em(face: str) -> float:
    """The advance of :data:`make_script_probe.SCRIPT` in ems, from the installed face."""
    import face_metrics
    from fontTools.ttLib import TTFont

    path, number = face_metrics._index()[(face.lower(), False, False)]
    if face == "Aptos":
        path = path.replace("Aptos-Light", "Aptos")  # the index finds the Light first
    font = TTFont(path, fontNumber=number, lazy=True)
    cmap, hmtx = font.getBestCmap(), font["hmtx"]
    return sum(hmtx[cmap[ord(c)]][0] for c in probe.SCRIPT) / font["head"].unitsPerEm


def measure(p: probe.Probe):
    """Word's lines of ``p``: ``(drawn lines, {line index: [[dy, size in half points]]})``.

    The size is the advance of the run over ``SCRIPT``'s advance at 1 em, in half points
    to 1e-3 (the pen x of the text object after the run is exact to 1e-4 px).
    """
    import oracle
    import quartz_pdf

    pages = quartz_pdf.read(oracle.export(probe.build(p), name=p.name))
    lines, extras = fold(quartz_pdf.lines(pages))
    faces = [g.face for g in p.groups for _ in range(probe.LINES)] + [face for face, _ in p.sizes]
    scripts = {}
    for index, (line, folded) in enumerate(zip(lines, extras)):
        runs = line.runs
        found = []
        for position, run in enumerate(runs):
            if run.text != probe.SCRIPT or position + 1 == len(runs):
                continue
            dy = next((d for d, r in folded if r is run), 0.0)
            em = advance_em(faces[index])
            half_points = (runs[position + 1].x - run.x) / em / (300 / 72) * 2
            found.append([round(dy), round(half_points, 3)])
        if found:
            scripts[index] = found
    drawn = [baselines.DrawnLine(line.page, line.y, tuple((r.font, r.size_px, r.text) for r in line.runs))
             for line in lines]
    return drawn, scripts


def compact(p: probe.Probe, drawn) -> list:
    """The baselines of each page, ``[[baseline, ...], ...]`` (every page holds a line);
    the text is the generator's."""
    texts = paragraph_texts(p)
    assert len(drawn) == len(texts), (p.name, len(drawn), len(texts))
    pages: list[list] = []
    for line, text in zip(drawn, texts):
        assert baselines._norm(line.text) == baselines._norm(text), (p.name, line.text, text)
        assert line.y == int(line.y) and line.page in (len(pages) - 1, len(pages)), (p.name, line)
        if line.page == len(pages):
            pages.append([])
        pages[-1].append(int(line.y))
    return pages


def expand(p: probe.Probe, pages) -> list[baselines.DrawnLine]:
    lines = [(page, y) for page, ys in enumerate(pages) for y in ys]
    return [baselines.DrawnLine(page, y, (("", 0, text),))
            for (page, y), text in zip(lines, paragraph_texts(p))]


def by_group(p: probe.Probe, found: dict) -> dict:
    """``{group (or size-sweep line): sorted distinct [offset px, size half points]}``:
    every line of a group draws its extra runs alike."""
    names = kinds(p)
    out: dict = defaultdict(set)
    for index, runs in found.items():
        if names[index].startswith("bdr"):
            continue  # a bordered run is not moved, and its advance holds the border's space
        out[names[index]].update(tuple(run) for run in runs)
    return {name: sorted(map(list, runs)) for name, runs in out.items()}


# -- scoring -------------------------------------------------------------------------------


def kinds(p: probe.Probe) -> list[str]:
    return probe.kinds(p)


def tabulate(scored, key=lambda kind: kind.split("/")[0]) -> dict:
    """``{key(group name): [exact, scored]}`` over ``[(probe, results)]``, every line
    counted by the paragraph it belongs to."""
    table: dict = defaultdict(lambda: [0, 0])
    for p, results in scored:
        names = kinds(p)
        for r in results:
            if r.status not in ("exact", "miss") or r.block < 0:
                continue
            cell = table[key(names[r.block])]
            cell[0] += r.status == "exact"
            cell[1] += 1
    return dict(table)


def family(name: str) -> str:
    """``sup@44`` -> ``sup, 2x``; ``pos+6`` -> ``pos``; ``bdr4x0@11`` -> ``bdr, 1/2x``."""
    kind, _, face_size = name.split("/")[:3] if name.count("/") >= 2 else (name, "", "")
    base = kind.split("@")[0].rstrip("+-0123456789x")
    if "@" not in kind:
        return base
    extra = int(kind.split("@")[1])
    text = int("".join(c for c in name.split("/")[1] if c.isdigit()))
    return f"{base}, {'2x' if extra > text else '1/2x'}"


def recorded(data: dict, section: str, name: str):
    """A document's recording; one Word drew exactly as another is recorded as that
    other's name (the three compatibility settings draw every line alike)."""
    value = data[section][name]
    return data[section][value] if isinstance(value, str) else value


def offline(data: dict | None = None, which=lambda p: True):
    """``[(probe, results)]`` from the recording, for the documents ``which`` selects."""
    data = data or json.loads(OBSERVATIONS.read_text())
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(p, probe_documents.score(probe.build(p), expand(p, recorded(data, "documents", p.name)), metrics))
            for p in probe.PROBES if p.name in data["documents"] and which(p)]


def report(scored) -> None:
    by_family = tabulate(scored, lambda k: k.split("/")[0] if k.startswith("size") else family(k))
    for name, (exact, count) in sorted(by_family.items()):
        print(f"{name:14} {exact:6} / {count}")
    by_probe = {p.name: tabulate([(p, r)], lambda k: "all")["all"] for p, r in scored}
    for name, (exact, count) in by_probe.items():
        print(f"  {name:16} {exact:6} / {count}")
    for p, results in scored:
        names = kinds(p)
        pages: dict = defaultdict(list)
        for r in results:
            if r.block >= 0 and r.predicted is not None:
                pages[names[r.block]].append((r.observed, r.predicted))
        shown = 0
        for group, ys in pages.items():
            if all(o == q for o, q in ys) or group.startswith("size") or shown > 40:
                continue
            shown += 1
            n = len(ys) - 1
            observed = (ys[-1][0] - ys[0][0]) / n
            predicted = (ys[-1][1] - ys[0][1]) / n
            misses = sum(o != q for o, q in ys)
            print(f"  {p.name:14} {group:44} pitch observed {observed:8.3f} model {predicted:8.3f}"
                  f"  first {ys[0][0] - ys[0][1]:+d}  misses {misses}")


def main(argv: list[str]) -> int:
    recording = probe_documents.RecordingMetrics()
    documents, scripts, scored = {}, {}, []
    for p in probe.PROBES:
        drawn, found = measure(p)
        documents[p.name] = compact(p, drawn)
        scripts[p.name] = by_group(p, found)
        scored.append((p, probe_documents.score(probe.build(p), drawn, recording)))
    report(scored)
    if "--record" in argv[1:]:
        for section in (documents, scripts):
            for name in list(section):
                same = next((other for other in section if other != name and not isinstance(section[other], str)
                             and section[other] == section[name]), None)
                if same is not None and list(section).index(same) < list(section).index(name):
                    section[name] = same
        payload = {
            "_about": ("Every line Word 16.106 drew for each document of tools/make_script_probe.py: "
                       "the baselines in 1/300-inch device px of each page, paragraph by paragraph (the "
                       "text is the generator's); per group (per line in the size sweep), every "
                       "distinct [drawn offset from the baseline in px, size in half points] of its "
                       "extra runs; and the four hhea/typo integers of every face the "
                       "model asked for. Measurements only; regenerate with "
                       "tools/read_script_probe.py --record."),
            "faces": dict(sorted(recording.faces.items())),
            "documents": documents,
            "scripts": scripts,
        }
        OBSERVATIONS.write_text(json.dumps(payload, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
