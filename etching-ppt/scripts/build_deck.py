#!/usr/bin/env python3
"""Build an etching-style .pptx from a JSON spec.

    python scripts/build_deck.py examples/demo_spec.json -o demo.pptx

Spec shape (see references/layouts.md for every layout and its keys):

    {
      "deck_title": "footer title",
      "font": "EB Garamond",
      "slides": [
        {"layout": "cover", "title": "...", "subtitle": "...", "image": "art/cover.jpg", "meta": ["..."]},
        {"layout": "panels", "title": "...", "subtitle": "...", "panels": [{"label": "...", "body": "..."}]}
      ]
    }

Image values may be: a path relative to the spec file, "placeholder:skyline", "placeholder:hills",
or an object {"chart": "bars", "labels": [...], "values": [...]} / {"chart": "lines", "series": {...}}.
"""
from __future__ import annotations

import argparse
import hashlib
import inspect
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from etching_kit import Deck, Theme, chart_bars, chart_lines, placeholder_illustration  # noqa: E402

LAYOUTS = ("cover", "section", "panels", "text_image", "equation", "numbered", "table", "figure", "closing")


def resolve_image(value, base: Path, cache: Path, theme: Theme):
    if value is None:
        return None
    if isinstance(value, dict):
        key = hashlib.sha1(json.dumps(value, sort_keys=True).encode()).hexdigest()[:12]
        out = cache / f"chart_{key}.png"
        kind = value["chart"]
        if kind == "bars":
            return chart_bars(out, value["labels"], value["values"], theme, value.get("ylabel", ""))
        if kind == "lines":
            return chart_lines(out, value["series"], value.get("xlabel", ""), value.get("ylabel", ""), theme)
        raise ValueError(f"unknown chart kind: {kind}")
    if isinstance(value, str) and value.startswith("placeholder:"):
        kind = value.split(":", 1)[1]
        return placeholder_illustration(cache / f"ph_{kind}.png", kind, theme)
    p = Path(value)
    return p if p.is_absolute() else base / p


def build(spec_path: Path, out: Path) -> Path:
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base = spec_path.parent
    theme = Theme(**spec.get("theme", {}))
    if "font" in spec:
        theme.font = spec["font"]
    cache = Path(spec.get("cache_dir", base / ".etching_cache"))
    deck = Deck(theme, spec.get("deck_title", ""), cache, page_numbers=spec.get("page_numbers", True))
    for i, sl in enumerate(spec["slides"], 1):
        sl = dict(sl)
        layout = sl.pop("layout")
        if layout not in LAYOUTS:
            raise SystemExit(f"slide {i}: unknown layout {layout!r}; choose from {', '.join(LAYOUTS)}")
        for k in ("image",):
            if k in sl:
                sl[k] = resolve_image(sl[k], base, cache, theme)
        for p in sl.get("panels", []):
            if "image" in p:
                p["image"] = resolve_image(p["image"], base, cache, theme)
        fn = getattr(deck, layout)
        allowed = set(inspect.signature(fn).parameters) | {"source_note", "notes"}
        extra = set(sl) - allowed
        if extra:
            raise SystemExit(f"slide {i} ({layout}): unexpected keys {sorted(extra)}")
        fn(**sl)
    deck.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=Path("deck.pptx"))
    a = ap.parse_args()
    out = build(a.spec, a.out)
    print(f"saved {out}")


if __name__ == "__main__":
    main()
