---
name: etching-ppt
description: Build PowerPoint decks in a calm monochrome "etching" style (off-white paper, pen-and-ink illustrations, spaced-caps serif titles, Roman-numeral labels). Use when the user asks for slides, a deck or a presentation in an etching, sketchbook, New Yorker or book-essay look, or wants an academic talk that reads like an illustrated essay. Produces a .pptx from a JSON spec with python-pptx and checks the result by rendering it.
---

# etching-ppt

Turns a slide outline into a `.pptx` with one consistent visual language. The look is quiet and academic:
ink on warm paper, no colour, no icons, no stock photos, lots of empty space.

## When to use

- The user wants a presentation, deck or slides and names this style, or asks for something "calm", "classic",
  "book-like", "sketchbook" or "etching".
- A thesis, seminar, lecture or research talk where content accuracy matters more than flash.

Do not use it for colourful marketing decks or data-dense dashboards.

## Workflow

1. **Outline first.** One claim per slide. Title as a short noun phrase, subtitle as an italic epigraph. Body text
   is at most 3 to 5 short lines. Move detail into the speaker notes (`"notes"`), not onto the slide.
2. **Pick a layout per slide** from the table below. Mix them; avoid more than two identical layouts in a row.
3. **Illustrations.** Use real etching-style images when the user has them (`references/illustration_prompts.md`
   has prompts for an image generator). Otherwise use `"placeholder:skyline"` or `"placeholder:hills"` and tell the
   user these are stand-ins.
4. **Write the spec** (JSON, see `examples/demo_spec.json`) and build:
   ```bash
   pip install -r requirements.txt
   python scripts/build_deck.py my_spec.json -o my_deck.pptx
   ```
5. **Look at it.** Render and inspect every slide before delivering:
   ```bash
   python scripts/render_check.py my_deck.pptx -o preview/
   ```
   Open `preview/contact_sheet.png`, then single slides that look crowded. Fix overflow by cutting words, not by
   shrinking below 12 pt. Watch the console for `text may overflow` warnings.
6. **Report honestly**: what was built, which images are placeholders, which fonts must be installed.

## Layouts

| Layout | Use for | Key fields |
|---|---|---|
| `cover` | first slide | `title`, `subtitle`, `image`, `meta` (author lines) |
| `section` | divider between parts | `label`, `title`, `items` |
| `panels` | 2 to 6 parallel ideas, steps, comparisons | `panels` [{`label`, `body` or `image`, `caption`}], `cols`, `caption` |
| `text_image` | one illustration plus a short argument | `image`, `paragraphs`, `image_label`, `image_caption`, `image_side` |
| `equation` | one core formula | `equations` [{`latex`, `caption`}], `paragraph` |
| `numbered` | questions, objectives, steps | `items` |
| `table` | comparisons and schedules | `header`, `rows`, `widths` |
| `figure` | a chart, diagram or panorama | `image`, `caption`, `paragraph` |
| `closing` | last slide | `image`, `paragraph` |

Every slide also accepts `title`, `subtitle`, `source_note` (small italic label bottom-left, good for source tags
such as "[data] Statistics Office, 2025") and `notes` (speaker notes). Full details: `references/layouts.md`.

Charts need no external file: give `image` as `{"chart": "bars", "labels": [...], "values": [...]}` or
`{"chart": "lines", "series": {"name": [...]}}`. They are drawn with hatching and line styles, never colour.

## Style rules (summary; full spec in `references/design.md`)

- Palette: paper `#F7F5F0`, ink `#2B2B2B`, secondary `#555555`, wash `#E3E1DC`. Nothing else. To mark "our proposal",
  use diagonal hatching or a wash fill, not a colour.
- Serif only (EB Garamond by default; Garamond, Georgia or Cormorant also work via `"font"`).
- Titles are spaced capitals; subtitles and captions are italic and framed by dashes, for example
  `— a subtitle —`. Page numbers are lowercase Roman numerals in the footer.
- Equations are rendered as images (LaTeX mathtext, Computer Modern), at most one block per slide.
- No emoji, gradients, shadows, bullet icons or stock photos.
- The poetry lives in titles and captions only. Claims, numbers and citations must match the user's source text.
  Never invent a figure to fill a layout; mark examples and placeholders as such.

## Fonts

The deck stores the font name, not the font. Install EB Garamond (free, Google Fonts) on the machine that opens
the file, or set `"font": "Georgia"` in the spec. Rendering for the preview needs the same font to look right.

## Extending

`scripts/etching_kit.py` is a plain Python library. For a one-off layout, build the slide through
`Deck.slide(title, subtitle)` and add shapes with `deck.text`, `deck.frame`, `deck.rule` and `deck.pic`.
If a new layout is useful in general, add a method to `Deck` and list it in `LAYOUTS` in `build_deck.py`.
