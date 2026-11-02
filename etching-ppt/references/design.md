# Etching style: design specification

## 1. Character

Like a page from an illustrated essay or a classic magazine: calm, monochrome, generous empty space, pen-and-ink
illustrations with a soft grey wash. Academic but narrative. There are no bright colours, gradients, flat icons,
stock photos, drop shadows or emoji.

## 2. Palette

| Role | Colour |
|---|---|
| Paper (background) | `#F7F5F0` |
| Primary text, ink lines | `#2B2B2B` |
| Secondary text, captions, footer | `#555555` |
| Wash and hatching fills | `#BDBDBD` to `#E3E1DC` |
| Frame lines | `#2B2B2B`, 0.6 to 1.2 pt |

The only accent is contrast between ink and paper. To mark one thing as special (for example "the proposed
model"), use a wash fill or diagonal hatching.

## 3. Typography

- Classic serif throughout: EB Garamond, Cormorant Garamond or Libre Caslon. Georgia as a safe fallback.
- **Title**: capitals, wide letter-spacing (about 0.15 em), 28 to 40 pt, centred at the top.
- **Subtitle**: small italic, framed by dashes: `— the same blocks, marked by a different cartographer —`.
- **Panel label**: small capitals with a Roman numeral: `I · FIRST STEP`.
- **Caption under an image**: italic, framed by dashes.
- **Body**: regular serif, 14 to 18 pt, at most 3 to 5 short lines, essay tone.
- **Footer**: a thin full-width rule; the deck title in small italic at the centre; a lowercase Roman page number
  at the right, `— vii —`.

## 4. Layout principles

- Wide margins (about 6% each side). Everything is centred or aligned to a clear column.
- Illustrations always sit in a thin frame or float with soft edges.
- One idea per slide. If a slide needs more than five lines, split it or move detail to speaker notes.
- Avoid repeating the same layout more than twice in a row.

## 5. Illustrations

Pen-and-ink etching with a grey watercolour wash, objects that fit the topic (a ticker tape for markets, a
notebook of formulas for theory, a scale for trade-offs, a road for a future, a window for perspective).
No colour, no text inside the image. See `illustration_prompts.md`.

## 6. Equations and data

- Equations are rendered as transparent PNGs from LaTeX (Computer Modern), ink colour, inside a thin frame.
  At most one core block per slide; the caption explains it in words.
- Charts use hatching and line styles instead of colour, direct labels instead of a legend, and a frame
  without top and right spines.

## 7. Tone of text

Titles are short phrases, not sentences. Subtitles read like epigraphs. The body stays accurate to the source
material: the style may be poetic, the claims may not. Flag illustrative numbers as examples and show the source
of real ones in the `source_note`.

## 8. Accessibility and delivery

- Keep body text at 12 pt or larger; contrast of `#2B2B2B` on `#F7F5F0` is above 12:1.
- Give every illustration a meaningful caption, since the images carry no embedded text.
- Install the font on the presenting machine, or export to PDF first.
