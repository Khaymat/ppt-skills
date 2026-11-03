# Layout reference

Every layout is a method on `etching_kit.Deck` and a `"layout"` value in a JSON spec. Keys below are the JSON keys.
Common keys on every slide: `title`, `subtitle`, `source_note`, `notes`.

Image values: a path relative to the spec, `"placeholder:skyline"`, `"placeholder:hills"`, or a chart object.

## cover
`title`, `subtitle`, `image` (wide panorama, about 2:1), `meta` (list of lines under the rule), `kicker` (tiny label
top-left, for example "draft for review").

## section
`label` ("Part I"), `title`, `items` (italic contents list), optional `image` (about 4:3; the text block then moves to the left half).

## panels
`panels`: list of `{"label", "body", "caption", "image", "size"}`. Use `body` for text, `image` for an illustration;
if both are given the image wins. Labels get Roman numerals automatically (`"numbered": false` to turn this off).
`cols` defaults to `min(n, 3)`. `caption` adds one sentence under the grid. Text-only panels are kept compact.

## text_image
`paragraphs` (use `\n` for new paragraphs), `image`, `image_label`, `image_caption`, `image_side` (`left` or `right`).

## equation
`equations`: list of `{"latex", "caption", "size"}`. LaTeX uses matplotlib mathtext (no `\begin{align}`; put several
lines in one `latex` string separated by `\n`). `paragraph` adds a sentence below.

## numbered
`items`: up to about six statements. Roman numerals and thin rules are added.

## table
`header`, `rows` (list of lists of strings), `widths` (relative column widths), `size`. The last column is italic;
numeric cells are centred.

## figure
`image` (diagram, chart or panorama), `caption`, `paragraph`.

Chart objects:

```json
{"chart": "bars",  "labels": ["A", "B"], "values": [4, 7], "ylabel": "units"}
{"chart": "lines", "series": {"first": [1, 2, 3], "second": [2, 2, 2]}, "xlabel": "t", "ylabel": "y"}
```

## closing
`image`, `paragraph`.

## Python equivalents

```python
from etching_kit import Deck, Theme
d = Deck(Theme(font="Georgia"), deck_title="my deck")
s = d.panels("Three Ideas", "subtitle", [{"label": "One", "body": "..."}], notes="say this aloud")
d.text(s, 1, 6, 5, 0.4, "extra line", 14, italic=True)   # add your own shapes
d.save("out.pptx")
```

Geometry is in inches on a 13.333 x 7.5 slide. Content area: x from 0.9 to 12.4, y from 1.5 to 6.5.
