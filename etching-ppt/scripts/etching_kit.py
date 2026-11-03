"""etching_kit: build monochrome "etching" style PowerPoint decks with python-pptx.

The look: warm off-white paper, ink-and-wash illustrations, spaced-caps serif titles, italic
epigraph subtitles, Roman-numeral panel labels and page numbers, thin rules and a footer.

Typical use:

    from etching_kit import Deck
    d = Deck(deck_title="my deck")
    d.cover("A TITLE", "a subtitle", image="art/cover.jpg", meta=["Author", "Date"])
    d.panels("Three Ideas", "what matters", [{"label": "First", "body": "..."}, ...])
    d.save("out.pptx")

Every layout method returns the python-pptx slide, so callers can add shapes of their own.
"""
from __future__ import annotations

import hashlib
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

IN = 914400
ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


@dataclass
class Theme:
    paper: str = "#F7F5F0"
    ink: str = "#2B2B2B"
    ink2: str = "#555555"
    grey: str = "#BDBDBD"
    wash: str = "#E3E1DC"
    font: str = "EB Garamond"       # any installed serif; falls back to the viewer's default if missing
    width_in: float = 13.333
    height_in: float = 7.5
    margin: float = 0.6
    dash_open: str = "— "      # epigraph and caption framing; set both to "" to disable
    dash_close: str = " —"
    min_font: float = 11.0


def roman(n: int) -> str:
    out = ""
    for v, r in [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"),
                 (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]:
        while n >= v:
            out += r
            n -= v
    return out


def _rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#"))


def estimate_height(text: str, width_in: float, size: float, line: float = 1.2) -> float:
    """Rough height in inches of `text` wrapped into `width_in` (serif, ~0.42 em average glyph)."""
    cpl = max(1, int(width_in * 72 / (0.42 * size)))
    lines = sum(max(1, math.ceil(len(p) / cpl)) for p in text.split("\n"))
    return lines * size * line * 1.2 / 72


def fit_size(text: str, width_in: float, height_in: float, size: float, line: float, min_size: float) -> float:
    while size > min_size and estimate_height(text, width_in, size, line) > height_in:
        size -= 0.5
    return size


class Deck:
    def __init__(self, theme: Theme | None = None, deck_title: str = "", cache_dir: str | Path | None = None,
                 page_numbers: bool = True, warn=print):
        self.t = theme or Theme()
        self.deck_title = deck_title
        self.page_numbers = page_numbers
        self.warn = warn
        self.cache = Path(cache_dir or Path.cwd() / ".etching_cache")
        self.cache.mkdir(parents=True, exist_ok=True)
        self.prs = Presentation()
        self.prs.slide_width = Emu(int(self.t.width_in * IN))
        self.prs.slide_height = Emu(int(self.t.height_in * IN))
        self.W, self.H = self.t.width_in, self.t.height_in

    # primitives
    def _frame_text(self, t: str) -> str:
        return f"{self.t.dash_open}{t}{self.t.dash_close}"

    def text(self, s, x, y, w, h, t, size=18, italic=False, bold=False, color=None, align="left",
             spacing=0, anchor=MSO_ANCHOR.TOP, line=1.15, fit=False):
        color = color or self.t.ink
        if fit:
            new = fit_size(t, w, h, size, line, self.t.min_font)
            if estimate_height(t, w, new, line) > h:
                self.warn(f"[etching_kit] text may overflow its box ({w:.1f}x{h:.1f} in): {t[:50]!r}")
            size = new
        tb = s.shapes.add_textbox(Emu(int(x * IN)), Emu(int(y * IN)), Emu(int(w * IN)), Emu(int(h * IN)))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        for i, para in enumerate(t.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = ALIGN[align] if isinstance(align, str) else align
            p.line_spacing = line
            r = p.add_run()
            r.text = para
            f = r.font
            f.name, f.size, f.italic, f.bold = self.t.font, Pt(size), italic, bold
            f.color.rgb = _rgb(color)
            if spacing:
                r._r.get_or_add_rPr().set("spc", str(spacing))
        return tb

    def rule(self, s, x1, y, x2, color=None, w=0.75):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1 * IN)), Emu(int(y * IN)),
                                    Emu(int(x2 * IN)), Emu(int(y * IN)))
        ln.line.color.rgb = _rgb(color or self.t.ink)
        ln.line.width = Pt(w)

    def frame(self, s, x, y, w, h, fill=None, lw=0.9):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * IN)), Emu(int(y * IN)), Emu(int(w * IN)),
                               Emu(int(h * IN)))
        r.shadow.inherit = False
        if fill:
            r.fill.solid()
            r.fill.fore_color.rgb = _rgb(fill)
        else:
            r.fill.background()
        r.line.color.rgb = _rgb(self.t.ink)
        r.line.width = Pt(lw)
        return r

    def pic(self, s, path, x, y, w=None, h=None, center=True):
        """Place an image scaled to fit inside (w, h); returns the drawn (w, h)."""
        path = Path(path)
        if not path.exists():
            self.warn(f"[etching_kit] image not found, skipped: {path}")
            return 0, 0
        iw, ih = Image.open(path).size
        if w and h:
            sc = min(w / iw, h / ih)
            pw, ph = iw * sc, ih * sc
            if center:
                x, y = x + (w - pw) / 2, y + (h - ph) / 2
            w, h = pw, ph
        elif w:
            h = w * ih / iw
        else:
            w = h * iw / ih
        s.shapes.add_picture(str(path), Emu(int(x * IN)), Emu(int(y * IN)), Emu(int(w * IN)), Emu(int(h * IN)))
        return w, h

    def equation_image(self, latex: str, size: float = 24) -> Path:
        """Render LaTeX (matplotlib mathtext, Computer Modern) to a transparent PNG; lines split on \\n."""
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        key = hashlib.sha1(f"{latex}|{size}|{self.t.ink}".encode()).hexdigest()[:16]
        out = self.cache / f"eq_{key}.png"
        if out.exists():
            return out
        rows = latex.split("\n")
        fig = plt.figure(figsize=(9, 0.7 * len(rows)))
        for i, row in enumerate(rows):
            fig.text(0.5, 1 - (i + 0.5) / len(rows), f"${row}$", ha="center", va="center", fontsize=size,
                     color=self.t.ink, math_fontfamily="cm")
        fig.savefig(out, dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.08)
        plt.close(fig)
        return out

    # slide chrome
    def slide(self, title=None, subtitle=None, source_note=None, notes=None, title_size=30):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = _rgb(self.t.paper)
        n = len(self.prs.slides)
        m = self.t.margin
        if title:
            self.text(s, m, 0.38, self.W - 2 * m, 0.6, title.upper(), title_size, align="center", spacing=450)
        if subtitle:
            self.text(s, m, 0.98, self.W - 2 * m, 0.4, self._frame_text(subtitle), 15, italic=True,
                      color=self.t.ink2, align="center")
        self.rule(s, m, self.H - 0.55, self.W - m)
        if self.deck_title:
            self.text(s, m, self.H - 0.47, self.W - 2 * m, 0.3, self._frame_text(self.deck_title), 11, italic=True,
                      color=self.t.ink2, align="center")
        if self.page_numbers:
            self.text(s, self.W - m - 1.5, self.H - 0.88, 1.5, 0.3, self._frame_text(roman(n)), 12, italic=True,
                      color=self.t.ink2, align="right")
        if source_note:
            self.text(s, m, self.H - 0.88, self.W - 2 * m - 1.8, 0.3, source_note, 11, italic=True, color=self.t.ink2)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def _panel(self, s, x, y, w, h, label=None, body=None, caption=None, image=None, size=15, number=None):
        if label:
            lab = (f"{roman(number).upper()} · " if number else "") + label.upper()
            self.text(s, x, y, w, 0.3, lab, 12, align="center", spacing=150)
        self.frame(s, x, y + 0.32, w, h)
        if image:
            self.pic(s, image, x + 0.08, y + 0.4, w - 0.16, h - 0.16)
        elif body:
            self.text(s, x + 0.22, y + 0.5, w - 0.44, h - 0.3, body, size, line=1.2, fit=True)
        if caption:
            self.text(s, x, y + 0.42 + h, w, 0.35, self._frame_text(caption), 13, italic=True, color=self.t.ink2,
                      align="center")

    # layouts
    def cover(self, title, subtitle=None, image=None, meta: Sequence[str] = (), kicker=None, **kw):
        """Title, epigraph, one wide panorama above a baseline rule, then author lines."""
        s = self.slide(**kw)
        m = self.t.margin
        if kicker:
            self.text(s, m, 0.12, 6, 0.25, kicker, 10, italic=True, color=self.t.ink2)
        self.text(s, m, 0.45, self.W - 2 * m, 0.5, title.upper(), 28, align="center", spacing=350)
        if subtitle:
            self.text(s, m, 1.05, self.W - 2 * m, 0.4, self._frame_text(subtitle), 17, italic=True,
                      color=self.t.ink2, align="center")
        if image:
            self.pic(s, image, 1.4, 1.6, self.W - 2.8, 3.8)
        self.rule(s, 1.4, 5.5, self.W - 1.4, w=1.2)
        if meta:
            self.text(s, m, 5.62, self.W - 2 * m, 1.2, "\n".join(meta), 14, italic=True, color=self.t.ink2,
                      align="center", line=1.3)
        return s

    def section(self, label, title, items: Sequence[str] = (), image=None, **kw):
        """Divider slide: small spaced label, big title, a short rule, italic contents list.

        With `image`, the text block moves to the left half and the picture fills the right half.
        """
        s = self.slide(**kw)
        m = self.t.margin
        if image:
            tx, tw = 0.6, 6.4
            self.pic(s, image, 7.3, 1.6, 5.2, 4.4)
        else:
            tx, tw = m, self.W - 2 * m
        cx = tx + tw / 2
        self.text(s, tx, 1.6, tw, 0.5, label.upper(), 18, color=self.t.ink2, align="center", spacing=600)
        self.text(s, tx, 2.15, tw, 0.8, title.upper(), (34 if len(title) <= 12 else 26) if image else 40, align="center",
                  spacing=450)
        self.rule(s, cx - 2.1, 3.15, cx + 2.1)
        if items:
            ix, iw = (tx, tw) if image else (1.5, self.W - 3.0)
            self.text(s, ix, 3.4, iw, 2.8, "\n".join(items), 16, italic=True, color=self.t.ink2,
                      align="center", line=1.35)
        return s

    def panels(self, title, subtitle, panels: Sequence[dict], cols: int | None = None, caption=None,
               numbered=True, **kw):
        """Grid of framed panels. Each panel: {label, body | image, caption}. 1-3 per row."""
        s = self.slide(title, subtitle, **kw)
        n = len(panels)
        cols = cols or min(n, 3)
        rows = math.ceil(n / cols)
        gap, x0 = 0.37, 0.9
        w = (self.W - 2 * x0 - gap * (cols - 1)) / cols
        top, bottom = 1.55, (5.55 if caption else 6.1)
        h = (bottom - top - 0.8 * rows) / rows
        if not any(p.get("image") for p in panels):
            h = min(h, 2.6)  # text-only panels stay compact
        for i, p in enumerate(panels):
            r, c = divmod(i, cols)
            self._panel(s, x0 + c * (w + gap), top + r * (h + 0.8), w, h, p.get("label"), p.get("body"),
                        p.get("caption"), p.get("image"), p.get("size", 15), i + 1 if numbered else None)
        if caption:
            self.text(s, 1.3, 5.6, self.W - 2.6, 0.8, caption, 16, align="center", line=1.25, fit=True)
        return s

    def text_image(self, title, subtitle, paragraphs: str, image=None, image_caption=None, image_label=None,
                   image_side="left", **kw):
        """Framed illustration (about 40% of the width) beside a column of text."""
        s = self.slide(title, subtitle, **kw)
        iw, tw = 4.3, 7.1
        ix, tx = (0.75, 5.5) if image_side == "left" else (self.W - 0.75 - iw, 0.9)
        if image:
            self._panel(s, ix, 1.55, iw, 3.3, image_label, None, image_caption, image)
        self.text(s, tx, 1.7, tw, 4.6, paragraphs, 16, line=1.25, fit=True)
        return s

    def equation(self, title, subtitle, equations: Sequence[dict], paragraph=None, **kw):
        """Framed equation blocks. Each: {latex, caption, size}. Keep it to one or two blocks."""
        s = self.slide(title, subtitle, **kw)
        y = 1.55
        avail = 3.9 if paragraph else 4.8
        each = avail / max(1, len(equations))
        for eq in equations:
            img = self.equation_image(eq["latex"], eq.get("size", 24))
            iw, ih = Image.open(img).size
            maxw, maxh = self.W - 3.4, each - 0.55
            sc = min(maxw / (iw / 300), maxh / (ih / 300), 1.6)
            w, h = iw / 300 * sc, ih / 300 * sc
            bx, bw, bh = (self.W - (w + 0.6)) / 2, w + 0.6, h + 0.3
            self.frame(s, bx, y, bw, bh)
            self.pic(s, img, bx + 0.3, y + 0.15, w, h, center=False)
            if eq.get("caption"):
                self.text(s, 1.4, y + bh + 0.06, self.W - 2.8, 0.35, self._frame_text(eq["caption"]), 13, italic=True,
                          color=self.t.ink2, align="center")
            y += bh + 0.55
        if paragraph:
            self.text(s, 1.6, max(y, 5.0), self.W - 3.2, 1.2, paragraph, 16, align="center", line=1.25, fit=True)
        return s

    def numbered(self, title, subtitle, items: Sequence[str], **kw):
        """Roman-numbered statements separated by thin grey rules (up to about six)."""
        s = self.slide(title, subtitle, **kw)
        step = min(0.9, 4.9 / max(1, len(items)))
        x, w = 1.3, self.W - 2.6
        for k, q in enumerate(items):
            y = 1.55 + k * step
            self.text(s, x, y, 0.7, 0.5, roman(k + 1).upper() + ".", 18, align="right")
            self.text(s, x + 0.95, y + 0.03, w - 0.95, step - 0.1, q, 16, line=1.25, fit=True)
            self.rule(s, x + 0.95, y + step - 0.12, x + w, self.t.grey, 0.5)
        return s

    def table(self, title, subtitle, header: Sequence[str], rows: Sequence[Sequence[str]],
              widths: Sequence[float] | None = None, size=14, **kw):
        """Ruled table: spaced-caps header, italic last column."""
        s = self.slide(title, subtitle, **kw)
        total = self.W - 1.5
        widths = widths or [total / len(header)] * len(header)
        scale = total / sum(widths)
        widths = [w * scale for w in widths]
        rh = min(0.7, 4.8 / (len(rows) + 1))
        x0, y0 = 0.75, 1.55
        for i, r in enumerate([header, *rows]):
            x = x0
            for j, c in enumerate(r):
                self.frame(s, x, y0 + i * rh, widths[j], rh, lw=0.6)
                self.text(s, x + 0.08, y0 + i * rh + 0.04, widths[j] - 0.16, rh - 0.08, c.upper() if i == 0 else c,
                          11 if i == 0 else size, italic=(i > 0 and j == len(r) - 1), anchor=MSO_ANCHOR.MIDDLE,
                          align="center" if i == 0 or (c.replace(".", "").replace(",", "").replace("%", "").isdigit() or c in ("inf", "∞")) else "left", spacing=100 if i == 0 else 0, line=1.05)
                x += widths[j]
        return s

    def figure(self, title, subtitle, image, caption=None, paragraph=None, **kw):
        """One wide image (diagram, chart, panorama) with an italic caption and an optional paragraph."""
        s = self.slide(title, subtitle, **kw)
        h = 3.4 if paragraph else 4.4
        self.pic(s, image, 0.9, 1.5, self.W - 1.8, h)
        if caption:
            self.text(s, 1.0, 1.55 + h, self.W - 2.0, 0.35, self._frame_text(caption), 13, italic=True,
                      color=self.t.ink2, align="center")
        if paragraph:
            self.text(s, 1.3, 5.45, self.W - 2.6, 1.0, paragraph, 16, align="center", line=1.25, fit=True)
        return s

    def closing(self, title="Thank You", subtitle=None, image=None, paragraph=None, **kw):
        """Ending slide: a receding-road style illustration, a rule, a closing paragraph."""
        s = self.slide(title, subtitle, **kw)
        if image:
            self.pic(s, image, 2.0, 1.55, self.W - 4.0, 3.25)
        self.rule(s, 3.0, 4.95, self.W - 3.0)
        if paragraph:
            self.text(s, 2.2, 5.1, self.W - 4.4, 1.2, paragraph, 18, align="center", line=1.25, fit=True)
        return s

    def save(self, path):
        self.prs.save(path)
        return path


# ink-style helpers (matplotlib)
def _ink_axes(theme: Theme, figsize):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "serif", "text.color": theme.ink, "axes.edgecolor": theme.ink,
                         "axes.labelcolor": theme.ink, "xtick.color": theme.ink, "ytick.color": theme.ink,
                         "mathtext.fontset": "cm", "hatch.linewidth": 0.7})
    fig, ax = plt.subplots(figsize=figsize)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    return plt, fig, ax


def chart_bars(path, labels: Sequence[str], values: Sequence[float], theme: Theme | None = None,
               ylabel: str = "", figsize=(9, 4.2)):
    """Bar chart in ink: diagonal hatching only, no colour."""
    theme = theme or Theme()
    plt, fig, ax = _ink_axes(theme, figsize)
    hatches = ["////", "\\\\\\\\", "xxxx", "....", "----"]
    for i, (lab, v) in enumerate(zip(labels, values)):
        ax.bar(i, v, facecolor="none", edgecolor=theme.ink, hatch=hatches[i % len(hatches)], lw=1.0)
        ax.text(i, v, f" {v:g}", ha="center", va="bottom", fontstyle="italic")
    ax.set_xticks(range(len(labels)), labels, fontstyle="italic")
    ax.set_ylabel(ylabel, fontstyle="italic")
    fig.savefig(path, dpi=220, transparent=True, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    return Path(path)


def chart_lines(path, series: dict, xlabel: str = "", ylabel: str = "", theme: Theme | None = None,
                figsize=(9, 4.2)):
    """Line chart in ink: solid, dashed, dotted lines to tell series apart; direct labels at line ends."""
    theme = theme or Theme()
    plt, fig, ax = _ink_axes(theme, figsize)
    styles = ["-", (0, (5, 3)), (0, (1, 2)), (0, (5, 2, 1, 2))]
    for i, (name, ys) in enumerate(series.items()):
        ax.plot(range(len(ys)), ys, color=theme.ink, lw=1.1, ls=styles[i % len(styles)])
        ax.text(len(ys) - 1, ys[-1], f"  {name}", va="center", fontstyle="italic")
    ax.set_xlabel(xlabel, fontstyle="italic")
    ax.set_ylabel(ylabel, fontstyle="italic")
    fig.savefig(path, dpi=220, transparent=True, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    return Path(path)


def placeholder_illustration(path, kind: str = "skyline", theme: Theme | None = None, seed: int = 3,
                             size=(1920, 960)):
    """Procedural pen-and-wash stand-in so a deck never ships with an empty frame.

    Replace it with a real illustration (see references/illustration_prompts.md) before presenting.
    """
    import numpy as np
    theme = theme or Theme()
    plt, fig, ax = _ink_axes(theme, (size[0] / 200, size[1] / 200))
    fig.patch.set_facecolor(theme.paper)
    rng = np.random.default_rng(seed)
    ax.axis("off")
    if kind == "skyline":
        x = 0.0
        while x < 10:
            w, h = rng.uniform(0.4, 1.0), rng.uniform(1.0, 4.2)
            ax.add_patch(plt.Rectangle((x, 0), w, h, facecolor=theme.wash, edgecolor=theme.ink, lw=0.9,
                                       hatch=rng.choice(["////", "\\\\\\\\", "...."])))
            x += w + rng.uniform(0.05, 0.2)
        ax.plot([0, 10], [0, 0], color=theme.ink, lw=1.4)
        ax.set_xlim(0, 10)
        ax.set_ylim(-0.2, 4.6)
    else:  # hills
        xs = np.linspace(0, 10, 200)
        for k, a in enumerate([2.4, 1.6, 0.9]):
            ys = a + 0.6 * np.sin(xs * (0.7 + 0.3 * k) + k) + 0.15 * rng.normal(size=1)
            ax.fill_between(xs, 0, ys, facecolor=theme.wash, edgecolor=theme.ink, lw=0.9, hatch="////" if k == 1 else None,
                            alpha=1 - 0.2 * k)
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 3.6)
    fig.savefig(path, dpi=200, facecolor=theme.paper, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    return Path(path)
