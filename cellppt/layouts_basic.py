"""Full-page layouts: cover, agenda, section divider, thanks, and the shared header bar."""
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from .builder import connector, gradient_fill, shape_box, text_box
from .theme import HEADER_H, SLIDE_H, SLIDE_W, Theme

IN = Inches(1)


def header_bar(slide, theme: Theme, title: str, number=None):
    """Purple title bar with a white numbered disc, as in the reference deck."""
    shape_box(slide, 0, 0, SLIDE_W, HEADER_H, fill=theme.primary, name="HeaderBar")
    x = Inches(0.32)
    if number is not None:
        d = Inches(0.66)
        disc = text_box(slide, x, (HEADER_H - d) / 2, d, d, f"{int(number):02d}", theme,
                        size=20, bold=True, color=theme.primary, align="center",
                        fill=theme.white, kind=MSO_SHAPE.OVAL, name="HeaderNumber")
        disc.text_frame.margin_left = disc.text_frame.margin_right = 0
        x += d + Inches(0.18)
    text_box(slide, x, 0, SLIDE_W - x - Inches(0.3), HEADER_H, title, theme, size=28,
             bold=True, color=theme.white, name="HeaderTitle")


def cover(slide, theme: Theme, meta: dict, spec: dict):
    band_top = Inches(4.35)
    shape_box(slide, 0, band_top, SLIDE_W, SLIDE_H - band_top, fill=theme.primary, name="CoverBand")
    byline = spec.get("byline") or " | ".join(
        str(v) for v in (meta.get("authors"), meta.get("venue"), meta.get("year")) if v)
    if byline:
        y = Inches(0.62)
        for side in (-1, 1):
            cx = SLIDE_W / 2 + side * Inches(3.6)
            connector(slide, cx - Inches(1.3), y, cx + Inches(1.3), y, theme.line, 1.0, arrow=False)
            shape_box(slide, cx + side * Inches(1.3) - Inches(0.05), y - Inches(0.05), Inches(0.1),
                      Inches(0.1), fill=theme.primary, kind=MSO_SHAPE.DIAMOND)
        text_box(slide, Inches(2.5), Inches(0.35), SLIDE_W - Inches(5), Inches(0.55), byline,
                 theme, size=20, bold=True, color=theme.primary, align="center")
    title = spec.get("title") or meta.get("title", "")
    text_box(slide, Inches(0.6), Inches(1.05), SLIDE_W - Inches(1.2), Inches(2.75), title, theme,
             size=spec.get("title_size", 50), bold=True, color=theme.primary, align="center",
             line_spacing=1.05)
    tag = spec.get("tag") or meta.get("tag")
    if tag:
        w = Inches(3.4)
        text_box(slide, (SLIDE_W - w) / 2, band_top - Inches(0.3), w, Inches(0.6), tag, theme,
                 size=20, bold=True, color=theme.white, align="center", fill=theme.primary,
                 line=theme.white, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    info = spec.get("info") or meta.get("cover_info")
    if info:
        text_box(slide, Inches(1.6), band_top + Inches(0.75), SLIDE_W - Inches(3.2), Inches(1.9),
                 info, theme, size=16, color=theme.white, anchor="top", line_spacing=1.3)


def agenda(slide, theme: Theme, meta: dict, spec: dict):
    shape_box(slide, 0, Inches(4.35), SLIDE_W, SLIDE_H - Inches(4.35), fill=theme.primary)
    panel = shape_box(slide, Inches(1.4), Inches(0.35), Inches(10.5), Inches(6.55),
                      fill=theme.white, name="AgendaPanel", shadow=True)
    panel.shadow.inherit = True
    title = spec.get("title", "汇\n报\n目\n录")
    text_box(slide, Inches(3.1), Inches(0.8), Inches(1.3), Inches(5.6), title, theme, size=54,
             bold=True, color=theme.primary, align="center", font="SimSun", line_spacing=1.0)
    connector(slide, Inches(3.05), Inches(2.1), Inches(3.05), Inches(5.6), theme.line, 1.2, False)
    sub = text_box(slide, Inches(1.9), Inches(2.6), Inches(1.2), Inches(2.6), "CONTENTS", theme,
                   size=22, bold=True, color="D0D0D0", align="center")
    sub.rotation = 90
    items = spec.get("items", [])
    active = spec.get("active", 0)
    top, step = Inches(0.62), min(Inches(0.96), Inches(5.9) / max(1, len(items)))
    for i, item in enumerate(items):
        on = i == active
        y = top + i * step
        row = shape_box(slide, Inches(5.05), y, Inches(5.9), step - Inches(0.2),
                        fill=theme.softer if on else theme.white,
                        line=theme.primary if on else "E6E6E6", kind=MSO_SHAPE.ROUNDED_RECTANGLE,
                        radius=0.5, line_w=1.0)
        if on:
            gradient_fill(row, [theme.soft, theme.white], 0)
        c = theme.primary if on else "C8C8C8"
        text_box(slide, Inches(5.25), y, Inches(1.0), step - Inches(0.2), f"{i + 1:02d}", theme,
                 size=26, bold=True, color=c, align="center")
        connector(slide, Inches(6.35), y + Inches(0.15), Inches(6.35), y + step - Inches(0.35),
                  c, 1.0, False)
        text_box(slide, Inches(6.5), y, Inches(4.3), step - Inches(0.2), item, theme, size=24,
                 color=theme.primary if on else "BDBDBD")


def section(slide, theme: Theme, meta: dict, spec: dict):
    shape_box(slide, 0, 0, Inches(4.6), SLIDE_H, fill=theme.primary)
    text_box(slide, Inches(0.5), Inches(2.2), Inches(3.6), Inches(2.4),
             f"{int(spec.get('number', 1)):02d}", theme, size=110, bold=True,
             color=theme.white, align="center")
    text_box(slide, Inches(5.3), Inches(2.5), Inches(7.5), Inches(1.2), spec.get("title", ""),
             theme, size=40, bold=True, color=theme.primary)
    if spec.get("subtitle"):
        text_box(slide, Inches(5.3), Inches(3.7), Inches(7.5), Inches(1.4), spec["subtitle"],
                 theme, size=20, color=theme.muted, anchor="top")
    connector(slide, Inches(5.35), Inches(3.62), Inches(9.5), Inches(3.62), theme.primary, 2.0, False)


def thanks(slide, theme: Theme, meta: dict, spec: dict):
    band_top = Inches(4.75)
    shape_box(slide, 0, band_top, SLIDE_W, SLIDE_H - band_top, fill=theme.primary)
    top_line = spec.get("byline") or meta.get("short_title") or ""
    if top_line:
        y = Inches(0.95)
        connector(slide, Inches(2.0), y, Inches(4.1), y, theme.line, 1.0, False)
        connector(slide, Inches(9.2), y, Inches(11.3), y, theme.line, 1.0, False)
        text_box(slide, Inches(4.1), Inches(0.65), Inches(5.1), Inches(0.6), top_line, theme,
                 size=22, bold=True, color=theme.primary, align="center")
    text_box(slide, Inches(1), Inches(1.45), SLIDE_W - Inches(2), Inches(3.0),
             spec.get("title", "感谢聆听\n欢迎交流讨论"), theme, size=spec.get("title_size", 60),
             bold=True, color=theme.primary, align="center", line_spacing=1.1)
    tag = spec.get("tag") or meta.get("tag_short") or "文献汇报"
    w = Inches(2.8)
    text_box(slide, (SLIDE_W - w) / 2, band_top - Inches(0.3), w, Inches(0.6), tag, theme,
             size=22, bold=True, color=theme.white, align="center", fill=theme.primary,
             line=theme.white, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
