"""Reusable content blocks: bullet rows, banners, native charts."""
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

from .builder import gradient_fill, shape_box, text_box
from .theme import Theme, hex_rgb

MARKERS = {"diamond": MSO_SHAPE.DIAMOND, "arrow": MSO_SHAPE.RIGHT_ARROW,
           "dot": MSO_SHAPE.OVAL, "square": MSO_SHAPE.RECTANGLE}


def bullets(slide, theme: Theme, items, x, y, w, style="diamond", size=22, row_h=None):
    """One marker shape + one text box per bullet (each reveals separately)."""
    row_h = row_h or Inches(0.5)
    for i, item in enumerate(items):
        ry = y + i * row_h
        m = Inches(0.24) if style != "arrow" else Inches(0.34)
        mh = Inches(0.24) if style != "arrow" else Inches(0.22)
        filled = style in ("arrow", "dot", "square")
        shape_box(slide, x, ry + (row_h - mh) / 2, m, mh,
                  fill=theme.primary if filled else None,
                  line=None if filled else theme.primary, kind=MARKERS.get(style, MSO_SHAPE.DIAMOND),
                  line_w=1.5)
        text_box(slide, x + m + Inches(0.15), ry, w - m - Inches(0.15), row_h, item, theme,
                 size=size, bold=True)
    return y + len(items) * row_h


def banner(slide, theme: Theme, text, x, y, w, h, style="text", size=26):
    if style == "bar":
        return text_box(slide, x, y, w, h, text, theme, size=size, bold=True, color=theme.white,
                        align="center", fill=theme.primary)
    if style == "gradient":
        box = shape_box(slide, x, y, w, h, fill=theme.primary)
        gradient_fill(box, theme.gradient or [theme.soft, theme.primary, theme.soft], 0)
        if text:
            text_box(slide, x, y, w, h, text, theme, size=size, bold=True, color=theme.white,
                     align="center")
        return box
    if style == "soft":
        return text_box(slide, x, y, w, h, text, theme, size=size, bold=True, color=theme.primary,
                        align="center", fill=theme.soft)
    return text_box(slide, x, y, w, h, text, theme, size=size, bold=True, color=theme.accent,
                    align="center")


CHARTS = {"line": XL_CHART_TYPE.LINE_MARKERS, "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
          "barh": XL_CHART_TYPE.BAR_CLUSTERED, "pie": XL_CHART_TYPE.PIE,
          "doughnut": XL_CHART_TYPE.DOUGHNUT, "scatter": XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
          "area": XL_CHART_TYPE.AREA}
PALETTE_EXTRA = ["3B7DD8", "E0A030", "2E9E6B", "8E8E8E", "D9534F"]


def chart(slide, theme: Theme, spec: dict, x, y, w, h):
    """Native, editable PowerPoint chart.

    spec: {type: line|bar|barh|pie|doughnut|scatter|area, categories: [...],
           series: {name: [values]} or for scatter {name: [[x, y], ...]},
           title, x_title, y_title, legend: bool}
    """
    kind = spec.get("type", "bar")
    if kind == "scatter":
        data = XyChartData()
        for name, pts in spec["series"].items():
            s = data.add_series(name)
            for px, py in pts:
                s.add_data_point(px, py)
    else:
        data = CategoryChartData()
        data.categories = spec.get("categories", [])
        for name, vals in spec["series"].items():
            data.add_series(name, vals)
    frame = slide.shapes.add_chart(CHARTS.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED),
                                   Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)), data)
    ch = frame.chart
    ch.font.size = Pt(spec.get("font_size", 11))
    ch.font.name = theme.font_en
    colors = [theme.primary, theme.accent] + PALETTE_EXTRA
    if kind in ("pie", "doughnut"):
        for i, point in enumerate(ch.plots[0].series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = hex_rgb(colors[i % len(colors)])
    else:
        for i, s in enumerate(ch.plots[0].series):
            col = hex_rgb(colors[i % len(colors)])
            if kind in ("line", "scatter"):
                s.format.line.color.rgb = col
                s.format.line.width = Pt(2.25)
                s.smooth = kind == "scatter"
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = col
        if kind in ("bar", "barh"):
            ch.plots[0].gap_width = 60
        try:
            ch.value_axis.has_major_gridlines = False
            ch.value_axis.format.line.color.rgb = hex_rgb("808080")
            ch.category_axis.format.line.color.rgb = hex_rgb("808080")
            for ax, key in ((ch.category_axis, "x_title"), (ch.value_axis, "y_title")):
                if spec.get(key):
                    ax.has_title = True
                    ax.axis_title.text_frame.text = spec[key]
                    ax.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
        except (AttributeError, ValueError):
            pass
    many = len(spec["series"]) > 1 or kind in ("pie", "doughnut")
    ch.has_legend = spec.get("legend", many)
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    if spec.get("title"):
        ch.has_title = True
        ch.chart_title.text_frame.text = spec["title"]
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(13)
    else:
        ch.has_title = False
    return frame
