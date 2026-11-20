# ==============================================================================
# 【文件作用】PPT 正文页的高频“积木块组件” (cellppt/blocks.py)
# 包含三大积木：
#   1. bullets : 标题栏下方的带图标要点行（菱形/箭头/圆点 + 富文本句子）
#   2. banner  : 页面最底下的总结横幅条（实心紫底白字条、渐变底条、或纯红字结论条）
#   3. chart   : PowerPoint 原生可编辑的数据图表（柱状图、折线图、饼图、圆环图、散点图）
# ==============================================================================

from pptx.chart.data import CategoryChartData, XyChartData     # 导入图表数据结构
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION  # 导入图表类型与图例位置枚举
from pptx.enum.shapes import MSO_SHAPE                         # 导入几何形状枚举
from pptx.util import Emu, Inches, Pt                          # 导入尺寸单位

from .builder import gradient_fill, shape_box, text_box        # 导入底层画笔
from .theme import Theme, hex_rgb                              # 导入主题与颜色工具

# 要点行前面的小图标形状映射表：diamond=菱形, arrow=右箭头, dot=圆点, square=方块
MARKERS = {
    "diamond": MSO_SHAPE.DIAMOND,
    "arrow": MSO_SHAPE.RIGHT_ARROW,
    "dot": MSO_SHAPE.OVAL,
    "square": MSO_SHAPE.RECTANGLE,
}


def bullets(slide, theme: Theme, items, x, y, w, style="diamond", size=22, row_h=None):
    # 【函数作用】在坐标 (x, y) 处往下画多行核心要点列表
    # 每一行都拆成“1个独立小图形图标 + 1个独立文本框”，这样在 PowerPoint 逐元素回放时图标和文字会依次跳出来
    # 参数微调说明：
    #   style : 图标形状（"diamond"=空心紫框菱形, "arrow"=实心紫箭头, "dot"=实心圆点）
    #   size  : 要点文字字号（默认 22 磅，字多时可微调为 19 或 20）
    #   row_h : 每行占用的高度（默认 0.5 英寸，如果某行有两行文字可调大到 0.75）
    row_h = row_h or Inches(0.5)
    for i, item in enumerate(items):
        ry = y + i * row_h                                         # 计算第 i 行的顶部 Y 坐标
        m = Inches(0.24) if style != "arrow" else Inches(0.34)     # 小图标的宽度
        mh = Inches(0.24) if style != "arrow" else Inches(0.22)    # 小图标的高度
        filled = style in ("arrow", "dot", "square")               # 判断图标是实心填充还是空心描边
        # 1. 画要点前面的小图标
        shape_box(slide, x, ry + (row_h - mh) / 2, m, mh,
                  fill=theme.primary if filled else None,
                  line=None if filled else theme.primary,
                  kind=MARKERS.get(style, MSO_SHAPE.DIAMOND),
                  line_w=1.5)
        # 2. 画小图标右边的富文本要点文字（与图标间隔 0.15 英寸）
        text_box(slide, x + m + Inches(0.15), ry, w - m - Inches(0.15), row_h, item, theme,
                 size=size, bold=True)
    return y + len(items) * row_h  # 返回所有要点画完后的底部 Y 坐标


def banner(slide, theme: Theme, text, x, y, w, h, style="text", size=26):
    # 【函数作用】在页面底部画一句话总结横幅条（take-home message）
    # 参数 style 可选值：
    #   "bar"      : 深紫色实心背景 + 白色加粗文字（如第 4、8、11 页底栏）
    #   "gradient" : 左右浅紫向中间深紫过渡的渐变条 + 白色文字
    #   "soft"     : 浅紫色背景 + 深紫色文字
    #   "text"     : 无背景框、直接写深红色大号总结字（如第 9 页底部红字）
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


# PowerPoint 原生图表类型映射表
CHARTS = {
    "line": XL_CHART_TYPE.LINE_MARKERS,                  # 带数据点的折线图
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,               # 竖向簇状柱形图
    "barh": XL_CHART_TYPE.BAR_CLUSTERED,                 # 横向簇状条形图
    "pie": XL_CHART_TYPE.PIE,                            # 饼图
    "doughnut": XL_CHART_TYPE.DOUGHNUT,                  # 圆环图
    "scatter": XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,# 平滑曲线散点图
    "area": XL_CHART_TYPE.AREA,                          # 面积图
}
# 多系列图表的补充调色板（蓝、金黄、翠绿、中灰、珊瑚红）
PALETTE_EXTRA = ["3B7DD8", "E0A030", "2E9E6B", "8E8E8E", "D9534F"]


def chart(slide, theme: Theme, spec: dict, x, y, w, h):
    # 【函数作用】在幻灯片 (x, y, w, h) 区域插入一个 PowerPoint 原生可编辑数据图表（在 PPT 里右键可“编辑数据”）
    kind = spec.get("type", "bar")  # 读取图表类型，默认柱状图 "bar"
    if kind == "scatter":
        # 如果是 XY 散点曲线图，按 (px, py) 坐标对填入数据
        data = XyChartData()
        for name, pts in spec["series"].items():
            s = data.add_series(name)
            for px, py in pts:
                s.add_data_point(px, py)
    else:
        # 如果是分类图表（柱状图/折线图/圆环图），填入横轴分类 categories 和各系列数值
        data = CategoryChartData()
        data.categories = spec.get("categories", [])
        for name, vals in spec["series"].items():
            data.add_series(name, vals)
    # 调用 python-pptx 在幻灯片上创建原生图表对象
    frame = slide.shapes.add_chart(CHARTS.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED),
                                   Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)), data)
    ch = frame.chart
    ch.font.size = Pt(spec.get("font_size", 11))  # 设置图表坐标轴与标签字号（默认 11 磅）
    ch.font.name = theme.font_en                  # 设置图表字体为英文字体 Arial
    colors = [theme.primary, theme.accent] + PALETTE_EXTRA  # 组合图表配色序列
    if kind in ("pie", "doughnut"):
        # 饼图与圆环图：按扇区逐个上色
        for i, point in enumerate(ch.plots[0].series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = hex_rgb(colors[i % len(colors)])
    else:
        # 柱状图与折线图：按数据系列逐个设置线条或柱子颜色
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
            ch.plots[0].gap_width = 60  # 设置柱子之间的间距比例（60% 让柱子显得更饱满）
        try:
            # 去掉纵轴默认杂乱的网格横线，并将横纵坐标轴线设为深灰色
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
    # 设置图例显示（多于 1 个系列或饼图/圆环图时默认显示在底部）
    many = len(spec["series"]) > 1 or kind in ("pie", "doughnut")
    ch.has_legend = spec.get("legend", many)
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    # 设置图表顶部标题
    if spec.get("title"):
        ch.has_title = True
        ch.chart_title.text_frame.text = spec["title"]
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(13)
    else:
        ch.has_title = False
    return frame
