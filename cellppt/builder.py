# ==============================================================================
# 【文件作用】PPT 的“底层画笔工具箱” (cellppt/builder.py)
# 所有在幻灯片上画矩形、圆角框、写带颜色的富文本（包括化学式上下标、深红高亮词）、
# 涂渐变色、无损插图片、画箭头连接线的操作，全都在这里封装。
# ==============================================================================

import re                                             # 导入正则表达式模块，用于解析 **高亮** 和 _{下标}
from pathlib import Path                              # 导入路径处理模块

from lxml import etree                                # 导入 XML 处理库，用于直接操作 PowerPoint 底层 DrawingML 节点
from PIL import Image                                 # 导入图片处理库，用于读取图片的原始宽高比例
from pptx.enum.shapes import MSO_SHAPE                # 导入 PPT 内置几何形状枚举（矩形、圆角矩形、椭圆、箭头等）
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN       # 导入文字水平对齐与垂直居中枚举
from pptx.oxml.ns import qn                           # 导入 XML 命名空间转换工具
from pptx.util import Emu, Pt                         # 导入 PPT 尺寸单位：Emu(坐标长度) 和 Pt(字号/线宽磅值)

from .theme import Theme, hex_rgb                     # 导入主题皮肤类与颜色转换函数

# 文字水平对齐方式映射表（left=左对齐, center=居中, right=右对齐）
ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
# 文本框垂直对齐方式映射表（top=靠上, middle=垂直居中, bottom=靠下）
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}

# 富文本特殊标记正则表达式：
#   **文字**  -> 变成高亮强调色（如深红）并加粗
#   ==文字==  -> 变成主题主色调（如深紫）并加粗
#   _{文字}   -> 变成下标（如 C_{3} 变成 C₃）
#   ^{文字}   -> 变成上标（如 10^{4} 变成 10⁴）
TOKEN = re.compile(r"(\*\*.+?\*\*|==.+?==|_\{.+?\}|\^\{.+?\})")


def set_run_font(run, theme: Theme, size=None, bold=None, color=None, font=None, italic=None):
    # 【函数作用】给一小段连续文字（run）设置字号、粗体、斜体、颜色，并同时绑定中英文字体
    f = run.font
    if size:
        f.size = Pt(size)               # 设置字号（单位：磅 Pt，后期微调文字大小就是改传入的 size）
    if bold is not None:
        f.bold = bold                   # 设置是否加粗（True / False）
    if italic is not None:
        f.italic = italic               # 设置是否斜体（True / False）
    if color:
        f.color.rgb = hex_rgb(color)    # 设置文字 RGB 颜色
    latin = theme.font_en if font is None else font
    f.name = latin                      # 设置英文与数字的字体（默认 Arial）
    rpr = run._r.get_or_add_rPr()       # 获取底层 XML 的字符属性节点 <a:rPr>
    # 遍历设置东亚中文字体(a:ea)与复杂脚本字体(a:cs)，防止中文掉回难看的系统默认宋体
    for tag, face in (("a:ea", font or theme.font_cn), ("a:cs", font or theme.font_cn)):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", face)        # 写入中文字体名（默认 Microsoft YaHei）


def _segments(text: str):
    # 【函数作用】把一行带有 **高亮** 或 _{下标} 的字符串切分成一段段小碎片，返回 (文字内容, 颜色风格, 上下标偏移值)
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            # **包裹的词**：去掉前后的 **，标记为 'accent'（深红强调色），基线偏移为 0
            yield part[2:-2], "accent", 0
        elif part.startswith("==") and part.endswith("=="):
            # ==包裹的词==：去掉前后的 ==，标记为 'primary'（主题主色），基线偏移为 0
            yield part[2:-2], "primary", 0
        elif part.startswith("_{"):
            # _{包裹的词}：去掉 _{ 和 }，基线偏移设为 -25000（PowerPoint 标准下标偏移量）
            yield part[2:-1], None, -25000
        elif part.startswith("^{"):
            # ^{包裹的词}：去掉 ^{ 和 }，基线偏移设为 +30000（PowerPoint 标准上标偏移量）
            yield part[2:-1], None, 30000
        else:
            # 普通文字：保持默认颜色，基线偏移为 0
            yield part, None, 0


def write_rich(tf, text, theme: Theme, size=16, color=None, bold=False, align="left",
               anchor="middle", line_spacing=1.1, font=None, wrap=True, autofit=False):
    # 【函数作用】往一个 PPT 文本框（TextFrame）里写入富文本，支持 '\n' 多行换行、行距控制与内边距微调
    tf.clear()                                                    # 先清空文本框里原有的默认空段落
    tf.word_wrap = wrap                                           # 开启自动换行（True=文字到框边缘自动折行）
    tf.vertical_anchor = ANCHOR.get(anchor, MSO_ANCHOR.MIDDLE)    # 设置垂直对齐（默认 middle 垂直居中）
    # 设置文本框左右内边距（45720 EMU 约等于 0.05 英寸，如果觉得字离框边太近可调大这个数）
    for side in ("margin_left", "margin_right"):
        setattr(tf, side, Emu(45720))
    # 设置文本框上下内边距（22860 EMU 约等于 0.025 英寸）
    tf.margin_top = tf.margin_bottom = Emu(22860)
    color = color or theme.text                                   # 未指定颜色时使用主题默认深灰字色
    # 按换行符 '\n' 把文本拆成多段
    for i, line in enumerate(str(text).split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph() # 第 1 行复用自带段落，后续行新建段落
        para.alignment = ALIGN.get(align, PP_ALIGN.LEFT)          # 设置水平对齐方式（left / center / right）
        para.line_spacing = line_spacing                          # 设置行距倍数（默认 1.1 倍行距，觉得行距挤可调成 1.25）
        # 遍历该行切分出来的每个富文本片段
        for chunk, style, baseline in _segments(line):
            run = para.add_run()                                  # 在段落末尾新建一个文字片段 run
            run.text = chunk                                      # 填入文字
            # 如果是 **高亮** 则取 theme.accent，如果是 ==主色== 则取 theme.primary，否则用普通字色
            c = {"accent": theme.accent, "primary": theme.primary}.get(style, color)
            # 设置该片段的字号、粗体（高亮词自动加粗）、颜色和字体
            set_run_font(run, theme, size, bold or style is not None, c, font)
            # 如果有上下标偏移（baseline != 0），直接写入底层 XML 的 baseline 属性
            if baseline:
                run._r.get_or_add_rPr().set("baseline", str(baseline))
    # 如果开启了 autofit=True，则在文字超出框大小时让 PowerPoint 自动缩小字号适应框体
    if autofit:
        body = tf._txBody.find(qn("a:bodyPr"))
        for child in list(body):
            if child.tag in (qn("a:spAutoFit"), qn("a:noAutofit"), qn("a:normAutofit")):
                body.remove(child)
        etree.SubElement(body, qn("a:normAutofit"))
    return tf


def no_line(shape):
    # 【函数作用】去除形状的外边框线（让它变成纯色块，不带默认细黑边）
    shape.line.fill.background()


def shape_box(slide, x, y, w, h, fill=None, line=None, kind=MSO_SHAPE.RECTANGLE,
              line_w=1.0, radius=None, name=None, shadow=False):
    # 【函数作用】在幻灯片坐标 (x, y) 处画一个宽 w、高 h 的几何图形（背景框、圆角卡片、箭头、圆环等都靠它画）
    # 参数说明：
    #   x, y   : 左上角横纵坐标（可用 Inches(数值) 传参）
    #   w, h   : 宽度和高度
    #   fill   : 填充背景色（6位十六进制字符串，传 None 则透明无背景）
    #   line   : 边框线颜色（传 None 则无边框）
    #   kind   : 形状类型（默认矩形 RECTANGLE，圆角矩形是 ROUNDED_RECTANGLE，圆是 OVAL）
    #   line_w : 边框线粗细（单位 Pt，默认 1.0 磅）
    #   radius : 圆角弧度系数（0.0 ~ 1.0 之间，0.5 表示胶囊半圆角）
    shp = slide.shapes.add_shape(kind, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    if fill:
        shp.fill.solid()                         # 设置为纯色填充
        shp.fill.fore_color.rgb = hex_rgb(fill)  # 填入指定的 RGB 颜色
    else:
        shp.fill.background()                    # 无填充（透明）
    if line:
        shp.line.color.rgb = hex_rgb(line)       # 设置边框线颜色
        shp.line.width = Pt(line_w)              # 设置边框线粗细
    else:
        no_line(shp)                             # 无边框线
    # 如果传了圆角系数 radius 且该形状支持几何控制点，则写入控制点调整圆角大小
    if radius is not None and shp.adjustments and len(shp.adjustments) > 0:
        shp.adjustments[0] = radius
    # 默认关闭 PowerPoint 新建形状自带的难看阴影（只有 shadow=True 时才保留阴影）
    if not shadow:
        sppr = shp._element.spPr
        eff = etree.SubElement(sppr, qn("a:effectLst"))
        del eff
    if name:
        shp.name = name                          # 给形状起个名字，方便调试识别
    if shp.has_text_frame:
        shp.text_frame.text = ""                 # 初始化清空内部文字
    return shp


def text_box(slide, x, y, w, h, text, theme, size=16, color=None, bold=False, align="left",
             anchor="middle", fill=None, line=None, kind=None, name=None, font=None,
             line_spacing=1.1, autofit=False, radius=None):
    # 【函数作用】最常用的“画文字卡片/文本框”函数：一步完成“画框 + 填底色 + 写富文本”
    # 如果既没指定形状 kind，也没指定背景色 fill 和边框 line，就创建一个纯透明文本框
    if kind is None and fill is None and line is None:
        shp = slide.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    else:
        # 否则先调用 shape_box 画出带颜色/边框/圆角的几何形状
        shp = shape_box(slide, x, y, w, h, fill, line, kind or MSO_SHAPE.RECTANGLE, radius=radius)
    # 往形状的文本框里写入富文本内容
    write_rich(shp.text_frame, text, theme, size, color, bold, align, anchor, line_spacing,
               font, autofit=autofit)
    if name:
        shp.name = name
    return shp


def gradient_fill(shape, stops, angle=0):
    # 【函数作用】把一个形状的背景改成线性渐变色
    # 参数 stops : 颜色列表，如 ["F3E8F5", "6A1069", "F3E8F5"] 表示从浅紫 -> 深紫 -> 浅紫
    # 参数 angle : 渐变角度（0=从左到右水平渐变，90=从上到下垂直渐变）
    sppr = shape._element.spPr
    # 先移除形状原有的纯色或旧渐变节点
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)
    grad = etree.Element(qn("a:gradFill"), rotWithShape="1")  # 创建渐变节点 <a:gradFill>
    lst = etree.SubElement(grad, qn("a:gsLst"))               # 创建渐变色标列表 <a:gsLst>
    n = max(1, len(stops) - 1)
    for i, col in enumerate(stops):
        # pos 取值 0 ~ 100000（代表 0% ~ 100% 的位置）
        gs = etree.SubElement(lst, qn("a:gs"), pos=str(int(100000 * i / n)))
        etree.SubElement(gs, qn("a:srgbClr"), val=col.lstrip("#"))
    # PowerPoint 底层角度单位是 1/60000 度，因此乘以 60000
    etree.SubElement(grad, qn("a:lin"), ang=str(int(angle * 60000)), scaled="0")
    geom = sppr.find(qn("a:prstGeom"))
    # According to OOXML schema, <a:gradFill> must appear right after <a:prstGeom>
    (geom.addnext if geom is not None else sppr.append)(grad)
    return shape


def picture_fit(slide, path, x, y, w, h, align="center"):
    # 【函数作用】把本地图片（path）无损等比缩放到目标矩形框 (x, y, w, h) 内，自动居中，绝不拉伸变形
    path = Path(path)
    with Image.open(path) as im:
        iw, ih = im.size                         # 读取图片原始像素宽 iw、高 ih
    scale = min(w / iw, h / ih)                  # 取宽缩放比与高缩放比中的较小值，确保图片完整塞进框内
    pw, ph = iw * scale, ih * scale              # 计算缩放后的实际宽高
    px = x + (w - pw) / 2 if align == "center" else x  # 水平居中计算左上角 X
    py = y + (h - ph) / 2                        # 垂直居中计算左上角 Y
    return slide.shapes.add_picture(str(path), Emu(int(px)), Emu(int(py)), Emu(int(pw)), Emu(int(ph)))


def connector(slide, x1, y1, x2, y2, color, width=1.25, arrow=True):
    # 【函数作用】在起点 (x1, y1) 和终点 (x2, y2) 之间画一条直线；如果 arrow=True 则在终点加上三角箭头
    from pptx.enum.shapes import MSO_CONNECTOR
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)),
                                      Emu(int(x2)), Emu(int(y2)))
    line.line.color.rgb = hex_rgb(color)         # 设置线条颜色
    line.line.width = Pt(width)                  # 设置线条粗细（默认 1.25 磅）
    if arrow:
        ln = line.line._get_or_add_ln()
        # 在线条末端添加三角形箭头 <a:tailEnd type="triangle"/>
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return line
