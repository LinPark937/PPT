# ==============================================================================
# 【文件作用】PPT 的“标准骨架页模板” (cellppt/layouts_basic.py)
# 负责绘制每一套学术汇报 PPT 都会用到的固定结构页：
#   1. header_bar : 所有正文页最顶上那条带圆形编号的紫色标题栏
#   2. cover      : 第 1 页封面页
#   3. agenda     : 第 2 页汇报目录页
#   4. section    : 章节过渡页
#   5. thanks     : 最后一页感谢聆听封底页
# ==============================================================================

from pptx.enum.shapes import MSO_SHAPE                        # 导入 PPT 内置几何形状枚举
from pptx.util import Inches, Pt                              # 导入英寸与磅值单位

from .builder import connector, gradient_fill, shape_box, text_box  # 导入底层画笔函数
from .theme import HEADER_H, SLIDE_H, SLIDE_W, Theme          # 导入全局页面宽高与标题栏高度

IN = Inches(1)  # 1 英寸常量简写


def header_bar(slide, theme: Theme, title: str, number=None):
    # 【函数作用】画正文页顶部的通栏标题条：左边一个白底圆形页码（如 03、04），右边白色大字标题
    # 1. 画顶部的全宽深色背景矩形条（宽 SLIDE_W=13.333英寸，高 HEADER_H=0.92英寸）
    shape_box(slide, 0, 0, SLIDE_W, HEADER_H, fill=theme.primary, name="HeaderBar")
    x = Inches(0.32)  # 圆形编号距离左边缘的起始 X 坐标（0.32 英寸）
    if number is not None:
        d = Inches(0.66)  # 圆形编号徽章的直径（0.66 英寸，想调大圆圈改这里）
        # 2. 画白底圆形页码徽章（如 "03"），字号 20 磅加粗，颜色为主题主色
        disc = text_box(slide, x, (HEADER_H - d) / 2, d, d, f"{int(number):02d}", theme,
                        size=20, bold=True, color=theme.primary, align="center",
                        fill=theme.white, kind=MSO_SHAPE.OVAL, name="HeaderNumber")
        disc.text_frame.margin_left = disc.text_frame.margin_right = 0  # 清零圆圈内左右边距防止数字换行
        x += d + Inches(0.18)  # 标题文字的起始 X 坐标 = 圆圈右边 + 0.18 英寸间距
    # 3. 画标题栏上的白色大标题文字（默认字号 28 磅加粗，后期嫌标题字大可改成 26）
    text_box(slide, x, 0, SLIDE_W - x - Inches(0.3), HEADER_H, title, theme, size=28,
             bold=True, color=theme.white, name="HeaderTitle")


def cover(slide, theme: Theme, meta: dict, spec: dict):
    # 【函数作用】画第 1 页“封面页”：上方作者与期刊年份、中间论文大标题、胶囊标签、下半部深色色带与文献 DOI
    band_top = Inches(4.35)  # 下半部深紫色大色带的起始 Y 坐标（距离顶部 4.35 英寸）
    # 1. 画下半部分的深紫色大色带（从 Y=4.35 英寸一直铺到底部 7.5 英寸）
    shape_box(slide, 0, band_top, SLIDE_W, SLIDE_H - band_top, fill=theme.primary, name="CoverBand")
    # 拼接顶部的署名行（例如 "Wang 等 | Nature | 2026"）
    byline = spec.get("byline") or " | ".join(
        str(v) for v in (meta.get("authors"), meta.get("venue"), meta.get("year")) if v)
    if byline:
        y = Inches(0.62)  # 顶部署名装饰线的 Y 坐标
        # 2. 在署名文字左右两侧各画一条细横线和一个小菱形装饰点
        for side in (-1, 1):
            cx = SLIDE_W / 2 + side * Inches(3.6)
            connector(slide, cx - Inches(1.3), y, cx + Inches(1.3), y, theme.line, 1.0, arrow=False)
            shape_box(slide, cx + side * Inches(1.3) - Inches(0.05), y - Inches(0.05), Inches(0.1),
                      Inches(0.1), fill=theme.primary, kind=MSO_SHAPE.DIAMOND)
        # 3. 写顶部的作者与期刊年份文字（字号 20 磅加粗居中）
        text_box(slide, Inches(2.5), Inches(0.35), SLIDE_W - Inches(5), Inches(0.55), byline,
                 theme, size=20, bold=True, color=theme.primary, align="center")
    # 4. 写封面正中央的论文主标题（默认超大字号 50 磅加粗居中，如果标题字数特别多可在 spec 里把 title_size 调小到 40）
    title = spec.get("title") or meta.get("title", "")
    text_box(slide, Inches(0.6), Inches(1.05), SLIDE_W - Inches(1.2), Inches(2.75), title, theme,
             size=spec.get("title_size", 50), bold=True, color=theme.primary, align="center",
             line_spacing=1.05)
    # 5. 画卡在白底与紫底交界线上的圆角胶囊标签（例如 "文献汇报 | 15分钟"）
    tag = spec.get("tag") or meta.get("tag")
    if tag:
        w = Inches(3.4)  # 胶囊标签宽度 3.4 英寸
        text_box(slide, (SLIDE_W - w) / 2, band_top - Inches(0.3), w, Inches(0.6), tag, theme,
                 size=20, bold=True, color=theme.white, align="center", fill=theme.primary,
                 line=theme.white, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    # 6. 在下半部紫色色带里写文献英文原名、期刊与 DOI 信息（白色 16 磅字）
    info = spec.get("info") or meta.get("cover_info")
    if info:
        text_box(slide, Inches(1.6), band_top + Inches(0.75), SLIDE_W - Inches(3.2), Inches(1.9),
                 info, theme, size=16, color=theme.white, anchor="top", line_spacing=1.3)


def agenda(slide, theme: Theme, meta: dict, spec: dict):
    # 【函数作用】画第 2 页“汇报目录页”：底部紫色衬底 + 中央带阴影白面板 + 左侧竖排“汇报目录” + 右侧 01~06 章节条
    # 1. 画下半部紫色背景衬底
    shape_box(slide, 0, Inches(4.35), SLIDE_W, SLIDE_H - Inches(4.35), fill=theme.primary)
    # 2. 画中央悬浮的白色大卡片面板（开启 shadow=True 带立体投影效果）
    panel = shape_box(slide, Inches(1.4), Inches(0.35), Inches(10.5), Inches(6.55),
                      fill=theme.white, name="AgendaPanel", shadow=True)
    panel.shadow.inherit = True
    # 3. 写面板左侧竖排的“汇\n报\n目\n录”四个宋体大字（字号 54 磅）
    title = spec.get("title", "汇\n报\n目\n录")
    text_box(slide, Inches(3.1), Inches(0.8), Inches(1.3), Inches(5.6), title, theme, size=54,
             bold=True, color=theme.primary, align="center", font="SimSun", line_spacing=1.0)
    # 4. 画竖排大字左边的竖向分割细线和旋转 90 度的灰色 "CONTENTS" 装饰字
    connector(slide, Inches(3.05), Inches(2.1), Inches(3.05), Inches(5.6), theme.line, 1.2, False)
    sub = text_box(slide, Inches(1.9), Inches(2.6), Inches(1.2), Inches(2.6), "CONTENTS", theme,
                   size=22, bold=True, color="D0D0D0", align="center")
    sub.rotation = 90
    # 5. 循环绘制右侧的目录条目（例如 01 研究背景 ~ 06 结论与展望）
    items = spec.get("items", [])
    active = spec.get("active", 0)  # 当前高亮第几个章节（默认 0 即高亮第 1 章 01）
    top, step = Inches(0.62), min(Inches(0.96), Inches(5.9) / max(1, len(items)))
    for i, item in enumerate(items):
        on = i == active  # 判断当前行是不是高亮章节
        y = top + i * step
        # 画每个章节的胶囊外框（高亮项带主色边框与水平渐变底色，非高亮项为浅灰边框）
        row = shape_box(slide, Inches(5.05), y, Inches(5.9), step - Inches(0.2),
                        fill=theme.softer if on else theme.white,
                        line=theme.primary if on else "E6E6E6", kind=MSO_SHAPE.ROUNDED_RECTANGLE,
                        radius=0.5, line_w=1.0)
        if on:
            gradient_fill(row, [theme.soft, theme.white], 0)
        c = theme.primary if on else "C8C8C8"
        # 写胶囊左边的章节序号（01、02...）
        text_box(slide, Inches(5.25), y, Inches(1.0), step - Inches(0.2), f"{i + 1:02d}", theme,
                 size=26, bold=True, color=c, align="center")
        # 画序号与章节名之间的短竖线
        connector(slide, Inches(6.35), y + Inches(0.15), Inches(6.35), y + step - Inches(0.35),
                  c, 1.0, False)
        # 写胶囊右边的章节名称
        text_box(slide, Inches(6.5), y, Inches(4.3), step - Inches(0.2), item, theme, size=24,
                 color=theme.primary if on else "BDBDBD")


def section(slide, theme: Theme, meta: dict, spec: dict):
    # 【函数作用】画“章节过渡页”：左边占 4.6 英寸宽的深色块配超大章节数字，右边写章节标题与副标题
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
    # 【函数作用】画第 16 页“感谢聆听 / 欢迎交流讨论”封底致谢页
    band_top = Inches(4.75)  # 封底页下半部紫色色带的起始 Y 坐标
    shape_box(slide, 0, band_top, SLIDE_W, SLIDE_H - band_top, fill=theme.primary)
    # 顶部短标题与左右装饰细线
    top_line = spec.get("byline") or meta.get("short_title") or ""
    if top_line:
        y = Inches(0.95)
        connector(slide, Inches(2.0), y, Inches(4.1), y, theme.line, 1.0, False)
        connector(slide, Inches(9.2), y, Inches(11.3), y, theme.line, 1.0, False)
        text_box(slide, Inches(4.1), Inches(0.65), Inches(5.1), Inches(0.6), top_line, theme,
                 size=22, bold=True, color=theme.primary, align="center")
    # 正中央超大号“感谢聆听\n欢迎交流讨论”（默认 60 磅粗体）
    text_box(slide, Inches(1), Inches(1.45), SLIDE_W - Inches(2), Inches(3.0),
             spec.get("title", "感谢聆听\n欢迎交流讨论"), theme, size=spec.get("title_size", 60),
             bold=True, color=theme.primary, align="center", line_spacing=1.1)
    # 交界线处的“文献汇报”圆角胶囊标签
    tag = spec.get("tag") or meta.get("tag_short") or "文献汇报"
    w = Inches(2.8)
    text_box(slide, (SLIDE_W - w) / 2, band_top - Inches(0.3), w, Inches(0.6), tag, theme,
             size=22, bold=True, color=theme.white, align="center", fill=theme.primary,
             line=theme.white, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
