# ==============================================================================
# 【文件作用】PPT 的“皮肤配色与页面尺寸配置中心” (cellppt/theme.py)
# 后期如果你想改 PPT 的整体主色调、强调色、默认中英文字体、或者页边距，都在这个文件里微调。
# ==============================================================================

from dataclasses import dataclass, field  # 导入数据类工具，用来方便地定义主题配置结构
from pptx.dml.color import RGBColor       # 导入 python-pptx 的 RGB 颜色类
from pptx.util import Emu, Inches         # 导入长度单位：Inches(英寸) 和 Emu(PPT底层长度单位)


def hex_rgb(value: str) -> RGBColor:
    # 【函数作用】把 6 位十六进制颜色字符串（如 "6A1069" 或 "#6A1069"）转换成 PPT 认识的 RGBColor 对象
    value = value.lstrip("#")  # 去掉开头可能带的 '#' 号
    # 分别截取前2位(红R)、中2位(绿G)、后2位(蓝B)，从16进制转为整数并生成 RGBColor
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


@dataclass
class Theme:
    # 【类作用】定义一套 PPT 主题皮肤包含的所有颜色和字体属性（后期微调字体或默认字色改这里）
    name: str                                  # 主题名称（如 "nature_purple"）
    primary: str                               # 主色调：用于顶部标题栏、主卡片表头、封面底栏（6位十六进制颜色）
    primary_dark: str                          # 深主色：用于深色边框或次级强调
    accent: str                                # 高亮强调色：正文中用 **包裹的文字** 会自动变成这个颜色（如深红 C00000）
    soft: str                                  # 浅主色：用于浅色卡片的表头、胶囊标签底色
    softer: str                                # 极浅主色：用于卡片正文的浅色背景填充
    text: str = "1F1F1F"                       # 默认正文文字颜色（深灰接近黑，比纯黑更柔和，可微调）
    muted: str = "6B6B6B"                      # 次要/注释文字颜色（中灰色，可微调）
    line: str = "D9D2E0"                       # 卡片外边框线、分割线的默认颜色（浅灰紫，可微调粗细和深浅）
    white: str = "FFFFFF"                      # 纯白色常量
    font_cn: str = "Microsoft YaHei"           # 默认中文字体（微软雅黑，后期想改黑体/思源黑体改这里）
    font_en: str = "Arial"                     # 默认英文字体与数字字体（Arial，后期想改 Times New Roman 改这里）
    font_title_cn: str = "Microsoft YaHei"     # 标题中文字体
    gradient: list = field(default_factory=list)  # 渐变色停靠点列表（左->中->右的颜色过渡）

    def rgb(self, key_or_hex: str) -> RGBColor:
        # 【函数作用】既可以传属性名（如 theme.rgb("primary")），也可以直接传颜色码（如 theme.rgb("FF0000")）获取颜色对象
        return hex_rgb(getattr(self, key_or_hex, key_or_hex))


# ==============================================================================
# 【预设的 4 套顶刊学术配色库】（后期想新增自己的配色，照着格式加一行即可）
# ==============================================================================
THEMES = {
    # 1. Nature 经典深紫配色（PDF 视频同款：深紫标题栏 + 深红高亮词 + 浅紫卡片）
    "nature_purple": Theme(
        "nature_purple",
        primary="6A1069",       # 微调主紫色
        primary_dark="4E0B4D",  # 微调深紫色
        accent="C00000",        # 微调 **高亮文字** 的深红色
        soft="EADCF0",          # 微调卡片浅紫表头
        softer="F6F0F8",        # 微调卡片极浅紫底色
        gradient=["F3E8F5", "6A1069", "F3E8F5"],  # 底部渐变条的左-中-右颜色
    ),
    # 2. Cell 经典深蓝配色（深海蓝标题栏 + 橙红高亮词 + 冰蓝卡片）
    "cell_blue": Theme(
        "cell_blue",
        primary="0B4F8A",       # 微调主蓝色
        primary_dark="073763",  # 微调深蓝色
        accent="D9480F",        # 微调高亮橙红色
        soft="DCEAF7",          # 微调浅蓝卡片表头
        softer="F1F7FC",        # 微调极浅蓝底色
        gradient=["E6F0FA", "0B4F8A", "E6F0FA"],
    ),
    # 3. Science 经典深红配色（勃艮第红标题栏 + 宝石蓝高亮词）
    "science_red": Theme(
        "science_red",
        primary="9E1B1B",
        primary_dark="701313",
        accent="1D5FA8",
        soft="F6DEDE",
        softer="FCF3F3",
        gradient=["FAEAEA", "9E1B1B", "FAEAEA"],
    ),
    # 4. 森林绿学术配色（墨绿标题栏 + 琥珀橙高亮词）
    "forest_green": Theme(
        "forest_green",
        primary="1E6B4A",
        primary_dark="134A33",
        accent="C2410C",
        soft="DDF0E6",
        softer="F2FAF5",
        gradient=["E5F4EC", "1E6B4A", "E5F4EC"],
    ),
}


def get_theme(name: str | None) -> Theme:
    # 【函数作用】根据主题名字取出配色对象；如果没传名字或名字不存在，默认返回 "nature_purple"
    return THEMES.get(name or "nature_purple", THEMES["nature_purple"])


# ==============================================================================
# 【全局幻灯片几何尺寸常量】（16:9 宽屏标准尺寸，后期想微调标题栏高度或页边距改这里）
# ==============================================================================
SLIDE_W = Inches(13.333)      # 幻灯片总宽度：13.333 英寸（标准 16:9 宽屏宽度）
SLIDE_H = Inches(7.5)         # 幻灯片总高度：7.5 英寸（标准 16:9 宽屏高度）
MARGIN = Inches(0.35)         # 页面左右默认留白边距：0.35 英寸
HEADER_H = Inches(0.92)       # 正文页顶部紫色标题栏的高度：0.92 英寸（想让标题栏变窄可改成 0.85）
CONTENT_TOP = Inches(1.08)    # 正文内容区的起始 Y 坐标：1.08 英寸


def emu(v) -> Emu:
    # 【函数作用】把任意数值转换为 python-pptx 底层的 Emu 整数单位
    return Emu(int(v))
