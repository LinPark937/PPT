"""Slides 1 to 6 of the Nature 2026 presentation."""
from pathlib import Path
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from .blocks import banner, bullets
from .builder import connector, gradient_fill, picture_fit, shape_box, text_box
from .layouts_basic import agenda, cover, header_bar
from .nature2026_data import AGENDA_ITEMS, META
from .theme import SLIDE_W, Theme


def build_slide_01(slide, theme: Theme, fig_dir: Path):
    cover(slide, theme, META, {})


def build_slide_02(slide, theme: Theme, fig_dir: Path):
    agenda(slide, theme, META, {"items": AGENDA_ITEMS, "active": 0})


def build_slide_03(slide, theme: Theme, fig_dir: Path):
    header_bar(slide, theme, "从双轴走向网状三轴分子编织", 3)
    bullets(slide, theme, [
        "已有分子编织：主要依赖**线性构件**和**双轴结构**",
        "高阶拓扑设计：**网状三轴**的精确构筑及**性能研究**仍少",
    ], Inches(0.35), Inches(1.02), Inches(12.5), style="diamond", size=21, row_h=Inches(0.45))

    y_card, h_card = Inches(2.0), Inches(5.15)
    w_card = Inches(6.15)
    # Left card
    shape_box(slide, Inches(0.32), y_card, w_card, h_card, fill=theme.white, line=theme.line)
    text_box(slide, Inches(0.32), y_card, w_card, Inches(0.65), "已有结构：双轴编织",
             theme, size=24, bold=True, align="center", fill=theme.soft)
    text_box(slide, Inches(0.5), Inches(2.75), w_card - Inches(0.36), Inches(0.85),
             "线性与网状构件\n已形成**双轴编织结构**", theme, size=20, bold=True, align="center")
    picture_fit(slide, fig_dir / "s03_linear_biaxial.png", Inches(0.65), Inches(3.68), Inches(2.4), Inches(2.5))
    picture_fit(slide, fig_dir / "s03_reticular_biaxial.png", Inches(3.45), Inches(3.68), Inches(2.5), Inches(2.5))
    text_box(slide, Inches(0.55), Inches(6.32), Inches(2.75), Inches(0.62), "线性双轴结构",
             theme, size=18, bold=True, align="center", fill="EAE8EE")
    text_box(slide, Inches(3.45), Inches(6.32), Inches(2.75), Inches(0.62), "网状双轴结构",
             theme, size=18, bold=True, align="center", fill="EAE8EE")

    # Right card
    rx = Inches(6.82)
    shape_box(slide, rx, y_card, w_card, h_card, fill=theme.white, line=theme.line)
    text_box(slide, rx, y_card, w_card, Inches(0.65), "研究缺口：网状三轴",
             theme, size=24, bold=True, align="center", fill=theme.soft)
    text_box(slide, rx + Inches(0.2), Inches(2.75), w_card - Inches(0.4), Inches(0.85),
             "需要精确构筑**网状三轴**拓扑\n并检验其材料**性能**", theme, size=20, bold=True, align="center")
    picture_fit(slide, fig_dir / "s03_linear_triaxial.png", rx + Inches(0.2), Inches(3.7), Inches(1.8), Inches(1.9))
    picture_fit(slide, fig_dir / "s03_reticular_triaxial.png", rx + Inches(2.15), Inches(3.7), Inches(1.8), Inches(1.9))
    picture_fit(slide, fig_dir / "s03_c3_monomer.png", rx + Inches(4.2), Inches(3.8), Inches(1.6), Inches(1.5))
    text_box(slide, rx + Inches(0.2), Inches(5.6), Inches(1.8), Inches(0.5), "线性三轴",
             theme, size=18, bold=True, align="center")
    text_box(slide, rx + Inches(2.15), Inches(5.6), Inches(1.8), Inches(0.5), "网状三轴",
             theme, size=18, bold=True, align="center")
    text_box(slide, rx + Inches(4.05), Inches(5.45), Inches(1.95), Inches(0.75), "C_{3}构件\n对称性示意",
             theme, size=16, bold=True, align="center")
    text_box(slide, rx + Inches(0.25), Inches(6.32), w_card - Inches(0.5), Inches(0.62),
             "网状三轴的构筑与性能研究仍待拓展", theme, size=18, bold=True, align="center", fill="EAE8EE")


def build_slide_04(slide, theme: Theme, fig_dir: Path):
    header_bar(slide, theme, "C3大环与DB2的配位组装", 4)
    bullets(slide, theme, [
        "引入C_{3}对称大环与二连接DB2，通过**配位B–N聚合**构筑网络",
        "结合刚性结构支撑与柔性构象调节，为网状三轴交织提供条件",
    ], Inches(0.35), Inches(0.98), Inches(12.5), style="dot", size=20, row_h=Inches(0.42))

    # Left panel
    lx, ly, lw, lh = Inches(0.22), Inches(1.92), Inches(5.05), Inches(4.62)
    shape_box(slide, lx, ly, lw, lh, fill=theme.white, line=theme.line)
    text_box(slide, lx, ly, Inches(1.65), Inches(0.54), "构件选择", theme, size=19, bold=True,
             color=theme.white, align="center", fill=theme.primary)
    text_box(slide, lx + Inches(1.65), ly, lw - Inches(1.65), Inches(0.54), "C3大环与二连接分子",
             theme, size=19, bold=True, align="center", fill=theme.soft)
    text_box(slide, lx + Inches(0.15), Inches(2.5), lw - Inches(0.3), Inches(0.45),
             "**可逆配位B–N键驱动连接**", theme, size=20, bold=True, align="center")
    text_box(slide, lx + Inches(0.18), Inches(2.98), lw - Inches(0.36), Inches(0.76),
             "大环刚性片段引导有序延展\n亚甲基片段提供构象适应性", theme, size=16, bold=True,
             align="center", fill="EEECEF")
    picture_fit(slide, fig_dir / "s04_mcn3_db2.png", lx + Inches(0.15), Inches(3.8), lw - Inches(0.3), Inches(1.72))
    for i, lbl in enumerate(("C_{3}对称\n三连接", "吡啶位点\n氮供体", "DB2单体\n二连接")):
        text_box(slide, lx + Inches(0.14) + i * Inches(1.62), Inches(5.6), Inches(1.52), Inches(0.8),
                 lbl, theme, size=15, bold=True, align="center", fill=theme.softer)

    # Middle flow column
    mx, mw = Inches(5.42), Inches(1.55)
    steps = ["构件几何\n设计要素", "配位连接\nB–N聚合", "构象调节\n刚柔协同", "三轴交织\n设计目标"]
    for i, st in enumerate(steps):
        sy = Inches(1.92) + i * Inches(1.26)
        text_box(slide, mx, sy, mw, Inches(0.82), st, theme, size=16, bold=True,
                 align="center", fill=theme.white, line=theme.line)
        if i < 3:
            shape_box(slide, mx + Inches(0.58), sy + Inches(0.86), Inches(0.38), Inches(0.34),
                      fill=theme.soft, kind=MSO_SHAPE.DOWN_ARROW)

    # Right panel
    rx, rw = Inches(7.12), Inches(5.95)
    text_box(slide, rx, Inches(1.92), rw, Inches(0.54), "三轴交织的局部结构依据",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary)
    picture_fit(slide, fig_dir / "s04_coord_3d.png", rx + Inches(0.1), Inches(2.55), Inches(4.0), Inches(1.55))
    # Native N -> B badge
    text_box(slide, rx + Inches(4.3), Inches(3.05), Inches(0.52), Inches(0.52), "N", theme,
             size=16, color=theme.primary, align="center", fill=theme.softer, line=theme.primary, kind=MSO_SHAPE.OVAL)
    connector(slide, rx + Inches(4.85), Inches(3.31), rx + Inches(5.2), Inches(3.31), theme.primary, 1.5, True)
    text_box(slide, rx + Inches(5.22), Inches(3.05), Inches(0.52), Inches(0.52), "B", theme,
             size=16, color=theme.accent, align="center", fill="F9EBEB", line=theme.accent, kind=MSO_SHAPE.OVAL)

    bar = shape_box(slide, rx + Inches(0.1), Inches(4.12), rw - Inches(0.2), Inches(0.22),
                    fill=theme.primary, kind=MSO_SHAPE.RIGHT_ARROW)
    gradient_fill(bar, [theme.soft, theme.primary], 0)
    cols = [("三臂配位\n**大环供体**"), ("之字链\n**上下跨越**"), ("可逆连接\n**配位B–N**"), ("构象适应\n**网络交织**")]
    for i, txt in enumerate(cols):
        text_box(slide, rx + Inches(0.1) + i * Inches(1.46), Inches(4.36), Inches(1.4), Inches(0.68),
                 txt, theme, size=14, bold=True, align="center")
    for i, (head, body) in enumerate([("刚性支撑", "刚性片段引导链段\n定向、有序地延展"),
                                      ("柔性调节", "亚甲基与动态配位\n提供构象适应性")]):
        bx = rx + Inches(0.1) + i * Inches(2.95)
        text_box(slide, bx, Inches(5.12), Inches(2.8), Inches(0.45), head, theme,
                 size=17, bold=True, align="center", fill=theme.soft)
        text_box(slide, bx, Inches(5.6), Inches(2.8), Inches(0.92), body, theme,
                 size=15, bold=True, align="center", fill="F5F3F7")

    banner(slide, theme, "C_{3}大环与可逆B–N连接协同，为网状三轴编织提供构筑条件",
           Inches(0.22), Inches(6.68), SLIDE_W - Inches(0.44), Inches(0.58), style="bar", size=20)


def build_slide_05(slide, theme: Theme, fig_dir: Path):
    header_bar(slide, theme, "80℃、24小时获得TWPN单晶", 5)
    text_box(slide, Inches(0.45), Inches(0.95), Inches(12.2), Inches(1.0),
             "MCN3与DB2在甲苯中于80℃加热24小时，\n缓慢冷却至室温后获得淡黄色TWPN单晶",
             theme, size=24, bold=True)
    text_box(slide, Inches(1.5), Inches(1.96), Inches(10.3), Inches(0.55),
             "**单晶解析：C2/c，B–N键长1.65(1) Å**", theme, size=26, bold=True, align="center")

    cards = [
        ("合成示意", "s05_synthesis.png", Inches(0.22), Inches(3.05)),
        ("单晶实物", "s05_crystal_photo.png", Inches(3.58), Inches(3.05)),
        ("网络结构", "s05_hex_net.png", Inches(6.94), Inches(3.05)),
        ("键长示意", None, Inches(10.30), Inches(2.75)),
    ]
    for title, img, x, w in cards:
        text_box(slide, x, Inches(2.65), w, Inches(0.58), title, theme, size=22, bold=True,
                 color=theme.white, align="center", fill=theme.primary,
                 kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.18)
        shape_box(slide, x, Inches(3.28), w, Inches(2.12), fill=theme.white, line=theme.line)
        if img:
            picture_fit(slide, fig_dir / img, x + Inches(0.1), Inches(3.34), w - Inches(0.2), Inches(1.98))
        else:
            # Native N - B bond length diagram
            text_box(slide, x + Inches(0.32), Inches(3.62), Inches(0.62), Inches(0.62), "N", theme,
                     size=20, color=theme.primary, align="center", fill=theme.softer,
                     line=theme.primary, kind=MSO_SHAPE.OVAL)
            connector(slide, x + Inches(0.95), Inches(3.93), x + Inches(1.8), Inches(3.93), "333333", 1.8, False)
            text_box(slide, x + Inches(1.82), Inches(3.62), Inches(0.62), Inches(0.62), "B", theme,
                     size=20, color=theme.accent, align="center", fill="F9EBEB",
                     line=theme.accent, kind=MSO_SHAPE.OVAL)
            connector(slide, x + Inches(0.62), Inches(4.42), x + Inches(2.12), Inches(4.42), "555555", 1.2, False)
            text_box(slide, x + Inches(0.5), Inches(4.5), w - Inches(1.0), Inches(0.45),
                     "1.65(1) Å", theme, size=16, align="center")

    connector(slide, Inches(0.22), Inches(5.65), SLIDE_W - Inches(0.22), Inches(5.65), theme.line, 1.5, False)
    text_box(slide, Inches(0.28), Inches(5.9), Inches(3.65), Inches(1.05),
             "R1: Wang等, Nature 2026\nFig.1b与合成、单晶解析正文", theme, size=16, bold=True,
             fill=theme.white, line=theme.line, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
    for i, txt in enumerate(["实验观察\n淡黄色单晶", "结构解析\nB–N: 1.65(1) Å", "理论计算\n122.95 kJ/mol"]):
        text_box(slide, Inches(4.2) + i * Inches(3.0), Inches(5.9), Inches(2.85), Inches(1.05),
                 txt, theme, size=20, bold=True, align="center")


def build_slide_06(slide, theme: Theme, fig_dir: Path):
    header_bar(slide, theme, "四网络交织构成层状分子织物", 6)
    cols = [
        ("层级一", "四套hcb网络交织", "s06_layer1_hcb.png"),
        ("层级二", "每组三网络构成节点", "s06_layer2_nodes.png"),
        ("层级三", "单层与AB层间堆积", "s06_layer3_stack.png"),
    ]
    for i, (tag, subtitle, img) in enumerate(cols):
        x = Inches(0.32) + i * Inches(4.28)
        w = Inches(4.08)
        shape_box(slide, x, Inches(1.38), w, Inches(4.85), fill=theme.white, line="B8B8B8")
        text_box(slide, x + Inches(0.55), Inches(1.02), w - Inches(1.1), Inches(0.66), tag,
                 theme, size=24, bold=True, color=theme.white, align="center",
                 fill=theme.primary, kind=MSO_SHAPE.PARALLELOGRAM)
        text_box(slide, x + Inches(0.15), Inches(1.82), w - Inches(0.3), Inches(0.65), subtitle,
                 theme, size=24, bold=True, align="center")
        shape_box(slide, x + Inches(0.22), Inches(2.62), w - Inches(0.44), Inches(3.4),
                  fill=theme.white, line=theme.line)
        picture_fit(slide, fig_dir / img, x + Inches(0.28), Inches(2.68), w - Inches(0.56), Inches(3.28))

    bar = shape_box(slide, Inches(0.32), Inches(6.35), SLIDE_W - Inches(0.64), Inches(0.82), fill=theme.soft)
    gradient_fill(bar, [theme.soft, theme.primary, theme.soft], 0)
    text_box(slide, Inches(0.5), Inches(6.38), SLIDE_W - Inches(1.0), Inches(0.76),
             "**四网络交织；单层约19.5 Å，层间AB堆积 (R1, Fig.1c–e)**",
             theme, size=25, bold=True, align="center")
