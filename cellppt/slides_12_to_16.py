# ==============================================================================
# 【文件作用】绘制 Nature 2026 示例 PPT 的第 12 页到第 16 页 (cellppt/slides_12_to_16.py)
# 后期如果你想微调第 12~16 页里的文字、圆环图百分比、柱状图说明或致谢语，直接改对应函数即可。
# ==============================================================================

from pathlib import Path                                           # 导入路径对象
from pptx.enum.shapes import MSO_SHAPE                             # 导入 PPT 内置形状枚举
from pptx.util import Inches                                       # 导入英寸单位

from .builder import connector, gradient_fill, picture_fit, shape_box, text_box  # 导入底层画笔
from .layouts_basic import header_bar, thanks                      # 导入顶部标题栏与封底致谢页模板
from .nature2026_data import META                                  # 导入全局元数据
from .theme import SLIDE_W, Theme                                  # 导入幻灯片总宽与主题类


def build_slide_12(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 12 页：《12 内源孔与编织协同孔构成多孔框架》（左右分栏互证：左栏孔道尺寸，右栏 CO2 吸附表征）
    header_bar(slide, theme, "内源孔与编织协同孔构成多孔框架", 12)
    # --- 1. 左半区：孔道与空隙连接 ---
    text_box(slide, Inches(1.4), Inches(1.0), Inches(3.6), Inches(0.58), "孔道与空隙连接",
             theme, size=22, bold=True, color=theme.white, align="center",
             fill=theme.primary, kind=MSO_SHAPE.PARALLELOGRAM)
    # 左侧竖排 4 个圆角胶囊指标
    pills = ["大环内源孔9.1 Å", "编织协同孔12.7 Å", "三轴节点支撑空隙", "通道连接层内空隙"]
    for i, p_txt in enumerate(pills):
        text_box(slide, Inches(0.28), Inches(1.78) + i * Inches(0.78), Inches(3.05), Inches(0.58),
                 p_txt, theme, size=19, bold=True, align="center", fill="E7DFEC",
                 kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    # 指向中间的渐变大右箭头
    arr = shape_box(slide, Inches(3.55), Inches(2.35), Inches(2.5), Inches(1.75),
                    fill=theme.primary, kind=MSO_SHAPE.RIGHT_ARROW)
    gradient_fill(arr, ["F6F0F8", theme.primary], 0)
    text_box(slide, Inches(3.6), Inches(2.75), Inches(2.2), Inches(0.95), "孔道共同连接\n层内大空隙",
             theme, size=19, bold=True, color=theme.white, align="center")

    # 左下角沿 b 轴与沿 a 轴侧视图
    text_box(slide, Inches(0.45), Inches(5.15), Inches(2.0), Inches(0.45), "沿b轴侧视",
             theme, size=19, bold=True, align="center")
    picture_fit(slide, fig_dir / "s12_side_b.png", Inches(0.38), Inches(5.65), Inches(2.1), Inches(1.42))
    text_box(slide, Inches(3.25), Inches(5.15), Inches(2.4), Inches(0.45), "沿a轴侧视",
             theme, size=19, bold=True, align="center")
    picture_fit(slide, fig_dir / "s12_side_a.png", Inches(2.85), Inches(5.65), Inches(3.15), Inches(1.42))

    # 中央竖向分割细线
    connector(slide, Inches(6.25), Inches(1.1), Inches(6.25), Inches(7.1), theme.line, 1.0, False)

    # --- 2. 右半区：吸附与孔隙证据 ---
    text_box(slide, Inches(7.9), Inches(1.0), Inches(3.6), Inches(0.58), "吸附与孔隙证据",
             theme, size=22, bold=True, color=theme.white, align="center",
             fill=theme.primary, kind=MSO_SHAPE.PARALLELOGRAM)
    text_box(slide, Inches(6.6), Inches(1.72), Inches(6.3), Inches(0.52),
             "**CO_{2}吸附确认可接近表面积**", theme, size=24, bold=True, align="center")
    shape_box(slide, Inches(6.48), Inches(2.35), Inches(6.4), Inches(1.92), fill="F3F2F5")
    text_box(slide, Inches(6.58), Inches(2.65), Inches(0.85), Inches(1.2), "吸附\n表征",
             theme, size=20, bold=True, align="center")
    text_box(slide, Inches(7.5), Inches(2.48), Inches(3.35), Inches(1.65),
             "◆  CO_{2}吸附温度：195 K\n◆  BET面积：91.6 m^{2}/g\n◆  结构来源：R1, ED Fig.2",
             theme, size=18, bold=True)
    # 用原生矢量圆圈画出 O=C=O 分子与气体进入孔隙示意图
    shape_box(slide, Inches(11.0), Inches(2.5), Inches(1.7), Inches(0.75), fill=theme.white)
    for idx, (lbl, col, dx) in enumerate([("O", theme.accent, 0.25), ("C", "666666", 0.72), ("O", theme.accent, 1.18)]):
        text_box(slide, Inches(11.0 + dx), Inches(2.6), Inches(0.28), Inches(0.28), lbl, theme,
                 size=9, bold=True, color=theme.white, align="center", fill=col, kind=MSO_SHAPE.OVAL)
    text_box(slide, Inches(11.0), Inches(2.92), Inches(1.7), Inches(0.28), "CO2分子 | 示意", theme, size=9, align="center")
    shape_box(slide, Inches(11.0), Inches(3.35), Inches(1.7), Inches(0.75), fill=theme.white)
    shape_box(slide, Inches(11.75), Inches(3.48), Inches(0.75), Inches(0.35), fill=theme.soft, line=theme.primary)
    text_box(slide, Inches(11.0), Inches(3.82), Inches(1.7), Inches(0.26), "气体可达孔隙 | 示意", theme, size=9, align="center")

    # 右下角：结构与吸附互证及孔道截面图
    shape_box(slide, Inches(7.35), Inches(4.42), Inches(0.65), Inches(0.55), fill=theme.primary, kind=MSO_SHAPE.DOWN_ARROW)
    text_box(slide, Inches(6.55), Inches(5.12), Inches(2.35), Inches(0.45), "结构与吸附互证",
             theme, size=20, bold=True, align="center")
    text_box(slide, Inches(6.55), Inches(5.85), Inches(2.35), Inches(0.95), "晶体给出通道\n吸附确认可达",
             theme, size=20, bold=True, align="center")
    shape_box(slide, Inches(9.1), Inches(5.95), Inches(0.45), Inches(0.55), fill=theme.soft, kind=MSO_SHAPE.RIGHT_ARROW)
    text_box(slide, Inches(9.75), Inches(5.12), Inches(3.1), Inches(0.45), "孔道连接层内空隙",
             theme, size=20, bold=True, align="center")
    picture_fit(slide, fig_dir / "s12_pore_insets.png", Inches(9.65), Inches(5.65), Inches(3.25), Inches(1.45))


def build_slide_13(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 13 页：《13 苯优先吸附与动态分离》（左下等温吸附/穿透曲线 + 右上原生苯环六边形 + 右下原生圆环图）
    header_bar(slide, theme, "苯优先吸附与动态分离", 13)
    text_box(slide, Inches(0.32), Inches(1.05), Inches(7.4), Inches(0.55),
             "苯/环己烷性质相近，TWPN仍表现选择性吸附", theme, size=23, bold=True)
    shape_box(slide, Inches(3.3), Inches(1.75), Inches(0.55), Inches(0.48), fill=theme.primary, kind=MSO_SHAPE.DOWN_ARROW)
    text_box(slide, Inches(0.32), Inches(2.35), Inches(7.3), Inches(0.55),
             "用**等温吸附**、**竞争吸附**与**穿透曲线**交叉验证", theme, size=23, bold=True, align="center")

    # 1. 右上：芳香烃/脂肪烃分离示意 + PowerPoint 原生可编辑六边形绘制的 Cyc(环己烷) 与 Ben(苯)
    shape_box(slide, Inches(7.95), Inches(0.98), Inches(2.35), Inches(1.75), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s13_arom_aliph.png", Inches(8.02), Inches(1.02), Inches(2.2), Inches(1.65))
    shape_box(slide, Inches(10.45), Inches(0.98), Inches(2.4), Inches(1.75), fill=theme.white, line=theme.line)
    # 画紫色六边形 Cyc
    shape_box(slide, Inches(10.68), Inches(1.22), Inches(0.75), Inches(0.75), fill=theme.white,
              line=theme.primary, kind=MSO_SHAPE.HEXAGON, line_w=2.0)
    text_box(slide, Inches(10.58), Inches(2.05), Inches(0.95), Inches(0.45), "Cyc", theme, size=18, align="center")
    # 画蓝色六边形 + 内圈圆环表示苯环 Ben
    shape_box(slide, Inches(11.78), Inches(1.22), Inches(0.75), Inches(0.75), fill=theme.white,
              line="1D70B8", kind=MSO_SHAPE.HEXAGON, line_w=2.0)
    shape_box(slide, Inches(11.96), Inches(1.4), Inches(0.39), Inches(0.39), fill=theme.white,
              line="1D70B8", kind=MSO_SHAPE.OVAL, line_w=1.2)
    text_box(slide, Inches(11.68), Inches(2.05), Inches(0.95), Inches(0.45), "Ben", theme, size=18, align="center")
    text_box(slide, Inches(7.95), Inches(2.78), Inches(4.9), Inches(0.48), "**分子对象与孔隙识别**",
             theme, size=20, bold=True, align="center")
    connector(slide, Inches(7.9), Inches(3.28), Inches(12.88), Inches(3.28), theme.primary, 2.0, False)

    # 2. 左下：等温吸附曲线(f) 与 动态穿透曲线(g)
    shape_box(slide, Inches(0.22), Inches(3.45), Inches(7.25), Inches(3.65), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s13_isotherm_breakthrough.png", Inches(0.32), Inches(3.52), Inches(7.05), Inches(3.5))

    # 3. 右下：释放蒸气组成卡片 + 两个原生可编辑圆环形状（97.9% TWPN vs 60.5% MCN3）
    rx, rw = Inches(7.72), Inches(5.3)
    shape_box(slide, rx, Inches(3.48), rw, Inches(3.62), fill="F5F3F7")
    text_box(slide, rx + Inches(0.18), Inches(3.55), rw - Inches(0.36), Inches(0.55), "苯优先吸附与动态分离",
             theme, size=24, bold=True)
    text_box(slide, rx + Inches(0.18), Inches(4.22), rw - Inches(0.36), Inches(0.85),
             "等体积混合物竞争吸附后，释放蒸气Ben含量：\nTWPN为**97.9%**，MCN3对照为**60.5%**。",
             theme, size=17, bold=True)
    text_box(slide, rx + Inches(0.8), Inches(5.18), Inches(2.6), Inches(0.48), "释放蒸气组成",
             theme, size=22, bold=True, align="center")
    for idx, (pct, name, col) in enumerate([("97.9%\nBen", "TWPN", "2B7BBA"), ("60.5%\nBen", "MCN3", theme.primary)]):
        dx = rx + Inches(0.52) + idx * Inches(1.88)
        shape_box(slide, dx, Inches(5.78), Inches(1.15), Inches(1.15), fill=col, kind=MSO_SHAPE.DONUT)
        text_box(slide, dx + Inches(0.18), Inches(5.96), Inches(0.79), Inches(0.78), pct,
                 theme, size=12, bold=True, align="center")
        text_box(slide, dx, Inches(6.9), Inches(1.15), Inches(0.28), name, theme, size=11, bold=True, align="center")
    text_box(slide, rx + Inches(3.85), Inches(5.55), Inches(1.35), Inches(1.45),
             "蓝：Ben\n紫：Cyc\n等体积\n竞争吸附", theme, size=17, bold=True)


def build_slide_14(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 14 页：《14 多组分离、热再生与百克级制备》（左栏 6 组混合物柱状图，右栏 100g 级放大制备实物）
    header_bar(slide, theme, "多组分离、热再生与百克级制备", 14)
    # 1. 左上：分离范围与循环使用两条红头标签
    for i, (tag, txt) in enumerate([("分离范围", "六组芳香/脂肪烃配对均显示吸附偏好"),
                                    ("循环使用", "加热脱附后可再生，长期表现仍需验证")]):
        ry = Inches(1.05) + i * Inches(0.72)
        text_box(slide, Inches(0.22), ry, Inches(1.52), Inches(0.56), tag, theme,
                 size=19, bold=True, color=theme.white, align="center", fill="AC1B30")
        text_box(slide, Inches(1.98), ry, Inches(4.55), Inches(0.56), txt, theme,
                 size=17, bold=True, align="center", fill="F7F7F7", line=theme.line)

    # 2. 左下：释放蒸气组成柱状图卡片
    shape_box(slide, Inches(0.22), Inches(2.5), Inches(6.32), Inches(4.65), fill="F6F0F8")
    text_box(slide, Inches(0.38), Inches(2.56), Inches(6.0), Inches(0.52), "释放蒸气组成",
             theme, size=23, bold=True, align="center", fill="E7DCED")
    shape_box(slide, Inches(0.38), Inches(3.15), Inches(6.0), Inches(2.85), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s14_bar_chart.png", Inches(0.48), Inches(3.2), Inches(5.8), Inches(2.75))
    text_box(slide, Inches(0.45), Inches(6.08), Inches(5.9), Inches(0.95),
             "芳香组分占比92.1%–98.0%，n=3，误差为SEM。\n结果对应本研究竞争条件，不是选择性因子。",
             theme, size=16, bold=True)

    # 3. 右栏：百克级制备（100 g，产率 87.2%）与新增五组比较对象分子结构
    rx, rw = Inches(6.72), Inches(6.35)
    text_box(slide, rx, Inches(0.98), rw, Inches(0.56), "百克级制备与应用基础",
             theme, size=24, bold=True, align="center", fill="E7DCED")
    text_box(slide, rx, Inches(1.58), rw, Inches(0.88),
             "制备规模达到100 g，报告产率87.2%。\n实验室放大为后续材料应用提供基础。",
             theme, size=18, bold=True, align="center", fill="F5F3F7")
    shape_box(slide, rx + Inches(0.05), Inches(2.6), Inches(4.75), Inches(2.0), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s14_reactor_100g.png", rx + Inches(0.15), Inches(2.65), Inches(4.55), Inches(1.9))
    text_box(slide, rx + Inches(4.92), Inches(2.6), Inches(1.38), Inches(2.0),
             "**100 g**\n\n**87.2%**\n产率", theme, size=21, bold=True, align="center", fill="F0E6F5")

    shape_box(slide, rx + Inches(0.05), Inches(4.78), rw - Inches(0.1), Inches(1.5), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s14_five_pairs.png", rx + Inches(0.15), Inches(4.85), rw - Inches(0.3), Inches(1.35))
    text_box(slide, rx + Inches(0.05), Inches(6.42), rw - Inches(0.1), Inches(0.65),
             "新增五组比较对象 (Ben/Cyc见前页)", theme, size=19, bold=True, align="center", fill="ECE5F2")


def build_slide_15(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 15 页：《15 结论与展望：拓扑成为材料设计变量》（左列“已建立的认识”对比右列“待验证的问题”）
    header_bar(slide, theme, "结论与展望：拓扑成为材料设计变量", 15)
    # 1. 左右两列的梯形渐变表头（已建立的认识 vs 待验证的问题）
    for i, head in enumerate(["已建立的认识", "待验证的问题"]):
        hx = Inches(0.22) + i * Inches(6.65)
        bar = shape_box(slide, hx, Inches(0.95), Inches(6.22), Inches(0.55),
                        fill=theme.soft, kind=MSO_SHAPE.TRAPEZOID)
        gradient_fill(bar, ["FFFFFF", theme.soft], 90)
        text_box(slide, hx, Inches(0.95), Inches(6.22), Inches(0.55), head,
                 theme, size=24, bold=True, align="center")

    # 2. 上下两行左右对应的结论框与展望框
    rows = [
        ("C_{3}大环与DB2构建三轴网状编织，\n**单层同手性**、相邻层构型相反。",
         "构建同组分、不同拓扑的材料对照，\n**区分编织与化学组成的贡献。**"),
        ("弹性与分离功能在同一体系中呈现，\n结构—性能关系已有多类证据。",
         "拓展循环压缩、湿度与多组分条件，\n检验长期回复与动态分离表现。"),
    ]
    for r, (left_t, right_t) in enumerate(rows):
        ry = Inches(1.56) + r * Inches(1.38)
        text_box(slide, Inches(0.22), ry, Inches(6.22), Inches(1.22), left_t,
                 theme, size=20, bold=True, fill=theme.white, line=theme.line)
        shape_box(slide, Inches(6.52), ry + Inches(0.35), Inches(0.28), Inches(0.5),
                  fill="DDD6E5", kind=MSO_SHAPE.RIGHT_ARROW)
        text_box(slide, Inches(6.87), ry, Inches(6.22), Inches(1.22), right_t,
                 theme, size=20, bold=True, fill=theme.white, line=theme.line)

    # 3. 下半部分微观拓扑手性对图与三个拟验证实验示意图
    shape_box(slide, Inches(0.22), Inches(4.38), Inches(5.32), Inches(2.22), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s15_chirality_pair.png", Inches(0.45), Inches(4.45), Inches(4.85), Inches(2.08))

    shape_box(slide, Inches(5.68), Inches(4.38), Inches(7.41), Inches(2.22), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s15_proposed_tests.png", Inches(5.78), Inches(4.45), Inches(7.2), Inches(2.08))

    # 4. 底部梯形渐变总结条
    foot = shape_box(slide, Inches(0.22), Inches(6.68), SLIDE_W - Inches(0.44), Inches(0.62),
                     fill=theme.soft, kind=MSO_SHAPE.TRAPEZOID)
    gradient_fill(foot, ["FFFFFF", theme.soft], 90)
    text_box(slide, Inches(0.5), Inches(6.72), SLIDE_W - Inches(1.0), Inches(0.55),
             "**拓扑提供新的材料设计变量**；因果归因与实用工况仍需进一步检验。",
             theme, size=19, bold=True, align="center")


def build_slide_16(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 16 页：封底致谢页（调用 thanks 模板并在下方紫带中加上讨论问题）
    thanks(slide, theme, META, {})
    text_box(slide, Inches(1.6), Inches(5.45), SLIDE_W - Inches(3.2), Inches(1.65),
             META["thanks_info"], theme, size=16, color=theme.white, anchor="top", line_spacing=1.3)
