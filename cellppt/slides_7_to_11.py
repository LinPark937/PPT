# ==============================================================================
# 【文件作用】绘制 Nature 2026 示例 PPT 的第 7 页到第 11 页 (cellppt/slides_7_to_11.py)
# 后期如果你想微调第 7~11 页里的文字、卡片尺寸或配图位置，直接改对应函数即可。
# ==============================================================================

from pathlib import Path                                           # 导入路径对象
from pptx.enum.shapes import MSO_SHAPE                             # 导入 PPT 内置几何形状枚举
from pptx.util import Inches                                       # 导入英寸单位

from .blocks import banner, bullets                                # 导入要点列表与底部横幅组件
from .builder import connector, gradient_fill, picture_fit, shape_box, text_box  # 导入底层画笔
from .layouts_basic import header_bar                              # 导入顶部紫色标题栏函数
from .theme import SLIDE_W, Theme                                  # 导入幻灯片总宽与主题类


def build_slide_07(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 7 页：《07 交织顺序与重复单元的结构证据》（上图 + 中间贯穿大箭头流程 + 下方 3 条要点）
    header_bar(slide, theme, "交织顺序与重复单元的结构证据", 7)
    # 1. 上半部并排插入 3 张结构拆解图（44.6 Å 六边形孔道、1->4 编织步骤、四网络重复单元）
    picture_fit(slide, fig_dir / "s07_hex_pore.png", Inches(0.45), Inches(0.96), Inches(3.2), Inches(2.15))
    picture_fit(slide, fig_dir / "s07_weave_steps.png", Inches(4.35), Inches(0.96), Inches(3.8), Inches(2.15))
    picture_fit(slide, fig_dir / "s07_quad_net.png", Inches(8.65), Inches(0.96), Inches(4.3), Inches(2.15))

    # 2. 中间贯穿全宽的浅紫色大右箭头背景衬底
    shape_box(slide, Inches(0.18), Inches(3.15), SLIDE_W - Inches(0.36), Inches(1.52),
              fill=theme.soft, kind=MSO_SHAPE.RIGHT_ARROW)
    # 在大箭头上依次摆放 4 个深紫色圆角胶囊步骤
    pills = ["一套hcb网络", "两套网络交叠", "三轴节点形成", "四网络织物"]
    for i, lbl in enumerate(pills):
        px = Inches(0.32) + i * Inches(3.15)
        text_box(slide, px, Inches(3.55), Inches(2.75), Inches(0.66), lbl, theme, size=22,
                 bold=True, color=theme.white, align="center", fill=theme.primary,
                 kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        if i < 3:
            shape_box(slide, px + Inches(2.8), Inches(3.72), Inches(0.3), Inches(0.32),
                      fill="C8A8D8", kind=MSO_SHAPE.RIGHT_ARROW)

    # 3. 下半部 3 条结构证据菱形要点（字号 20 磅，行距高 0.75 英寸）
    bullets(slide, theme, [
        "六边形通道直径**44.6 Å**；大环位于不同平面，\nDB2桥联呈之字形，使上、下网络交织。",
        "以**1→2→3→4**套网络图解还原交织顺序，多网络共同构成规则的三轴编织结构。",
        "重复单元由四网络构成；弱C–H···π与C–H···O作用参与稳定 (R1, Fig.2c–e) 。",
    ], Inches(0.35), Inches(4.92), Inches(12.5), style="diamond", size=20, row_h=Inches(0.75))


def build_slide_08(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 8 页：《08 三轴编织诱导层内一致的拓扑手性》（上半部三列递进 + 下半部左右竖栏夹心机制说明）
    header_bar(slide, theme, "三轴编织诱导层内一致的拓扑手性", 8)
    # 1. 上半部三列：纳米拓扑结 -> A层P构型 -> B层M构型
    cols = [
        ("纳米拓扑结呈手性", "拓扑结的两种镜像", "s08_chirality_knots.png"),
        ("A层节点均为P构型", "四类节点同一手性", "s08_layer_p.png"),
        ("B层节点均为M构型", "与A层互为镜像", "s08_layer_m.png"),
    ]
    for i, (head, sub, img) in enumerate(cols):
        x = Inches(0.22) + i * Inches(4.38)
        w = Inches(4.12)
        # 渐变表头框
        h_box = shape_box(slide, x, Inches(0.98), w, Inches(0.58), fill=theme.primary)
        gradient_fill(h_box, [theme.primary, theme.soft], 0)
        text_box(slide, x, Inches(0.98), w, Inches(0.58), head, theme, size=22,
                 bold=True, color=theme.white, align="center")
        # 浅灰副标题条
        text_box(slide, x + Inches(0.25), Inches(1.62), w - Inches(0.5), Inches(0.42), sub,
                 theme, size=17, bold=True, align="center", fill="EEEEF0", line=theme.line)
        # 手性配图
        picture_fit(slide, fig_dir / img, x + Inches(0.2), Inches(2.12), w - Inches(0.4), Inches(1.42))
    # 三列之间的过渡文字标签（尺度拓展 / 镜像对应）
    for i, lbl in enumerate(["尺度拓展", "镜像对应"]):
        ax = Inches(3.65) + i * Inches(4.4)
        text_box(slide, ax, Inches(2.55), Inches(1.2), Inches(0.45), f"**{lbl}**", theme,
                 size=17, bold=True, align="center")

    # 2. 中部通栏核心定理横幅
    text_box(slide, Inches(0.22), Inches(3.62), SLIDE_W - Inches(0.44), Inches(0.58),
             "同层四类节点具有相同构型，相邻A/B层分别呈**P/M**手性",
             theme, size=24, bold=True, align="center", fill="EAE4F0")

    # 3. 下半部左右两侧深紫色竖向标签（“拓扑手性来源” 与 “构型镜像关系”）
    text_box(slide, Inches(0.22), Inches(4.35), Inches(0.92), Inches(1.9), "拓扑\n手性\n来源",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary)
    text_box(slide, SLIDE_W - Inches(1.14), Inches(4.35), Inches(0.92), Inches(1.9), "构型\n镜像\n关系",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary)
    # 夹在中间的两条核心机制说明与底部 3 个灰色标签块
    text_box(slide, Inches(1.45), Inches(4.35), Inches(10.4), Inches(1.05),
             "»  同一层中，**DB2**之字形结构的**扭转方向**一致，节点手性随之统一。\n"
             "»  A层为**P**构型，B层为**M**构型；交织拓扑使手性扩展到整层网络。",
             theme, size=20, bold=True)
    connector(slide, Inches(1.35), Inches(5.48), SLIDE_W - Inches(1.35), Inches(5.48), theme.accent, 1.5, True)
    for i, pill in enumerate(["层内构型一致", "相邻层互为镜像", "手性来自交织拓扑"]):
        text_box(slide, Inches(1.35) + i * Inches(3.55), Inches(5.65), Inches(3.35), Inches(0.56),
                 pill, theme, size=18, bold=True, align="center", fill="EEEEF0", line=theme.line)

    # 4. 页面底栏总结
    banner(slide, theme, "从纳米结到网状织物：同层同手性，相邻层P/M镜像 (R1, Fig.3a,c)",
           Inches(0.22), Inches(6.48), SLIDE_W - Inches(0.44), Inches(0.68), style="bar", size=22)


def build_slide_09(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 9 页：《09 剥离纳米片揭示单层与双层厚度》（三列浅紫底图文卡片 + 底部大红字结论）
    header_bar(slide, theme, "剥离纳米片揭示单层与双层厚度", 9)
    # 顶部 2 行实心紫箭头要点
    bullets(slide, theme, [
        "液相超声剥离可获得横向尺寸**数百纳米**的TWPN纳米片。",
        "AFM测得约**2.0 nm**（单层）与**4.0 nm**（双层），与晶体结构厚度一致。",
    ], Inches(0.28), Inches(0.98), Inches(12.5), style="arrow", size=22, row_h=Inches(0.5))

    # 三列并排 AFM 表征卡片
    cards = [
        ("液相超声剥离", "超声处理促进层间分离\n横向尺寸约数百纳米", "s09_exfoliation.png"),
        ("单层纳米片", "AFM台阶约2.0 nm\n与单层晶体结构相符", "s09_afm_mono.png"),
        ("双层纳米片", "AFM台阶约4.0 nm\n约为单层台阶的两倍", "s09_afm_bi.png"),
    ]
    for i, (head, body, img) in enumerate(cards):
        x = Inches(0.22) + i * Inches(4.32)
        w = Inches(4.15)
        shape_box(slide, x, Inches(2.05), w, Inches(4.45), fill=theme.soft)               # 外层浅紫大卡片
        shape_box(slide, x + Inches(0.14), Inches(2.15), w - Inches(0.28), Inches(2.75), fill=theme.white) # 内嵌白底图框
        picture_fit(slide, fig_dir / img, x + Inches(0.2), Inches(2.18), w - Inches(0.4), Inches(2.68))
        text_box(slide, x + Inches(0.2), Inches(5.0), w - Inches(0.4), Inches(0.48), head,
                 theme, size=22, bold=True)
        text_box(slide, x + Inches(0.2), Inches(5.48), w - Inches(0.4), Inches(0.95), body,
                 theme, size=18, bold=True)

    # 底部红色大字定量结论
    banner(slide, theme, "3.9×10^{4} μm^{2}单层估计含 > 177亿同手性单元 (R1)",
           Inches(0.5), Inches(6.62), SLIDE_W - Inches(1.0), Inches(0.65), style="text", size=26)


def build_slide_10(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 10 页：《10 柔软织物仍表现出显著弹性回弹》（上排条件卡 + 原生压痕示意，下排实验曲线与定量对照）
    header_bar(slide, theme, "柔软织物仍表现出显著弹性回弹", 10)
    # 1. 左上卡片：纳米压痕条件
    shape_box(slide, Inches(0.22), Inches(0.98), Inches(3.85), Inches(1.92), fill="EDE7F3", line=theme.line)
    text_box(slide, Inches(0.32), Inches(1.04), Inches(3.65), Inches(0.45), "纳米压痕条件",
             theme, size=22, bold=True, align="center")
    text_box(slide, Inches(0.38), Inches(1.52), Inches(3.5), Inches(1.3),
             "➢  (001)晶面受压\n➢  最大载荷2.0 mN\n➢  记录加载与卸载", theme, size=18, bold=True)

    # 2. 中上卡片：柔软与回弹指标
    shape_box(slide, Inches(4.22), Inches(0.98), Inches(3.95), Inches(1.92), fill="EDE7F3", line=theme.line)
    text_box(slide, Inches(4.32), Inches(1.04), Inches(3.75), Inches(0.45), "**柔软与回弹**",
             theme, size=22, bold=True, align="center")
    text_box(slide, Inches(4.38), Inches(1.5), Inches(3.7), Inches(1.35),
             "➢  约化模量4.27 GPa\n     处于最软25%有机晶体范围\n➢  **回弹420 nm，70.3%**",
             theme, size=17, bold=True)

    # 3. 右上：用 PowerPoint 原生矢量箭头和平行四边形绘制“(001)面纳米压痕示意图”
    rx = Inches(8.6)
    text_box(slide, rx, Inches(1.02), Inches(3.8), Inches(0.35), "(001)面纳米压痕 | 示意",
             theme, size=14, align="center")
    shape_box(slide, rx + Inches(1.8), Inches(1.4), Inches(0.18), Inches(0.55),
              fill=theme.primary, kind=MSO_SHAPE.DOWN_ARROW)
    text_box(slide, rx + Inches(2.1), Inches(1.48), Inches(2.0), Inches(0.35), "最大载荷 2.0 mN",
             theme, size=13)
    shape_box(slide, rx + Inches(0.5), Inches(2.05), Inches(2.8), Inches(0.35),
              fill="C589D4", line=theme.primary, kind=MSO_SHAPE.PARALLELOGRAM)
    text_box(slide, rx + Inches(0.5), Inches(2.45), Inches(2.8), Inches(0.35), "TWPN 晶体",
             theme, size=14, align="center")

    # 4. 左下卡片：加载与卸载曲线
    text_box(slide, Inches(0.22), Inches(3.05), Inches(3.85), Inches(0.56), "加载与卸载曲线",
             theme, size=22, bold=True, color=theme.white, align="center", fill=theme.primary,
             kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.18)
    shape_box(slide, Inches(0.22), Inches(3.65), Inches(3.85), Inches(2.98), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s10_load_curve.png", Inches(0.32), Inches(3.72), Inches(3.65), Inches(2.84))
    text_box(slide, Inches(0.22), Inches(6.72), Inches(3.85), Inches(0.5), "**2.0mN下结构保持机械稳定**",
             theme, size=18, bold=True, align="center")

    # 5. 中下卡片：约化模量与回弹的定量对照
    text_box(slide, Inches(4.22), Inches(3.05), Inches(5.85), Inches(0.56), "约化模量与回弹的定量对照",
             theme, size=22, bold=True, color=theme.white, align="center", fill=theme.primary,
             kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.18)
    shape_box(slide, Inches(4.22), Inches(3.65), Inches(5.85), Inches(2.98), fill=theme.white, line=theme.line)
    picture_fit(slide, fig_dir / "s10_modulus_rebound.png", Inches(4.32), Inches(3.72), Inches(5.65), Inches(2.84))
    text_box(slide, Inches(4.22), Inches(6.72), Inches(5.85), Inches(0.5), "**TWPN模量高于未编织单体**",
             theme, size=18, bold=True, align="center")

    # 6. 右下：回弹表现与比较边界说明文字
    text_box(slide, Inches(10.35), Inches(3.35), Inches(2.75), Inches(1.75),
             "**回弹表现**\n回弹深度420nm\n回弹比例70.3%\n(R1, ED Fig.1b)", theme, size=18, bold=True)
    text_box(slide, Inches(10.35), Inches(5.35), Inches(2.8), Inches(1.5),
             "**比较边界**\n约化模量并非拉伸强度\n结果限于本压痕条件", theme, size=16, bold=True)


def build_slide_11(slide, theme: Theme, fig_dir: Path):
    # 【函数作用】画第 11 页：《11 大环与编织节点协同分散应力》（左栏结构基础 + 右栏三分支多尺度模拟树）
    header_bar(slide, theme, "大环与编织节点协同分散应力", 11)
    bullets(slide, theme, [
        "分子模拟显示：c轴压缩**约20%**时，a/b方向几乎不膨胀。",
        "从**结构形变**、**角度变化**与**运动耦合**三个层面解释弹性响应。",
    ], Inches(0.35), Inches(0.95), Inches(12.5), style="diamond", size=20, row_h=Inches(0.42))

    # --- 1. 左栏：大环与节点的运动分工 ---
    lx, lw = Inches(0.22), Inches(3.2)
    text_box(slide, lx, Inches(1.88), Inches(1.65), Inches(0.5), "结构基础", theme, size=18,
             bold=True, color=theme.white, align="center", fill=theme.primary)
    text_box(slide, lx + Inches(1.65), Inches(1.88), Inches(1.55), Inches(0.5), "刚柔协同", theme,
             size=18, bold=True, align="center", fill=theme.soft)
    text_box(slide, lx, Inches(2.42), lw, Inches(0.42), "**大环与节点的运动分工**",
             theme, size=18, bold=True, align="center")
    picture_fit(slide, fig_dir / "s11_node_schem.png", lx + Inches(0.1), Inches(2.9), lw - Inches(0.2), Inches(1.4))
    text_box(slide, lx, Inches(4.35), lw, Inches(0.4), "可变构象与局部运动", theme, size=17, bold=True, align="center")
    text_box(slide, lx, Inches(4.82), Inches(1.55), Inches(1.45), "大环率先\n发生弯曲\n\n刚性骨架",
             theme, size=16, bold=True, align="center")
    text_box(slide, lx + Inches(1.65), Inches(4.82), Inches(1.55), Inches(1.45), "节点整体\n几何较稳定\n\n柔性调适",
             theme, size=16, bold=True, align="center")

    # --- 2. 右栏：分层形变与多尺度模拟三分支 ---
    rx, rw = Inches(3.65), Inches(9.45)
    text_box(slide, rx, Inches(1.88), rw, Inches(0.52), "大环先弯曲，节点随后调整：分层形变促进应力再分配",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary)
    text_box(slide, rx + Inches(2.1), Inches(2.5), Inches(5.25), Inches(0.45), "模拟证据与解释范围",
             theme, size=19, bold=True, color=theme.white, align="center", fill="9E1B32")
    sub_cards = [
        "压缩前后结构对照\n识别大环与节点形变",
        "压缩量与角度响应\n追踪α、β、γ与θ",
        "关联大环与节点运动\n分析局部运动耦合",
    ]
    for i, sc in enumerate(sub_cards):
        sx = rx + i * Inches(3.2)
        text_box(slide, sx, Inches(3.15), Inches(3.05), Inches(0.72), sc, theme,
                 size=15, bold=True, align="center", fill=theme.soft)

    text_box(slide, rx + Inches(3.1), Inches(3.92), Inches(3.2), Inches(0.38), "多尺度模拟机制解释",
             theme, size=18, bold=True, align="center")
    bot_cols = [
        ("结构形变", "s11_comp_struct.png", MSO_SHAPE.PENTAGON),
        ("角度响应", "s11_angle_resp.png", MSO_SHAPE.RECTANGLE),
        ("运动耦合", "s11_corr_map.png", MSO_SHAPE.CHEVRON),
    ]
    for i, (head, img, shp_kind) in enumerate(bot_cols):
        bx = rx + i * Inches(3.2)
        text_box(slide, bx, Inches(4.38), Inches(3.0), Inches(0.48), head, theme,
                 size=18, bold=True, color=theme.white, align="center", fill=theme.primary, kind=shp_kind)
        picture_fit(slide, fig_dir / img, bx + Inches(0.05), Inches(4.92), Inches(2.9), Inches(1.52))

    # 3. 底部紫色总结条
    banner(slide, theme, "结构、角度与运动耦合共同构成模拟解释，支持应力分散与回弹 (R1, Fig.4b–e)",
           Inches(0.22), Inches(6.58), SLIDE_W - Inches(0.44), Inches(0.58), style="bar", size=19)
