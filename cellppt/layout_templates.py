# ==============================================================================
# 【文件作用】v0.2 百变学术版式模板库 (cellppt/layout_templates.py)
# 从对标视频中提炼出的 8 种不同排版结构，全部接收标准字典数据渲染，彻底告别“每页排版都一样”。
# ==============================================================================

from pathlib import Path
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from .blocks import banner, bullets, chart
from .builder import connector, gradient_fill, picture_fit, shape_box, text_box
from .layouts_basic import header_bar
from .theme import SLIDE_W, Theme


def _draw_visual_or_cards(slide, theme: Theme, box_x, box_y, box_w, box_h, item: dict):
    # 【函数作用】在指定区域内：如果有配图路径就无损插图；如果没配图，就自动画原生矢量框图/要点卡，绝不空洞或乱塞校徽
    fig = item.get("figure")
    if fig and Path(fig).is_file():
        picture_fit(slide, fig, box_x + Inches(0.08), box_y + Inches(0.06),
                    box_w - Inches(0.16), box_h - Inches(0.12))
        return
    # 无位图时：自动绘制原生矢量对比子卡片列表
    sub_items = item.get("points", [])
    if not sub_items:
        return
    n = len(sub_items)
    gap = Inches(0.12)
    ch = (box_h - gap * (n + 1)) / max(1, n)
    for i, pt in enumerate(sub_items):
        cy = box_y + gap + i * (ch + gap)
        # 左边深色圆角标签 + 右边浅紫内容条（全原生矢量形状）
        if "：" in pt or ":" in pt:
            sep = "：" if "：" in pt else ":"
            k, v = pt.split(sep, 1)
            k = k.replace("**", "").strip()
            text_box(slide, box_x + Inches(0.12), cy, Inches(1.48), ch, k, theme,
                     size=15, bold=True, color=theme.white, align="center",
                     fill=theme.primary if i % 2 == 0 else theme.accent,
                     kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
            text_box(slide, box_x + Inches(1.68), cy, box_w - Inches(1.8), ch, v.strip(), theme,
                     size=14, bold=True, fill=theme.softer, line=theme.line, line_spacing=1.05)
        else:
            text_box(slide, box_x + Inches(0.15), cy, box_w - Inches(0.3), ch, pt, theme,
                     size=16, bold=True, align="center", fill=theme.softer, line=theme.line)


def tpl_dual_compare(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 1：双栏深度对比版式】（对标视频 Slide 3：顶部2行菱形要点 + 左右两大对比卡片 + 卡片底部结论条）
    header_bar(slide, theme, data["title"], s_no)
    bullets(slide, theme, data.get("bullets", []), Inches(0.35), Inches(1.02),
            Inches(12.5), style="diamond", size=20, row_h=Inches(0.45))
    y_card, h_card, w_card = Inches(2.0), Inches(5.15), Inches(6.15)
    for idx, col in enumerate(data.get("columns", [])[:2]):
        x = Inches(0.32) if idx == 0 else Inches(6.82)
        shape_box(slide, x, y_card, w_card, h_card, fill=theme.white, line=theme.line)
        text_box(slide, x, y_card, w_card, Inches(0.62), col.get("header", ""),
                 theme, size=23, bold=True, align="center", fill=theme.soft)
        text_box(slide, x + Inches(0.2), Inches(2.7), w_card - Inches(0.4), Inches(0.75),
                 col.get("subtitle", ""), theme, size=19, bold=True, align="center")
        _draw_visual_or_cards(slide, theme, x + Inches(0.15), Inches(3.5),
                              w_card - Inches(0.3), Inches(2.7), col)
        text_box(slide, x + Inches(0.22), Inches(6.32), w_card - Inches(0.44), Inches(0.62),
                 col.get("footer", ""), theme, size=17, bold=True, align="center", fill="EAE8EE")


def tpl_three_tier_cards(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 2：平行四边形悬浮头三列递进版式】（对标视频 Slide 6：三列大卡片 + 平行四边形徽章 + 底部对称渐变条）
    header_bar(slide, theme, data["title"], s_no)
    for i, col in enumerate(data.get("columns", [])[:3]):
        x = Inches(0.32) + i * Inches(4.28)
        w = Inches(4.08)
        shape_box(slide, x, Inches(1.38), w, Inches(4.85), fill=theme.white, line="B8B8B8")
        # 顶部跨界平行四边形标签（如：阶段一 / 阶段二 / 阶段三）
        text_box(slide, x + Inches(0.55), Inches(1.02), w - Inches(1.1), Inches(0.64),
                 col.get("tag", f"模块 {i+1}"), theme, size=22, bold=True, color=theme.white,
                 align="center", fill=theme.primary, kind=MSO_SHAPE.PARALLELOGRAM)
        text_box(slide, x + Inches(0.15), Inches(1.78), w - Inches(0.3), Inches(0.6),
                 col.get("header", ""), theme, size=22, bold=True, align="center")
        shape_box(slide, x + Inches(0.2), Inches(2.48), w - Inches(0.4), Inches(3.52),
                  fill=theme.white, line=theme.line)
        _draw_visual_or_cards(slide, theme, x + Inches(0.22), Inches(2.52),
                              w - Inches(0.44), Inches(3.44), col)
    bar = shape_box(slide, Inches(0.32), Inches(6.35), SLIDE_W - Inches(0.64), Inches(0.82), fill=theme.soft)
    gradient_fill(bar, [theme.soft, theme.primary, theme.soft], 0)
    text_box(slide, Inches(0.5), Inches(6.38), SLIDE_W - Inches(1.0), Inches(0.76),
             data.get("banner", ""), theme, size=22, bold=True, align="center")


def tpl_tri_zone_flow(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 3：左-中-右三区流程与系统框图版式】（对标视频 Slide 4：左栏核心模块 + 中栏4步竖向箭头流 + 右栏系统大图与子卡）
    header_bar(slide, theme, data["title"], s_no)
    bullets(slide, theme, data.get("bullets", []), Inches(0.35), Inches(0.98),
            Inches(12.5), style="dot", size=19, row_h=Inches(0.42))
    # 1. 左区面板
    lx, ly, lw, lh = Inches(0.22), Inches(1.92), Inches(4.85), Inches(4.62)
    left = data.get("left", {})
    shape_box(slide, lx, ly, lw, lh, fill=theme.white, line=theme.line)
    text_box(slide, lx, ly, Inches(1.6), Inches(0.54), left.get("tag", "系统基础"), theme,
             size=18, bold=True, color=theme.white, align="center", fill=theme.primary)
    text_box(slide, lx + Inches(1.6), ly, lw - Inches(1.6), Inches(0.54), left.get("header", ""),
             theme, size=18, bold=True, align="center", fill=theme.soft)
    text_box(slide, lx + Inches(0.15), Inches(2.52), lw - Inches(0.3), Inches(0.45),
             left.get("highlight", ""), theme, size=19, bold=True, align="center")
    _draw_visual_or_cards(slide, theme, lx + Inches(0.12), Inches(3.05), lw - Inches(0.24), Inches(2.45), left)
    chips = left.get("chips", [])
    for i, ch_txt in enumerate(chips[:3]):
        text_box(slide, lx + Inches(0.14) + i * Inches(1.55), Inches(5.6), Inches(1.46), Inches(0.8),
                 ch_txt, theme, size=14, bold=True, align="center", fill=theme.softer)
    # 2. 中区 4 步竖向流程
    mx, mw = Inches(5.22), Inches(1.65)
    for i, st in enumerate(data.get("flow_steps", [])[:4]):
        sy = Inches(1.92) + i * Inches(1.26)
        text_box(slide, mx, sy, mw, Inches(0.82), st, theme, size=15, bold=True,
                 align="center", fill=theme.white, line=theme.line)
        if i < 3:
            shape_box(slide, mx + Inches(0.62), sy + Inches(0.86), Inches(0.38), Inches(0.34),
                      fill=theme.soft, kind=MSO_SHAPE.DOWN_ARROW)
    # 3. 右区系统架构图 + 双子卡
    rx, rw = Inches(7.02), Inches(6.05)
    right = data.get("right", {})
    text_box(slide, rx, Inches(1.92), rw, Inches(0.54), right.get("header", ""),
             theme, size=19, bold=True, color=theme.white, align="center", fill=theme.primary)
    if right.get("figure") and Path(right["figure"]).is_file():
        picture_fit(slide, right["figure"], rx + Inches(0.1), Inches(2.52), rw - Inches(0.2), Inches(2.65))
    for i, sc in enumerate(right.get("subcards", [])[:2]):
        bx = rx + Inches(0.1) + i * Inches(2.98)
        text_box(slide, bx, Inches(5.25), Inches(2.88), Inches(0.42), sc.get("title", ""),
                 theme, size=16, bold=True, align="center", fill=theme.soft)
        text_box(slide, bx, Inches(5.7), Inches(2.88), Inches(0.82), sc.get("body", ""),
                 theme, size=14, bold=True, align="center", fill="F5F3F7")
    banner(slide, theme, data.get("banner", ""), Inches(0.22), Inches(6.68),
           SLIDE_W - Inches(0.44), Inches(0.58), style="bar", size=19)



def tpl_arch_innovation(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 4：核心架构对比大图 + 右侧模块深度拆解版式】（左栏放核心创新架构图，右栏拆解对抗机制与公式要点）
    header_bar(slide, theme, data["title"], s_no)
    bullets(slide, theme, data.get("bullets", []), Inches(0.35), Inches(0.98),
            Inches(12.5), style="arrow", size=19, row_h=Inches(0.42))
    # 左栏：架构大图卡片
    lx, lw = Inches(0.28), Inches(6.15)
    text_box(slide, lx, Inches(1.92), lw, Inches(0.54), data.get("left_header", "核心网络架构对比"),
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary,
             kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.15)
    shape_box(slide, lx, Inches(2.48), lw, Inches(4.02), fill=theme.white, line=theme.line)
    if data.get("figure") and Path(data["figure"]).is_file():
        picture_fit(slide, data["figure"], lx + Inches(0.12), Inches(2.54), lw - Inches(0.24), Inches(3.9))
    # 右栏：3 层核心机制拆解卡
    rx, rw = Inches(6.65), Inches(6.4)
    for i, m in enumerate(data.get("modules", [])[:3]):
        ry = Inches(1.92) + i * Inches(1.52)
        text_box(slide, rx, ry, Inches(1.75), Inches(1.36), m.get("tag", ""), theme,
                 size=18, bold=True, color=theme.white, align="center",
                 fill=theme.primary if i != 1 else theme.accent)
        text_box(slide, rx + Inches(1.85), ry, rw - Inches(1.85), Inches(1.36), m.get("desc", ""),
                 theme, size=16, bold=True, fill=theme.softer, line=theme.line)
    banner(slide, theme, data.get("banner", ""), Inches(0.28), Inches(6.65),
           SLIDE_W - Inches(0.56), Inches(0.58), style="bar", size=19)


def tpl_focus_challenges(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 5：研究重点 + 竖栏夹心三大难点版式】（对标视频 Slide 8：上半部重点卡片 + 中间横幅 + 下半部左右竖栏夹三列难点）
    header_bar(slide, theme, data["title"], s_no)
    for i, fc in enumerate(data.get("focus_cards", [])[:2]):
        x = Inches(0.25) + i * Inches(6.48)
        w = Inches(6.35)
        h_box = shape_box(slide, x, Inches(1.0), w, Inches(0.54), fill=theme.primary)
        gradient_fill(h_box, [theme.primary, theme.soft], 0)
        text_box(slide, x, Inches(1.0), w, Inches(0.54), fc.get("title", ""), theme,
                 size=20, bold=True, color=theme.white, align="center")
        text_box(slide, x, Inches(1.58), w, Inches(1.72), fc.get("body", ""), theme,
                 size=17, bold=True, fill=theme.softer, line=theme.line)
    # 中部通栏核心过渡条
    text_box(slide, Inches(0.25), Inches(3.48), SLIDE_W - Inches(0.5), Inches(0.55),
             data.get("mid_banner", ""), theme, size=21, bold=True, align="center", fill="EAE4F0")
    # 下半部左右两侧深色竖向标签 + 中间三张难点卡片
    text_box(slide, Inches(0.25), Inches(4.18), Inches(0.92), Inches(2.22), "三大\n关键\n难点",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary)
    text_box(slide, SLIDE_W - Inches(1.17), Inches(4.18), Inches(0.92), Inches(2.22), "攻关\n突破\n方向",
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.accent)
    for i, ch in enumerate(data.get("challenges", [])[:3]):
        cx = Inches(1.32) + i * Inches(3.62)
        cw = Inches(3.45)
        text_box(slide, cx, Inches(4.18), cw, Inches(0.5), ch.get("title", ""), theme,
                 size=17, bold=True, align="center", fill=theme.soft)
        text_box(slide, cx, Inches(4.72), cw, Inches(1.68), ch.get("body", ""), theme,
                 size=15, bold=True, fill=theme.white, line=theme.line)
    banner(slide, theme, data.get("banner", ""), Inches(0.25), Inches(6.55),
           SLIDE_W - Inches(0.5), Inches(0.62), style="bar", size=19)


def tpl_four_grid_metrics(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 6：四卡片并排创新网格 + 底部定量指标栏版式】（对标视频 Slide 5）
    header_bar(slide, theme, data["title"], s_no)
    text_box(slide, Inches(0.4), Inches(0.98), Inches(12.5), Inches(0.52),
             data.get("lead", ""), theme, size=21, bold=True)
    text_box(slide, Inches(1.2), Inches(1.55), Inches(10.9), Inches(0.52),
             data.get("highlight", ""), theme, size=23, bold=True, align="center")
    for i, c in enumerate(data.get("cards", [])[:4]):
        x = Inches(0.25) + i * Inches(3.25)
        w = Inches(3.05)
        text_box(slide, x, Inches(2.22), w, Inches(0.56), c.get("title", ""), theme,
                 size=20, bold=True, color=theme.white, align="center", fill=theme.primary,
                 kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.18)
        shape_box(slide, x, Inches(2.82), w, Inches(2.65), fill=theme.white, line=theme.line)
        _draw_visual_or_cards(slide, theme, x + Inches(0.06), Inches(2.88), w - Inches(0.12), Inches(2.52), c)
    connector(slide, Inches(0.25), Inches(5.68), SLIDE_W - Inches(0.25), Inches(5.68), theme.line, 1.5, False)
    for i, m in enumerate(data.get("metrics", [])[:4]):
        mx = Inches(0.25) + i * Inches(3.25)
        text_box(slide, mx, Inches(5.85), Inches(3.05), Inches(1.12), m, theme,
                 size=18, bold=True, align="center",
                 fill=theme.softer if i == 0 else theme.white, line=theme.line,
                 kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)


def tpl_data_chart_and_shot(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 7：左栏 PowerPoint 原生可编辑柱状图 + 右栏实验/系统实景截图版式】（对标视频 Slide 10/14）
    header_bar(slide, theme, data["title"], s_no)
    bullets(slide, theme, data.get("bullets", []), Inches(0.35), Inches(0.98),
            Inches(12.5), style="arrow", size=19, row_h=Inches(0.42))
    # 左栏：PowerPoint 原生可编辑数据图表 + 结论说明
    lx, lw = Inches(0.28), Inches(5.85)
    shape_box(slide, lx, Inches(1.92), lw, Inches(4.58), fill=theme.softer, line=theme.line)
    text_box(slide, lx + Inches(0.15), Inches(1.98), lw - Inches(0.3), Inches(0.5),
             data.get("chart_header", "预实验性能定量对比"), theme, size=20, bold=True,
             align="center", fill=theme.soft)
    shape_box(slide, lx + Inches(0.15), Inches(2.55), lw - Inches(0.3), Inches(2.95),
              fill=theme.white, line=theme.line)
    if data.get("chart_spec"):
        chart(slide, theme, data["chart_spec"], lx + Inches(0.25), Inches(2.62),
              lw - Inches(0.5), Inches(2.8))
    text_box(slide, lx + Inches(0.18), Inches(5.58), lw - Inches(0.36), Inches(0.82),
             data.get("chart_note", ""), theme, size=15, bold=True, align="center")
    # 右栏：实景截图/系统原型图
    rx, rw = Inches(6.35), Inches(6.7)
    text_box(slide, rx, Inches(1.92), rw, Inches(0.54), data.get("right_header", "实验验证与原型截图"),
             theme, size=20, bold=True, color=theme.white, align="center", fill=theme.primary,
             kind=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, radius=0.15)
    shape_box(slide, rx, Inches(2.48), rw, Inches(4.02), fill=theme.white, line=theme.line)
    if data.get("figure") and Path(data["figure"]).is_file():
        picture_fit(slide, data["figure"], rx + Inches(0.12), Inches(2.55), rw - Inches(0.24), Inches(3.15))
    text_box(slide, rx + Inches(0.15), Inches(5.78), rw - Inches(0.3), Inches(0.62),
             data.get("right_note", ""), theme, size=16, bold=True, align="center", fill=theme.softer)
    banner(slide, theme, data.get("banner", ""), Inches(0.3), Inches(6.62),
           SLIDE_W - Inches(0.6), Inches(0.58), style="text", size=22)


def tpl_roadmap_solutions(slide, theme: Theme, s_no: int, data: dict):
    # 【版式 8：现存问题对策对照 + 贯穿全屏五阶段时间轴胶囊链版式】（对标视频 Slide 7 + Slide 15）
    header_bar(slide, theme, data["title"], s_no)
    for i, head in enumerate(["现存瓶颈问题", "针对性解决途径"]):
        hx = Inches(0.25) + i * Inches(6.6)
        bar = shape_box(slide, hx, Inches(0.95), Inches(6.22), Inches(0.48),
                        fill=theme.soft, kind=MSO_SHAPE.TRAPEZOID)
        gradient_fill(bar, ["FFFFFF", theme.soft], 90)
        text_box(slide, hx, Inches(0.95), Inches(6.22), Inches(0.48), head,
                 theme, size=20, bold=True, align="center")
    for r, pair in enumerate(data.get("pairs", [])[:3]):
        ry = Inches(1.5) + r * Inches(0.96)
        text_box(slide, Inches(0.25), ry, Inches(6.22), Inches(0.86), pair[0],
                 theme, size=16, bold=True, fill=theme.white, line=theme.line)
        shape_box(slide, Inches(6.54), ry + Inches(0.22), Inches(0.26), Inches(0.42),
                  fill=theme.soft, kind=MSO_SHAPE.RIGHT_ARROW)
        text_box(slide, Inches(6.85), ry, Inches(6.22), Inches(0.86), pair[1],
                 theme, size=16, bold=True, fill=theme.softer, line=theme.line)
    # 下半部：贯穿全屏的大右箭头时间轴（支持 4~5 个阶段胶囊）
    shape_box(slide, Inches(0.2), Inches(4.52), SLIDE_W - Inches(0.4), Inches(1.92),
              fill=theme.soft, kind=MSO_SHAPE.RIGHT_ARROW)
    timeline = data.get("timeline", [])[:5]
    n_t = max(1, len(timeline))
    step_w = (SLIDE_W - Inches(1.1)) / n_t
    for i, tp in enumerate(timeline):
        px = Inches(0.35) + i * step_w
        pw = step_w - Inches(0.18)
        text_box(slide, px, Inches(4.72), pw, Inches(0.52), tp.get("time", ""),
                 theme, size=15, bold=True, color=theme.white, align="center",
                 fill=theme.primary, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.45)
        text_box(slide, px, Inches(5.32), pw, Inches(0.95), tp.get("task", ""),
                 theme, size=14, bold=True, align="center", fill=theme.white, line=theme.line)
    banner(slide, theme, data.get("banner", ""), Inches(0.25), Inches(6.6),
           SLIDE_W - Inches(0.5), Inches(0.58), style="bar", size=18)


# 版式名称到渲染函数的注册表
LAYOUT_REGISTRY = {
    "dual_compare": tpl_dual_compare,
    "three_tier_cards": tpl_three_tier_cards,
    "tri_zone_flow": tpl_tri_zone_flow,
    "arch_innovation": tpl_arch_innovation,
    "focus_challenges": tpl_focus_challenges,
    "four_grid_metrics": tpl_four_grid_metrics,
    "data_chart_and_shot": tpl_data_chart_and_shot,
    "roadmap_solutions": tpl_roadmap_solutions,
}
