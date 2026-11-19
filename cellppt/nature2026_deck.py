"""Assemble the 16-slide Nature 2026 presentation and companion handoff documents."""
from pathlib import Path
import pymupdf as fitz
from pptx import Presentation
from pptx.util import Inches

from .builder import shape_box, text_box
from .layouts_basic import header_bar
from .slides_1_to_6 import (
    build_slide_01, build_slide_02, build_slide_03,
    build_slide_04, build_slide_05, build_slide_06,
)
from .slides_7_to_11 import (
    build_slide_07, build_slide_08, build_slide_09,
    build_slide_10, build_slide_11,
)
from .slides_12_to_16 import (
    build_slide_12, build_slide_13, build_slide_14,
    build_slide_15, build_slide_16,
)
from .theme import SLIDE_H, SLIDE_W, get_theme

BUILDERS = [
    build_slide_01, build_slide_02, build_slide_03, build_slide_04,
    build_slide_05, build_slide_06, build_slide_07, build_slide_08,
    build_slide_09, build_slide_10, build_slide_11, build_slide_12,
    build_slide_13, build_slide_14, build_slide_15, build_slide_16,
]

SLIDE_PLAN = [
    (1, "封面：二维网状三轴分子编织", "展示论文标题、作者(Wang等, Nature 2026)、汇报时长(15分钟)与DOI", "封面版式", "0.5 min"),
    (2, "汇报目录", "01研究背景 / 02设计与结构 / 03拓扑手性 / 04力学响应 / 05孔隙与分离 / 06结论与展望", "目录版式", "0.5 min"),
    (3, "从双轴走向网状三轴分子编织", "对比已有双轴编织(ref.12, ref.16)与网状三轴研究缺口(ref.13, 本工作)", "双栏对比卡片", "1.0 min"),
    (4, "C3大环与DB2的配位组装", "MCN3(C3对称吡啶供体)与DB2二连接单体通过可逆B–N配位键驱动组装", "三区流程版式", "1.2 min"),
    (5, "80℃、24小时获得TWPN单晶", "甲苯80℃加热24h缓慢冷却获单晶；C2/c空间群，B–N键长1.65(1) Å，键能122.95 kJ/mol", "四卡片网格", "1.0 min"),
    (6, "四网络交织构成层状分子织物", "层级一：四套hcb网络交织；层级二：三网络构成节点；层级三：单层19.5 Å与AB堆积", "三列层级卡片", "1.2 min"),
    (7, "交织顺序与重复单元的结构证据", "44.6 Å六边形通道；1→2→3→4套网络图解交织顺序；C–H···π与C–H···O稳定作用", "流程横幅+证据", "1.0 min"),
    (8, "三轴编织诱导层内一致的拓扑手性", "纳米拓扑结镜像 → A层P构型 / B层M构型；DB2之字形扭转方向一致扩展至全层", "三列+雙向机制", "1.2 min"),
    (9, "剥离纳米片揭示单层与双层厚度", "液相超声剥离获数百纳米薄片；AFM台阶2.0 nm(单层)与4.0 nm(双层)；>177亿同手性单元", "三列AFM对照", "1.0 min"),
    (10, "柔软织物仍表现出显著弹性回弹", "(001)晶面2.0 mN纳米压痕；约化模量4.27 GPa，回弹420 nm (70.3%)", "压痕曲线+对照", "1.2 min"),
    (11, "大环与编织节点协同分散应力", "c轴压缩20%时a/b几乎不膨胀；大环先弯曲、节点后调整；结构形变/角度响应/运动耦合", "层级树+多尺度模拟", "1.2 min"),
    (12, "内源孔与编织协同孔构成多孔框架", "大环内源孔9.1 Å + 编织协同孔12.7 Å；195 K CO2吸附确认BET表面积91.6 m2/g", "左右互证版式", "1.0 min"),
    (13, "苯优先吸附与动态分离", "苯/环己烷等温吸附、竞争吸附与穿透曲线；TWPN释放蒸气Ben达97.9% (MCN3为60.5%)", "曲线+圆环图", "1.2 min"),
    (14, "多组分离、热再生与百克级制备", "六组芳香/脂肪烃配对均达92.1%–98.0%；100 g级放大制备，产率87.2%", "柱状图+百克实物", "1.0 min"),
    (15, "结论与展望：拓扑成为材料设计变量", "已建立的认识 vs 待验证的问题（同组分拓扑对照、循环压缩与湿度多组分穿透）", "双列总结展望", "0.6 min"),
    (16, "致谢与交流讨论", "感谢聆听，欢迎交流讨论；讨论问题：如何进一步提高拓扑功能？", "封底致谢版式", "0.2 min"),
]


def default_fig_dir() -> Path:
    return Path(__file__).resolve().parent.parent / "assets" / "nature2026_figures"


def generate_nature2026_pptx(output_path: str | Path, theme_name: str = "nature_purple",
                             fig_dir: Path | None = None) -> Path:
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig_dir = Path(fig_dir) if fig_dir else default_fig_dir()
    theme = get_theme(theme_name)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for builder in BUILDERS:
        slide = prs.slides.add_slide(blank)
        builder(slide, theme, fig_dir)

    try:
        prs.save(str(output_path))
    except PermissionError:
        import time
        output_path = output_path.with_name(f"{output_path.stem}_{int(time.time())}.pptx")
        prs.save(str(output_path))
    return output_path


def generate_companion_files(output_dir: str | Path, theme_name: str = "nature_purple") -> dict:
    """Create 《逐页安排.pptx》 and 《问答与原始资料.pdf》 matching the cell_ppt workflow."""
    out = Path(output_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    theme = get_theme(theme_name)

    # 1. 《逐页安排.pptx》
    plan_pptx = out / "逐页安排.pptx"
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    for num, title, summary, layout_kind, dur in SLIDE_PLAN:
        s = prs.slides.add_slide(blank)
        header_bar(s, theme, f"逐页规划 · 第 {num:02d} 页：{title}", num)
        shape_box(s, Inches(0.45), Inches(1.25), Inches(12.4), Inches(5.65), fill=theme.softer, line=theme.line)
        text_box(s, Inches(0.75), Inches(1.5), Inches(11.8), Inches(4.8),
                 f"【页面标题】 {title}\n\n"
                 f"【核心内容与证据定位】\n{summary}\n\n"
                 f"【版式结构】 {layout_kind}\n"
                 f"【预计讲解时长】 {dur}\n"
                 f"【主题配色】 {theme.name} (#{theme.primary} / #{theme.accent})",
                 theme, size=20, bold=False, anchor="top", line_spacing=1.35)
    try:
        prs.save(str(plan_pptx))
    except PermissionError:
        import time
        plan_pptx = out / f"逐页安排_{int(time.time())}.pptx"
        prs.save(str(plan_pptx))

    # 2. 《问答与原始资料.pdf》
    handoff_pdf = out / "问答与原始资料.pdf"
    doc = fitz.open()
    p1 = doc.new_page(width=595, height=842)
    p1.insert_text((48, 64), "PPT 总页数：16 页", fontsize=18, fontname="china-ss")
    p1.insert_text((48, 96), f"模板选择：{theme.name} (16:9 学术汇报深紫配色)", fontsize=12, fontname="china-ss")
    p1.insert_text((48, 128), "一、需求问答与模式参数", fontsize=14, fontname="china-ss")
    qa_lines = [
        "• 汇报场景：面向研究生与教师的 15 分钟中文学术文献汇报",
        "• 目标文献：Wang et al., Nature (2026) DOI: 10.1038/s41586-026-11218-8",
        "• 语言规范：中文讲解，保留原文关键英文术语、晶体学符号与图内标签",
        "• 重点覆盖：研究问题 → 配位组装设计 → 晶体结构 → 拓扑手性 → 弹性回弹 → 吸附分离",
    ]
    for idx, line in enumerate(qa_lines):
        p1.insert_text((60, 156 + idx * 24), line, fontsize=11, fontname="china-ss")

    for num, title, summary, layout_kind, dur in SLIDE_PLAN:
        page = doc.new_page(width=595, height=842)
        page.insert_text((48, 64), f"第 {num} 页：{title}", fontsize=16, fontname="china-ss")
        page.insert_text((48, 104), f"版式：{layout_kind}    预计时长：{dur}", fontsize=11, fontname="china-ss")
        page.insert_text((48, 140), f"核心内容：{summary}", fontsize=11, fontname="china-ss")
    doc.save(str(handoff_pdf))
    doc.close()
    return {"plan_pptx": plan_pptx, "handoff_pdf": handoff_pdf}
