# ==============================================================================
# 【文件作用】v0.2 通用 PDF/Word 论文转高颜值百变版式 PPT 生成器 (cellppt/paper_to_ppt.py)
# 相比 v0.1 的三大核心升级：
#   1. 自动调用 is_decorative_image 过滤封面校名/院徽等无用装饰图，并把真科研图与图注锁死
#   2. 调用 smart_planner 清洗页眉噪音、提炼“四字锚点：高密度结论”
#   3. 调用 layout_templates 里的 8 种对标视频同款百变版式（双栏对比、三列递进、三区流程、架构拆解、重点难点、四卡网格、原生图表+实景截图、问题对策+时间轴），页页排版变换自如且风格高度统一！
# ==============================================================================

import json                                                   # 导入 JSON 序列化模块（用于保存可读可改的规划蓝图）
import re                                                     # 导入正则表达式模块
from pathlib import Path                                      # 导入路径处理对象
import pymupdf as fitz                                        # 导入 PyMuPDF 文档解析库
from pptx import Presentation                                 # 导入 python-pptx 主演示文稿类
from pptx.util import Inches                                  # 导入英寸单位

from .builder import text_box                                 # 导入文本框绘制函数
from .layout_templates import LAYOUT_REGISTRY                 # 导入 v0.2 的 8 大百变版式注册表
from .layouts_basic import agenda, cover, thanks              # 导入封面、目录与致谢页模板
from .smart_planner import (                                  # 导入 v0.2 智能过滤与语义规划器
    build_v02_slide_plan,
    clean_academic_text,
    extract_cover_metadata,
    is_decorative_image,
)
from .theme import SLIDE_H, SLIDE_W, get_theme                # 导入全局宽高与主题配色获取函数


def _ensure_pdf(input_path: Path, work_dir: Path) -> Path:
    # 【函数作用】如果用户传入的是 Word 文档 (.doc 或 .docx)，自动调用本机 Word COM 接口无损转换为 .pdf 再处理
    if input_path.suffix.lower() in (".doc", ".docx"):
        import win32com.client
        work_dir.mkdir(parents=True, exist_ok=True)
        out_pdf = work_dir / f"{input_path.stem}.pdf"
        if not out_pdf.exists():
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(input_path), ReadOnly=True)
            # 17 = wdFormatPDF (Word 导出为标准 PDF 格式)
            doc.SaveAs(str(out_pdf), FileFormat=17)
            doc.Close(False)
        return out_pdf
    return input_path


def extract_pdf_assets(pdf_path: Path, asset_dir: Path, max_figures: int = 18) -> dict:
    # 【函数作用】打开 PDF/Word 文档，剔除封面校徽/装饰图，提取有效科研图及其图注(caption)，并清洗每页正文
    asset_dir.mkdir(parents=True, exist_ok=True)
    real_pdf = _ensure_pdf(Path(pdf_path).resolve(), asset_dir)
    doc = fitz.open(str(real_pdf))
    total_pages = len(doc)

    # 先读取所有页的原始文本，供提取封面元数据与跨页图注匹配使用
    raw_pages = [(pno, page.get_text("text")) for pno, page in enumerate(doc, 1)]
    meta = extract_cover_metadata(raw_pages, Path(pdf_path).stem)

    figures = []
    skipped_logos = []
    seen_xrefs = set()
    for pno, page in enumerate(doc, 1):
        page_txt = raw_pages[pno - 1][1]
        next_txt = raw_pages[pno][1] if pno < total_pages else ""
        for img_idx, info in enumerate(page.get_images(full=True)):
            xref = info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            base = doc.extract_image(xref)
            if not base:
                continue
            w, h = base.get("width", 0), base.get("height", 0)
            # 【v0.2 升级点 1】调用智能装饰图过滤器，把封面上的“山东师范大学”等校名/院徽图片直接拦截剔除！
            if is_decorative_image(pno, total_pages, w, h, page_txt):
                skipped_logos.append({"page": pno, "w": w, "h": h, "reason": "decorative_or_cover_logo"})
                continue
            ext = base.get("ext", "png")
            out_fn = asset_dir / f"fig_p{pno:02d}_img{img_idx:02d}.{ext}"
            out_fn.write_bytes(base["image"])
            # 【v0.2 升级点 2】在当前页和下一页文本中自动抓取该图片的“图一/图二/Fig.”图注，实现图文锁死！
            combined = page_txt + "\n" + next_txt[:200]
            cap_match = re.findall(r"(图\s*[一二三四五六七八九十\d]+[，,：:\s][^\n]+)", combined)
            caption = cap_match[0].strip() if cap_match else f"第 {pno} 页核心科研配图"
            figures.append({
                "page": pno,
                "path": out_fn,
                "w": w,
                "h": h,
                "caption": caption,
                "context": clean_academic_text(combined)[:400],
            })
            if len(figures) >= max_figures:
                break
        if len(figures) >= max_figures:
            break

    # 清洗每一页正文（剥离页眉页脚与目录点线）
    cleaned_pages = []
    for pno, raw_t in raw_pages:
        ct = clean_academic_text(raw_t)
        if ct:
            cleaned_pages.append((pno, ct))
    doc.close()
    return {
        "title": meta["title"],
        "meta": meta,
        "asset_dir": str(asset_dir),
        "figures": figures,
        "skipped_logos": skipped_logos,
        "pages": cleaned_pages,
    }


def build_pptx_from_pdf(pdf_path: str | Path, output_pptx: str | Path,
                        theme_name: str = "nature_purple",
                        user_images_dir: str | Path | None = None) -> Path:
    # 【函数作用】v0.3 主构建入口：
    # 1. 提取文档图文 + 合并客户自传图片 + 过滤无用校名 Logo
    # 2. 调用 v0.3 智能规划器（含参考文献自动抓图/合成机理图 + 8大百变版式 + 同源逐字稿）
    # 3. 渲染输出高颜值 PPTX（每页底部备注栏自动写入演示者提词逐字稿）+ 独立导出 MD/PDF 演讲逐字稿手册
    from .ai_client import ensure_config_template
    from .image_hunter import collect_user_uploaded_images
    from .script_writer import export_speech_scripts, write_note_to_slide

    ensure_config_template()
    pdf_path = Path(pdf_path).resolve()
    output_pptx = Path(output_pptx).resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    asset_dir = output_pptx.parent / f"{output_pptx.stem}_assets"

    # 第一步：提取文档图文并合并客户自己上传的图片文件夹 (--user-images)
    info = extract_pdf_assets(pdf_path, asset_dir)
    user_figs = collect_user_uploaded_images(user_images_dir)
    if user_figs:
        info["figures"] = user_figs + info["figures"]

    # 第二步：调用 v0.3 智能规划器生成图文锁死、参考文献配图补齐、带同源逐字稿的幻灯片蓝图
    plan = build_v02_slide_plan(info)
    plan_json_path = output_pptx.with_name(f"{output_pptx.stem}_plan_v03.json")
    plan_json_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    theme = get_theme(theme_name)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # 1. 绘制第 1 页：封面页 + 写入第 1 页底部备注栏逐字稿
    meta = plan["meta"]
    s_cover = prs.slides.add_slide(blank)
    cover(s_cover, theme, meta, {"byline": meta.get("byline"), "title_size": 44})
    write_note_to_slide(s_cover, 1, f"封面：{meta.get('title', '')}", plan.get("cover_script"))

    # 2. 绘制第 2 页：汇报目录页 + 写入第 2 页底部备注栏逐字稿
    s_agenda = prs.slides.add_slide(blank)
    agenda(s_agenda, theme, meta, {"items": plan["agenda"], "active": 0})
    write_note_to_slide(s_agenda, 2, "汇报目录", plan.get("agenda_script"))

    # 3. 循环绘制正文页：调度 8 大百变版式 + 将每一页的同源逐字稿写入 PowerPoint 备注栏
    for idx, s_data in enumerate(plan["slides"], 3):
        slide = prs.slides.add_slide(blank)
        layout_name = s_data.get("layout", "dual_compare")
        renderer = LAYOUT_REGISTRY.get(layout_name, LAYOUT_REGISTRY["dual_compare"])
        renderer(slide, theme, idx, s_data)
        write_note_to_slide(slide, idx, s_data.get("title", ""), s_data.get("speaker_script"))

    # 4. 绘制最后一页：封底致谢页 + 写入致谢页备注栏逐字稿
    s_thanks = prs.slides.add_slide(blank)
    thanks(s_thanks, theme, meta, {})
    if meta.get("thanks_info"):
        text_box(s_thanks, Inches(1.6), Inches(5.45), SLIDE_W - Inches(3.2), Inches(1.65),
                 meta["thanks_info"], theme, size=16, color=theme.white, anchor="top", line_spacing=1.3)
    write_note_to_slide(s_thanks, len(plan["slides"]) + 3, "致谢与交流讨论", plan.get("thanks_script"))

    # 保存 PPTX（带防 PowerPoint 占用自动改名保护）
    try:
        prs.save(str(output_pptx))
    except PermissionError:
        import time
        output_pptx = output_pptx.with_name(f"{output_pptx.stem}_{int(time.time())}.pptx")
        prs.save(str(output_pptx))

    # 5. 同步导出独立排版好的《*_演讲逐字稿.md》与《*_演讲逐字稿.pdf》提词手册
    export_speech_scripts(plan, output_pptx)
    return output_pptx
