"""Universal PDF-to-PPTX and JSON-to-PPTX academic presentation generator."""
import json
import re
from pathlib import Path
import pymupdf as fitz
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from .blocks import banner, bullets
from .builder import picture_fit, shape_box, text_box
from .layouts_basic import agenda, cover, header_bar, thanks
from .theme import SLIDE_H, SLIDE_W, get_theme


def extract_pdf_assets(pdf_path: Path, asset_dir: Path, max_figures: int = 18) -> dict:
    """Extract metadata, structured sections, and embedded figures from any research PDF."""
    asset_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    first_text = doc[0].get_text("text") if len(doc) else ""
    lines = [ln.strip() for ln in first_text.splitlines() if ln.strip()]
    title = doc.metadata.get("title") or (lines[0] if lines else pdf_path.stem)
    if len(title) > 90:
        title = title[:88] + "…"

    figures = []
    seen_xrefs = set()
    for pno, page in enumerate(doc, 1):
        for img_idx, info in enumerate(page.get_images(full=True)):
            xref = info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            base = doc.extract_image(xref)
            if not base or base.get("width", 0) < 160 or base.get("height", 0) < 120:
                continue
            ext = base.get("ext", "png")
            out_fn = asset_dir / f"pdf_p{pno:02d}_img{img_idx:02d}.{ext}"
            out_fn.write_bytes(base["image"])
            figures.append({"page": pno, "path": out_fn, "w": base["width"], "h": base["height"]})
            if len(figures) >= max_figures:
                break
        if len(figures) >= max_figures:
            break

    pages_text = []
    for pno, page in enumerate(doc, 1):
        txt = re.sub(r"\s+", " ", page.get_text("text")).strip()
        if txt:
            pages_text.append((pno, txt))
    doc.close()
    return {"title": title, "figures": figures, "pages": pages_text}


def build_pptx_from_pdf(pdf_path: str | Path, output_pptx: str | Path,
                        theme_name: str = "nature_purple") -> Path:
    pdf_path = Path(pdf_path).resolve()
    output_pptx = Path(output_pptx).resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    asset_dir = output_pptx.parent / f"{pdf_path.stem}_assets"
    info = extract_pdf_assets(pdf_path, asset_dir)
    theme = get_theme(theme_name)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    meta = {
        "authors": "学术文献解析",
        "venue": "Research Report",
        "year": "2026",
        "title": info["title"],
        "short_title": info["title"][:28],
        "tag": "文献汇报 | 自动生成",
        "tag_short": "文献汇报",
        "cover_info": f"来源文件: {pdf_path.name}",
    }
    cover(prs.slides.add_slide(blank), theme, meta, {})

    sec_names = ["研究背景与问题", "核心方法与设计", "关键实验与证据", "拓展分析与讨论", "结论与展望"]
    agenda(prs.slides.add_slide(blank), theme, meta, {"items": sec_names, "active": 0})

    figs = info["figures"]
    pages = info["pages"]
    n_content = max(4, min(10, len(pages) * 2))
    for idx in range(n_content):
        slide = prs.slides.add_slide(blank)
        slide_no = idx + 3
        pno, raw = pages[idx % len(pages)] if pages else (1, "核心研究内容概述")
        sentences = [s.strip() for s in re.split(r"[。！？.!?；;]", raw) if len(s.strip()) > 12]
        head_title = sentences[0][:26] if sentences else f"核心进展与证据 ({idx + 1})"
        header_bar(slide, theme, head_title, slide_no)

        top_pts = sentences[1:3] if len(sentences) >= 3 else ["核心结论与实验数据梳理", "结合图表证据验证关键机制"]
        bullets(slide, theme, [s[:46] for s in top_pts], Inches(0.35), Inches(1.02), Inches(12.4), size=20)

        # Two-column card layout with extracted figures or structured summary
        for col in range(2):
            cx = Inches(0.32) + col * Inches(6.42)
            cw = Inches(6.22)
            shape_box(slide, cx, Inches(2.1), cw, Inches(4.25), fill=theme.white, line=theme.line)
            sub_h = f"关键模块 {col + 1} (P{pno})"
            text_box(slide, cx, Inches(2.1), cw, Inches(0.58), sub_h, theme,
                     size=20, bold=True, align="center", fill=theme.soft)
            fig_idx = (idx * 2 + col)
            if fig_idx < len(figs):
                picture_fit(slide, figs[fig_idx]["path"], cx + Inches(0.25), Inches(2.8), cw - Inches(0.5), Inches(2.65))
            body_sent = sentences[min(len(sentences) - 1, 2 + col)][:65] if sentences else "实验与理论数据相互印证。"
            text_box(slide, cx + Inches(0.2), Inches(5.55), cw - Inches(0.4), Inches(0.68),
                     body_sent, theme, size=16, bold=True, align="center", fill=theme.softer)

        banner(slide, theme, f"证据来源：{pdf_path.name} 第 {pno} 页",
               Inches(0.32), Inches(6.55), SLIDE_W - Inches(0.64), Inches(0.58), style="bar", size=19)

    thanks(prs.slides.add_slide(blank), theme, meta, {})
    prs.save(str(output_pptx))
    return output_pptx
