# ==============================================================================
# 【文件作用】v0.3 同源配套演讲逐字稿与演示者视图提词引擎 (cellppt/script_writer.py)
# 核心设计原则：“一份逐页规划蓝图，同时产出 PPT 画面 + PowerPoint 底部备注栏提词 + 独立 PDF/MD 演讲逐字稿”。
# 这样学生哪怕改了规划里的任何一页，PPT 和逐字稿页码永远 100% 同步，绝不会出现“图和讲稿对不上”的情况。
# ==============================================================================

from pathlib import Path
import pymupdf as fitz


def format_slide_script_text(slide_no: int, title: str, script: dict) -> str:
    # 【函数作用】把单页的逐字稿字典格式化为清晰易读的提词文本（用于写入 PPT 每一页底部的“备注栏”）
    sec = script.get("estimated_seconds", 60)
    t_in = script.get("transition_in", "")
    main = script.get("main_speech", "")
    fig_g = script.get("figure_guide", "")
    emph = script.get("key_emphasis", "")
    t_out = script.get("transition_out", "")

    lines = [f"【第 {slide_no:02d} 页：{title}】（建议用时：{sec} 秒）"]
    if t_in:
        lines.append(f"▶ [承上引入]：{t_in}")
    if fig_g:
        lines.append(f"▶ [指图引导]：{fig_g}")
    if main:
        lines.append(f"▶ [核心讲稿]：{main}")
    if emph:
        lines.append(f"★ [重点强调]：{emph}")
    if t_out:
        lines.append(f"▶ [启下翻页]：{t_out}")
    return "\n".join(lines)


def write_note_to_slide(slide, slide_no: int, title: str, script: dict | None):
    # 【函数作用】把格式化好的逐字稿直接写入 PowerPoint 当前页底部的“备注栏 (Notes Slide)”
    # 学生在答辩开启 PowerPoint“演示者视图”时，投影仪放 PPT，自己电脑屏幕直接显示这段逐字稿！
    if not script:
        return
    try:
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = format_slide_script_text(slide_no, title, script)
    except Exception:
        pass


def export_speech_scripts(plan: dict, output_pptx: Path) -> dict:
    # 【函数作用】根据同一份 plan 蓝图，同步导出独立的《*_演讲逐字稿.md》和《*_演讲逐字稿.pdf》提词手册
    meta = plan.get("meta", {})
    all_scripts = plan.get("all_scripts", [])
    refs = plan.get("references", [])

    md_path = output_pptx.with_name(f"{output_pptx.stem}_演讲逐字稿.md")
    pdf_path = output_pptx.with_name(f"{output_pptx.stem}_演讲逐字稿.pdf")

    total_sec = sum(item.get("estimated_seconds", 60) for item in all_scripts)
    total_min = round(total_sec / 60.0, 1)

    # --- 1. 生成 Markdown 格式逐字稿 ---
    md_lines = [
        f"# 《{meta.get('title', '学术课题汇报')}》配套演讲逐字稿（同源提词本）",
        "",
        f"- **汇报人**：{meta.get('authors', '')}  |  **总页数**：{len(all_scripts)} 页  |  **预计总时长**：约 {total_min} 分钟（{total_sec} 秒）",
        "- **使用提示**：本逐字稿与 PPT 画面及每页底部的 PowerPoint【备注栏】完全同源同步；答辩时开启 PowerPoint“演示者视图”即可自动看到每页提词。",
        "",
        "---",
        "",
    ]
    for item in all_scripts:
        s_no = item["slide_no"]
        s_title = item["title"]
        sec = item.get("estimated_seconds", 60)
        md_lines.append(f"## 第 {s_no:02d} 页：{s_title} （预计 {sec} 秒）")
        if item.get("layout"):
            md_lines.append(f"- **页面版式**：`{item['layout']}`")
        if item.get("transition_in"):
            md_lines.append(f"- **【承上引入】**：{item['transition_in']}")
        if item.get("figure_guide"):
            md_lines.append(f"- **【指图话术】**：**{item['figure_guide']}**")
        if item.get("main_speech"):
            md_lines.append(f"- **【开口逐字稿】**：\n  > {item['main_speech']}")
        if item.get("key_emphasis"):
            md_lines.append(f"- **【重音强调】**：*{item['key_emphasis']}*")
        if item.get("transition_out"):
            md_lines.append(f"- **【翻页衔接】**：{item['transition_out']}")
        md_lines.append("")

    if refs:
        md_lines.append("---")
        md_lines.append("## 附录：PPT 画面引用的核心参考文献索引")
        for r in refs:
            md_lines.append(f"- **{r['id']}** ({r.get('year', '')}): {r['citation']}")

    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    # --- 2. 生成 A4 打印版 PDF 逐字稿手册 ---
    doc = fitz.open()
    for idx, item in enumerate(all_scripts):
        if idx % 3 == 0:
            page = doc.new_page(width=595, height=842)
            y = 48
            if idx == 0:
                page.insert_text((42, y), f"《{meta.get('title', '')[:24]}》配套演讲逐字稿", fontsize=16, fontname="china-ss")
                y += 24
                page.insert_text((42, y), f"总页数：{len(all_scripts)} 页   |   预计总讲解时长：{total_min} 分钟 ({total_sec} 秒)", fontsize=10, fontname="china-ss")
                y += 24
        s_no = item["slide_no"]
        s_title = item["title"]
        sec = item.get("estimated_seconds", 60)
        block_text = (
            f"【第 {s_no:02d} 页：{s_title}】 (预计 {sec} 秒)\n"
            f"• 承上引入：{item.get('transition_in', '')}\n"
            f"• 指图引导：{item.get('figure_guide', '')}\n"
            f"• 开口讲稿：{item.get('main_speech', '')}\n"
            f"• 重点与翻页：{item.get('key_emphasis', '')} -> {item.get('transition_out', '')}"
        )
        rect = fitz.Rect(42, y, 553, y + 235)
        page.insert_textbox(rect, block_text, fontsize=10, fontname="china-ss")
        y += 245

    try:
        doc.save(str(pdf_path))
    except PermissionError:
        import time
        pdf_path = output_pptx.with_name(f"{output_pptx.stem}_演讲逐字稿_{int(time.time())}.pdf")
        doc.save(str(pdf_path))
    doc.close()
    return {"script_md": md_path, "script_pdf": pdf_path, "total_seconds": total_sec}
