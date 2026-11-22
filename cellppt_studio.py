# ==============================================================================
# 【文件作用】整个项目的总入口与命令行遥控器 (cellppt_studio.py)
# 支持三大子命令：
#   1. demo     : 一键生成 16 页 Nature 示例 PPT + 审阅文件 + 弹窗逐元素回放并录屏
#   2. replay   : 对任意已有的 .pptx 文件执行桌面 PowerPoint 逐元素生长动画与录屏
#   3. from-pdf : 传入任意一篇科研论文 PDF，自动提取图文生成 PPTX 并可选回放录屏
# ==============================================================================

import argparse          # 导入命令行参数解析模块
import json              # 导入 JSON 格式化输出模块
from pathlib import Path # 导入跨平台文件路径处理模块
import sys               # 导入系统退出与运行环境模块

from cellppt import (
    THEMES,                    # 所有可选的配色主题字典
    build_pptx_from_pdf,       # 从任意 PDF 论文生成 PPTX 的函数
    generate_companion_files,  # 生成《逐页安排.pptx》与《问答与原始资料.pdf》的函数
    generate_nature2026_pptx,  # 生成 16 页 Nature 2026《成品ppt.pptx》的函数
    run_live_playback,         # 启动桌面 PowerPoint 逐元素动态回放并录屏的函数
)


def cmd_demo(args):
    # 【函数作用】执行 `demo` 子命令：生成 16 页示例 PPT、配套审阅文件，并按需启动 PowerPoint 现场逐元素输出与录屏
    out_dir = Path(args.output_dir).resolve()  # 解析输出文件夹的绝对路径（默认是 output/）
    out_dir.mkdir(parents=True, exist_ok=True) # 如果输出文件夹不存在则自动创建
    pptx_path = out_dir / "成品ppt.pptx"       # 最终交付的 16 页可编辑 PPTX 文件路径
    mp4_path = out_dir / "ppt.mp4"             # 录屏输出的 MP4 视频文件路径

    print(f"[1/3] 正在生成 16 页全可编辑《成品ppt.pptx》 (主题: {args.theme})...")
    # 调用 16 页生成器；如果原文件正被 PowerPoint 打开占用，它会返回自动改名后的实际保存路径
    actual_pptx = generate_nature2026_pptx(pptx_path, theme_name=args.theme)

    print("[2/3] 正在生成审阅交接文件《逐页安排.pptx》与《问答与原始资料.pdf》...")
    # 生成配套的逐页规划 PPTX 和问答交接 PDF
    comp = generate_companion_files(out_dir, theme_name=args.theme)

    # 组装执行结果字典
    result = {
        "status": "completed",
        "pptx": str(actual_pptx),
        "plan_pptx": str(comp["plan_pptx"]),
        "handoff_pdf": str(comp["handoff_pdf"]),
        "video": None,
    }
    # 如果命令行参数 --play 为 yes，则启动桌面 PowerPoint 逐元素生长回放
    if args.play == "yes":
        print(f"[3/3] 正在启动 PowerPoint 桌面编辑窗口执行逐元素输出 (mode={args.mode}, record={args.record})...")
        play_info = run_live_playback(
            source_pptx=actual_pptx,                              # 传入刚生成好的成品 PPTX
            output_mp4=mp4_path if args.record == "yes" else None, # 如果 --record yes 则录制 MP4
            mode=args.mode,                                       # 回放模式：sequential(顺序) 或 swarm(多光标并发)
            step_ms=args.step_ms,                                 # 每个元素弹出的间隔毫秒数（微调动画快慢改这个参数）
            record=(args.record == "yes"),                        # 是否开启窗口录屏
            keep_open=args.keep_open,                             # 回放结束后是否保持 PowerPoint 窗口不关闭
        )
        result.update(play_info)                 # 把回放统计信息（耗时、帧数、元素数）合并进结果
        result["video"] = play_info.get("output_mp4")
    # 在终端打印 JSON 格式的完整交付结果
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_replay(args):
    # 【函数作用】执行 `replay` 子命令：拿任意一个现成的 .pptx 文件，在桌面 PowerPoint 中逐元素动态播放并录屏
    pptx_path = Path(args.pptx).resolve()  # 获取用户传入的 PPTX 绝对路径
    if not pptx_path.is_file():            # 检查文件是否存在
        raise FileNotFoundError(f"找不到 PPTX 文件: {pptx_path}")
    # 如果没指定输出 mp4 路径，默认保存在同目录下的 ppt.mp4
    mp4_path = Path(args.output_mp4).resolve() if args.output_mp4 else pptx_path.with_name("ppt.mp4")
    # 调用桌面 PowerPoint 实时回放与录屏引擎
    play_info = run_live_playback(
        source_pptx=pptx_path,
        output_mp4=mp4_path if args.record == "yes" else None,
        mode=args.mode,
        step_ms=args.step_ms,
        record=(args.record == "yes"),
        keep_open=args.keep_open,
    )
    print(json.dumps(play_info, ensure_ascii=False, indent=2))
    return 0


def cmd_from_pdf(args):
    # 【函数作用】执行 `from-pdf` 子命令：读取任意科研论文 PDF/Word + 可选客户自传图片文件夹，
    # 自动提取图文、从参考文献补齐科研配图、生成可编辑 PPTX（内置演示者视图备注栏逐字稿）及独立演讲逐字稿文档
    pdf_path = Path(args.pdf).resolve()    # 输入论文 PDF/Word 的绝对路径
    out_pptx = Path(args.output).resolve() # 输出 PPTX 的绝对路径
    actual_pptx = build_pptx_from_pdf(
        pdf_path,
        out_pptx,
        theme_name=args.theme,
        user_images_dir=args.user_images,
    )
    result = {
        "status": "completed",
        "pdf": str(pdf_path),
        "pptx": str(actual_pptx),
        "plan_json": str(actual_pptx.with_name(f"{actual_pptx.stem}_plan_v03.json")),
        "speech_script_md": str(actual_pptx.with_name(f"{actual_pptx.stem}_演讲逐字稿.md")),
        "speech_script_pdf": str(actual_pptx.with_name(f"{actual_pptx.stem}_演讲逐字稿.pdf")),
    }
    # 如果指定了 --play yes，同样启动 PowerPoint 逐元素回放
    if args.play == "yes":
        mp4_path = actual_pptx.with_suffix(".mp4")
        play_info = run_live_playback(
            source_pptx=actual_pptx,
            output_mp4=mp4_path if args.record == "yes" else None,
            mode=args.mode,
            step_ms=args.step_ms,
            record=(args.record == "yes"),
        )
        result.update(play_info)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def main():
    # 【函数作用】定义命令行参数规则（微调默认参数值如默认主题、默认间隔毫秒数 step-ms 都在这里改）
    parser = argparse.ArgumentParser(description="CellPPT Studio Pro — 科研 PPT 自动生成与 PowerPoint 实时逐元素输出系统")
    sub = parser.add_subparsers(dest="command", required=True)

    # --- 子命令 1: demo ---
    p_demo = sub.add_parser("demo", help="一键复现 Nature 2026 16页成品 PPTX + 逐元素实时回放与 MP4 录屏")
    p_demo.add_argument("--output-dir", default="output", help="输出目录 (默认: output)")
    p_demo.add_argument("--theme", choices=list(THEMES.keys()), default="nature_purple", help="配色主题皮肤")
    p_demo.add_argument("--play", choices=["yes", "no"], default="yes", help="是否弹窗启动 PowerPoint 逐元素回放")
    p_demo.add_argument("--record", choices=["yes", "no"], default="yes", help="回放时是否同步录制 MP4 视频")
    p_demo.add_argument("--mode", choices=["sequential", "swarm"], default="sequential", help="sequential=顺序生长, swarm=多光标并发")
    p_demo.add_argument("--step-ms", type=int, default=25, help="每个元素显现的间隔毫秒数 (默认25ms，16页总耗时约15秒)")
    p_demo.add_argument("--keep-open", action="store_true", help="输出完成后保持 PowerPoint 窗口打开不关闭")
    p_demo.set_defaults(func=cmd_demo)

    # --- 子命令 2: replay ---
    p_rep = sub.add_parser("replay", help="对任意已有 PPTX 执行 PowerPoint 桌面逐元素动态输出与录屏")
    p_rep.add_argument("--pptx", required=True, help="输入 PPTX 文件路径")
    p_rep.add_argument("--output-mp4", default=None, help="输出 MP4 视频路径")
    p_rep.add_argument("--record", choices=["yes", "no"], default="yes", help="是否录屏保存为 MP4")
    p_rep.add_argument("--mode", choices=["sequential", "swarm"], default="sequential", help="回放动画模式")
    p_rep.add_argument("--step-ms", type=int, default=25, help="每元素间隔毫秒数 (想慢放可改成 80~150)")
    p_rep.add_argument("--keep-open", action="store_true", help="回放结束后保持 PowerPoint 打开")
    p_rep.set_defaults(func=cmd_replay)

    # --- 子命令 3: from-pdf ---
    p_pdf = sub.add_parser("from-pdf", help="从任意论文 PDF/Word 提取图文、从参考文献补图并生成 PPTX + 同源演讲逐字稿")
    p_pdf.add_argument("--pdf", required=True, help="输入论文 PDF 或 Word (.doc/.docx) 路径")
    p_pdf.add_argument("--output", default="output/paper_presentation.pptx", help="输出 PPTX 路径")
    p_pdf.add_argument("--user-images", default=None, help="可选：客户自己上传的图片文件夹路径")
    p_pdf.add_argument("--theme", choices=list(THEMES.keys()), default="nature_purple", help="配色主题皮肤")
    p_pdf.add_argument("--play", choices=["yes", "no"], default="no", help="生成后是否立即弹窗逐元素回放")
    p_pdf.add_argument("--record", choices=["yes", "no"], default="no", help="回放时是否录屏")
    p_pdf.add_argument("--mode", choices=["sequential", "swarm"], default="sequential", help="回放动画模式")
    p_pdf.add_argument("--step-ms", type=int, default=35, help="每元素间隔毫秒数")
    p_pdf.set_defaults(func=cmd_from_pdf)

    args = parser.parse_args()  # 解析终端传入的参数
    return args.func(args)      # 调用对应的处理函数


if __name__ == "__main__":
    raise SystemExit(main())
