"""CellPPT Studio Pro — Unified CLI & Product Entry Point.

Usage examples:
  # 1. Reproduce the 16-slide Nature 2026 deck + companion files + live PowerPoint playback & MP4 recording
  python cellppt_studio.py demo --record yes --mode sequential

  # 2. Replay any existing PPTX in desktop PowerPoint (sequential or 10-cursor swarm mode)
  python cellppt_studio.py replay --pptx output/成品ppt.pptx --mode swarm --record yes

  # 3. Generate an editable academic PPTX from any research paper PDF
  python cellppt_studio.py from-pdf --pdf paper.pdf --output output/paper_deck.pptx --theme cell_blue
"""
import argparse
import json
from pathlib import Path
import sys

from cellppt import (
    THEMES,
    build_pptx_from_pdf,
    generate_companion_files,
    generate_nature2026_pptx,
    run_live_playback,
)


def cmd_demo(args):
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_path = out_dir / "成品ppt.pptx"
    mp4_path = out_dir / "ppt.mp4"

    print(f"[1/3] 正在生成 16 页全可编辑《成品ppt.pptx》 (主题: {args.theme})...")
    generate_nature2026_pptx(pptx_path, theme_name=args.theme)

    print("[2/3] 正在生成审阅交接文件《逐页安排.pptx》与《问答与原始资料.pdf》...")
    comp = generate_companion_files(out_dir, theme_name=args.theme)

    result = {
        "status": "completed",
        "pptx": str(pptx_path),
        "plan_pptx": str(comp["plan_pptx"]),
        "handoff_pdf": str(comp["handoff_pdf"]),
        "video": None,
    }
    if args.play == "yes":
        print(f"[3/3] 正在启动 PowerPoint 桌面编辑窗口执行逐元素输出 (mode={args.mode}, record={args.record})...")
        play_info = run_live_playback(
            source_pptx=pptx_path,
            output_mp4=mp4_path if args.record == "yes" else None,
            mode=args.mode,
            step_ms=args.step_ms,
            record=(args.record == "yes"),
            keep_open=args.keep_open,
        )
        result.update(play_info)
        result["video"] = play_info.get("output_mp4")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_replay(args):
    pptx_path = Path(args.pptx).resolve()
    if not pptx_path.is_file():
        raise FileNotFoundError(f"找不到 PPTX 文件: {pptx_path}")
    mp4_path = Path(args.output_mp4).resolve() if args.output_mp4 else pptx_path.with_name("ppt.mp4")
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
    pdf_path = Path(args.pdf).resolve()
    out_pptx = Path(args.output).resolve()
    build_pptx_from_pdf(pdf_path, out_pptx, theme_name=args.theme)
    result = {"status": "completed", "pdf": str(pdf_path), "pptx": str(out_pptx)}
    if args.play == "yes":
        mp4_path = out_pptx.with_suffix(".mp4")
        play_info = run_live_playback(
            source_pptx=out_pptx,
            output_mp4=mp4_path if args.record == "yes" else None,
            mode=args.mode,
            step_ms=args.step_ms,
            record=(args.record == "yes"),
        )
        result.update(play_info)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def main():
    parser = argparse.ArgumentParser(description="CellPPT Studio Pro — 科研 PPT 自动生成与 PowerPoint 实时逐元素输出系统")
    sub = parser.add_subparsers(dest="command", required=True)

    p_demo = sub.add_parser("demo", help="一键复现 Nature 2026 16页成品 PPTX + 逐元素实时回放与 MP4 录屏")
    p_demo.add_argument("--output-dir", default="output", help="输出目录 (默认: output)")
    p_demo.add_argument("--theme", choices=list(THEMES.keys()), default="nature_purple")
    p_demo.add_argument("--play", choices=["yes", "no"], default="yes")
    p_demo.add_argument("--record", choices=["yes", "no"], default="yes")
    p_demo.add_argument("--mode", choices=["sequential", "swarm"], default="sequential")
    p_demo.add_argument("--step-ms", type=int, default=28, help="每元素显现间隔毫秒数 (默认28ms，总长约15秒)")
    p_demo.add_argument("--keep-open", action="store_true", help="输出完成后保持 PowerPoint 窗口打开")
    p_demo.set_defaults(func=cmd_demo)

    p_rep = sub.add_parser("replay", help="对任意已有 PPTX 执行 PowerPoint 桌面逐元素动态输出与录屏")
    p_rep.add_argument("--pptx", required=True, help="输入 PPTX 路径")
    p_rep.add_argument("--output-mp4", default=None, help="输出 MP4 路径")
    p_rep.add_argument("--record", choices=["yes", "no"], default="yes")
    p_rep.add_argument("--mode", choices=["sequential", "swarm"], default="sequential")
    p_rep.add_argument("--step-ms", type=int, default=28)
    p_rep.add_argument("--keep-open", action="store_true")
    p_rep.set_defaults(func=cmd_replay)

    p_pdf = sub.add_parser("from-pdf", help="从任意论文 PDF 提取图文并生成全可编辑学术汇报 PPTX")
    p_pdf.add_argument("--pdf", required=True, help="输入论文 PDF 路径")
    p_pdf.add_argument("--output", default="output/paper_presentation.pptx", help="输出 PPTX 路径")
    p_pdf.add_argument("--theme", choices=list(THEMES.keys()), default="nature_purple")
    p_pdf.add_argument("--play", choices=["yes", "no"], default="no")
    p_pdf.add_argument("--record", choices=["yes", "no"], default="no")
    p_pdf.add_argument("--mode", choices=["sequential", "swarm"], default="sequential")
    p_pdf.add_argument("--step-ms", type=int, default=35)
    p_pdf.set_defaults(func=cmd_from_pdf)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
