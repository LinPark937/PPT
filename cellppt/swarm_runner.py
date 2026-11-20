# ==============================================================================
# 【文件作用】“多光标并发打字 + 拼音气泡 + Ctrl+V 贴图特效引擎” (cellppt/swarm_runner.py)
# 当回放模式选了 `--mode swarm` 时，它会在每一页幻灯片上临时画出 6 个不同颜色的小鼠标指针和白色悬浮小气泡，
# 模拟好几个人同时在 PPT 上移动光标、敲拼音输入汉字、按 Ctrl+V 粘贴图片的并发制作过程，播完该页后再把临时光标删干净。
# ==============================================================================

import time                                                 # 导入时间与延时模块
from pathlib import Path                                    # 导入文件路径对象
from pypinyin import lazy_pinyin                            # 导入汉字转拼音库，用于在气泡里显示拼音字母
import win32com.client                                      # 导入 Windows Office COM 客户端接口
import win32con                                             # 导入 Windows 窗口常量
import win32gui                                             # 导入 Windows 窗口管理接口

from .playback_engine import Win32WindowRecorder, _find_ppt_hwnd  # 导入窗口录屏机与窗口句柄查找函数

# 10 个并发模拟鼠标指针的十六进制颜色表（天蓝、紫罗兰、翠绿、爱马仕橙、玫红、青色、靛蓝、金黄、蓝绿、酒红）
COLORS = [
    "0284C7", "7C3AED", "16A34A", "EA580C", "DB2777",
    "0891B2", "4F46E5", "CA8A04", "0D9488", "BE123C",
]


def _rgb(h: str) -> int:
    # 【函数作用】把 6 位十六进制颜色码（如 "0284C7"）转换成 Office COM 接口要求的 BGR 整数格式
    return int(h[:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:], 16) << 16)


def execute_swarm_playback(staged_pptx: Path, targets: dict,
                           output_mp4: str | Path | None = None, record: bool = True) -> dict:
    # 【函数作用】打开桌面 PowerPoint，在每一页上创建多个彩色光标箭头和拼音/Ctrl+V 气泡，分头并发点亮页面上的所有元素并同步录屏
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = -1                                              # 让 PowerPoint 窗口可见
    deck = app.Presentations.Open(str(staged_pptx), 0, 0, -1)     # 打开所有元素已预隐藏的 PPTX 副本
    win = deck.Windows.Item(1)
    win.WindowState = 3                                           # 最大化窗口
    win.ViewType = 9                                              # 切换到带左侧缩略图栏的普通编辑视图 (ppViewNormal = 9)
    try:
        win.SplitHorizontal = 15                                  # 左侧缩略图栏占 15% 宽度
    except Exception:
        pass
    win.Activate()

    # 查找并置顶 PowerPoint 窗口句柄
    hwnd = 0
    for _ in range(30):
        hwnd = _find_ppt_hwnd(staged_pptx.stem)
        if hwnd:
            break
        time.sleep(0.1)
    if hwnd:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
    time.sleep(0.4)

    # 如果开启了录屏，启动 30fps 的窗口录屏机
    recorder = Win32WindowRecorder(hwnd, output_mp4, fps=30, width=1280, height=704) if (record and output_mp4 and hwnd) else None
    t0 = time.perf_counter()
    W, H = float(deck.PageSetup.SlideWidth), float(deck.PageSetup.SlideHeight)  # 读取幻灯片画布宽高（单位：磅）
    revealed = 0

    try:
        # 遍历每一页幻灯片
        for s_idx in range(1, deck.Slides.Count + 1):
            slide = deck.Slides.Item(s_idx)
            ids = targets.get(str(slide.SlideID), [])
            if not ids:
                continue
            win.View.GotoSlide(s_idx)                             # 切换到当前页
            shape_map = {int(slide.Shapes.Item(i).Id): slide.Shapes.Item(i)
                         for i in range(1, slide.Shapes.Count + 1)}
            ordered = [shape_map[i] for i in ids if i in shape_map]

            # 【微调点】在当前页上临时创建 n_cur 个彩色鼠标指针形状（默认最多 6 个并发光标，想更多可把 6 改成 10）
            n_cur = min(6, max(2, len(ordered)))
            pointers = []
            bubbles = []
            for ci in range(n_cur):
                col = COLORS[ci % len(COLORS)]
                # 使用多边形折线顶点画出逼真的鼠标箭头形状
                f = slide.Shapes.BuildFreeform(0, 0, 0)
                for a, b in [(0, 17), (4, 13), (8, 21), (12, 19), (8, 11), (16, 11), (0, 0)]:
                    f.AddNodes(0, 0, a, b)
                c = f.ConvertToShape()
                c.Fill.ForeColor.RGB = _rgb(col)                  # 鼠标箭头填充对应颜色
                c.Line.ForeColor.RGB = _rgb("FFFFFF")             # 鼠标箭头加白色描边
                c.Left, c.Top = W * 0.45 + ci * 18, H - 28        # 初始停靠在画布底部中央
                pointers.append(c)

                # 为每个鼠标指针配套创建一个跟在旁边的白色悬浮小气泡框（用来显示拼音或 Ctrl+V）
                tb = slide.Shapes.AddTextbox(1, c.Left, c.Top - 22, 76, 20)
                tb.Fill.Visible = -1
                tb.Fill.ForeColor.RGB = _rgb("FFFFFF")
                tb.Line.Visible = -1
                tb.Line.ForeColor.RGB = _rgb(col)
                tb.TextFrame.TextRange.Font.Size = 11
                tb.TextFrame.TextRange.Font.Color.RGB = _rgb(col)
                tb.Visible = 0
                bubbles.append(tb)

            # 把当前页所有待显现的元素均匀分配给 n_cur 个光标（分头干活）
            lanes = [ordered[i::n_cur] for i in range(n_cur)]
            max_len = max(len(ln) for ln in lanes)
            for step in range(max_len):
                for ci, ln in enumerate(lanes):
                    if step >= len(ln):
                        continue
                    shp = ln[step]
                    p, b = pointers[ci], bubbles[ci]
                    tx, ty = float(shp.Left), float(shp.Top)
                    # 把第 ci 号鼠标箭头瞬移到目标元素的左上角
                    p.Left, p.Top = max(0, min(W - 18, tx)), max(0, min(H - 22, ty))
                    has_txt = shp.HasTextFrame and shp.TextFrame.HasText
                    if has_txt:
                        # 如果目标元素里带文字，把前 4 个字转成拼音显示在悬浮气泡里（如 "erweiwangz"）
                        txt = shp.TextFrame.TextRange.Text.strip()
                        py = "".join(lazy_pinyin(txt[:4]))[:10] if txt else "type"
                        b.Left, b.Top = max(0, min(W - 80, tx + 14)), max(0, min(H - 22, ty - 20))
                        b.TextFrame.TextRange.Text = py
                        b.Visible = -1
                    else:
                        # 如果目标元素是纯图形或图片，在悬浮气泡里显示 "Ctrl+V" 粘贴手势
                        b.Left, b.Top = max(0, min(W - 80, tx + 14)), max(0, min(H - 22, ty - 20))
                        b.TextFrame.TextRange.Text = "Ctrl+V"
                        b.Visible = -1
                    # 将该隐藏元素点亮（设为可见）
                    shp.Visible = -1
                    revealed += 1
                # 【微调点】每批次并发点亮后的停顿时间（0.065 秒，想让并发打字慢一点可调成 0.12）
                time.sleep(0.065)

            # 当前页所有元素点亮完毕后，把临时创建的鼠标箭头和拼音气泡全部删除干净，绝不留在成品 PPT 里
            for h_shp in pointers + bubbles:
                try:
                    h_shp.Delete()
                except Exception:
                    pass
            time.sleep(0.08)
        time.sleep(0.4)
        deck.Save()
    finally:
        frames = recorder.stop() if recorder else 0
        try:
            deck.Close()
        except Exception:
            pass

    return {
        "staged_pptx": str(staged_pptx),
        "output_mp4": str(Path(output_mp4).resolve()) if (record and output_mp4) else None,
        "slides": len(targets),
        "revealed_elements": revealed,
        "frames_recorded": frames,
        "elapsed_seconds": round(time.perf_counter() - t0, 2),
        "mode": "swarm",
    }
