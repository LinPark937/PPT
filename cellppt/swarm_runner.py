"""10-Cursor Concurrent Swarm Playback with Pinyin Bubbles & Window Recording."""
import time
from pathlib import Path
from pypinyin import lazy_pinyin
import win32com.client
import win32con
import win32gui

from .playback_engine import Win32WindowRecorder, _find_ppt_hwnd

COLORS = ["0284C7", "7C3AED", "16A34A", "EA580C", "DB2777",
          "0891B2", "4F46E5", "CA8A04", "0D9488", "BE123C"]


def _rgb(h: str) -> int:
    return int(h[:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:], 16) << 16)


def execute_swarm_playback(staged_pptx: Path, targets: dict,
                           output_mp4: str | Path | None = None, record: bool = True) -> dict:
    """Open PowerPoint and animate 10 concurrent cursors revealing graphics & text."""
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = -1
    deck = app.Presentations.Open(str(staged_pptx), 0, 0, -1)
    win = deck.Windows.Item(1)
    win.WindowState = 3
    win.ViewType = 9
    try:
        win.SplitHorizontal = 15
    except Exception:
        pass
    win.Activate()

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

    recorder = Win32WindowRecorder(hwnd, output_mp4, fps=30, width=1280, height=704) if (record and output_mp4 and hwnd) else None
    t0 = time.perf_counter()
    W, H = float(deck.PageSetup.SlideWidth), float(deck.PageSetup.SlideHeight)
    revealed = 0

    try:
        for s_idx in range(1, deck.Slides.Count + 1):
            slide = deck.Slides.Item(s_idx)
            ids = targets.get(str(slide.SlideID), [])
            if not ids:
                continue
            win.View.GotoSlide(s_idx)
            shape_map = {int(slide.Shapes.Item(i).Id): slide.Shapes.Item(i)
                         for i in range(1, slide.Shapes.Count + 1)}
            ordered = [shape_map[i] for i in ids if i in shape_map]

            # Build 6 lightweight colored cursor pointers on this slide
            n_cur = min(6, max(2, len(ordered)))
            pointers = []
            bubbles = []
            for ci in range(n_cur):
                col = COLORS[ci % len(COLORS)]
                f = slide.Shapes.BuildFreeform(0, 0, 0)
                for a, b in [(0, 17), (4, 13), (8, 21), (12, 19), (8, 11), (16, 11), (0, 0)]:
                    f.AddNodes(0, 0, a, b)
                c = f.ConvertToShape()
                c.Fill.ForeColor.RGB = _rgb(col)
                c.Line.ForeColor.RGB = _rgb("FFFFFF")
                c.Left, c.Top = W * 0.45 + ci * 18, H - 28
                pointers.append(c)

                tb = slide.Shapes.AddTextbox(1, c.Left, c.Top - 22, 76, 20)
                tb.Fill.Visible = -1
                tb.Fill.ForeColor.RGB = _rgb("FFFFFF")
                tb.Line.Visible = -1
                tb.Line.ForeColor.RGB = _rgb(col)
                tb.TextFrame.TextRange.Font.Size = 11
                tb.TextFrame.TextRange.Font.Color.RGB = _rgb(col)
                tb.Visible = 0
                bubbles.append(tb)

            # Distribute shapes across cursors and reveal concurrently in lockstep batches
            lanes = [ordered[i::n_cur] for i in range(n_cur)]
            max_len = max(len(ln) for ln in lanes)
            for step in range(max_len):
                for ci, ln in enumerate(lanes):
                    if step >= len(ln):
                        continue
                    shp = ln[step]
                    p, b = pointers[ci], bubbles[ci]
                    tx, ty = float(shp.Left), float(shp.Top)
                    p.Left, p.Top = max(0, min(W - 18, tx)), max(0, min(H - 22, ty))
                    has_txt = shp.HasTextFrame and shp.TextFrame.HasText
                    if has_txt:
                        txt = shp.TextFrame.TextRange.Text.strip()
                        py = "".join(lazy_pinyin(txt[:4]))[:10] if txt else "type"
                        b.Left, b.Top = max(0, min(W - 80, tx + 14)), max(0, min(H - 22, ty - 20))
                        b.TextFrame.TextRange.Text = py
                        b.Visible = -1
                    else:
                        b.Left, b.Top = max(0, min(W - 80, tx + 14)), max(0, min(H - 22, ty - 20))
                        b.TextFrame.TextRange.Text = "Ctrl+V"
                        b.Visible = -1
                    shp.Visible = -1
                    revealed += 1
                time.sleep(0.065)

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
