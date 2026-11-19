"""Live PowerPoint Edit-Window Element-by-Element Playback & Window Recording Engine."""
from copy import copy
import ctypes
from ctypes import wintypes
from hashlib import sha256
import json
from pathlib import Path
import shutil
import subprocess
import threading
import time
from zipfile import ZipFile

import cv2
import imageio_ffmpeg
from lxml import etree
import numpy as np
from pptx import Presentation

NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
SHAPES = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}


def _p(name: str) -> str:
    return f"{{{NS}}}{name}"


def reading_order(slide_xml, slide=None):
    """Sort top-level shapes in visual reading order (top-to-bottom rows, left-to-right, z-index)."""
    geometry = {str(s.shape_id): s for s in slide.shapes} if slide is not None else {}
    tree = slide_xml.find(f"{_p('cSld')}/{_p('spTree')}")
    positioned = []
    for index, element in enumerate(tree):
        if not isinstance(element.tag, str):
            continue
        qname = etree.QName(element)
        if qname.namespace != NS or qname.localname not in SHAPES:
            continue
        info = element[0].find(_p("cNvPr"))
        if info is None or not info.get("id") or info.get("hidden") in {"1", "true"}:
            continue
        shape = geometry.get(info.get("id"))
        if shape is not None:
            left, top = shape.left, shape.top
        else:
            off = (
                element.find(f"{_p('spPr')}/{{{DRAWING}}}xfrm/{{{DRAWING}}}off")
                or element.find(f"{_p('grpSpPr')}/{{{DRAWING}}}xfrm/{{{DRAWING}}}off")
                or element.find(f"{_p('xfrm')}/{{{DRAWING}}}off")
            )
            left, top = (int(off.get("x", 0)), int(off.get("y", 0))) if off is not None else (0, 0)
        positioned.append((
            float("inf") if top is None else int(top),
            float("inf") if left is None else int(left),
            index,
            element,
        ))
    remaining = sorted(positioned, key=lambda item: item[:3])
    result = []
    while remaining:
        row_top = remaining[0][0]
        row, later = [], []
        for item in remaining:
            (row if item[0] <= row_top + 6 * 12700 else later).append(item)
        # Within the same visual row, sort by left, then z-order index so background bars precede labels
        result.extend(item[3] for item in sorted(row, key=lambda item: (item[1], item[2], item[0])))
        remaining = later
    return result


def prepare_hidden_copy(source_pptx: str | Path, hidden_pptx: str | Path) -> dict:
    """Create an isolated PPTX copy with all slide elements hidden (`hidden='1'`) in reading order."""
    source_pptx = Path(source_pptx).resolve()
    hidden_pptx = Path(hidden_pptx).resolve()
    hidden_pptx.parent.mkdir(parents=True, exist_ok=True)
    if hidden_pptx.exists():
        hidden_pptx.unlink()

    slides = Presentation(str(source_pptx)).slides
    parts = {str(slide.part.partname).lstrip("/"): (idx, slide) for idx, slide in enumerate(slides, 1)}
    targets = {}
    parser = etree.XMLParser(resolve_entities=False, no_network=True)

    with ZipFile(source_pptx, "r") as src_zip, ZipFile(hidden_pptx, "w") as dst_zip:
        for member in src_zip.infolist():
            payload = src_zip.read(member)
            if member.filename in parts:
                idx, slide = parts[member.filename]
                xml = etree.fromstring(payload, parser)
                ids = []
                for elem in reading_order(xml, slide):
                    info = elem[0].find(_p("cNvPr"))
                    if info is not None and info.get("id"):
                        ids.append(int(info.get("id")))
                        info.set("hidden", "1")
                if ids:
                    targets[str(slide.slide_id)] = ids
                payload = etree.tostring(xml, encoding="UTF-8", xml_declaration=True)
            dst_zip.writestr(copy(member), payload)
    return targets


class Win32WindowRecorder:
    """Captures the exact PowerPoint window (title bar, ribbon, thumbnails, canvas, status bar) into MP4."""

    def __init__(self, hwnd: int, output_mp4: str | Path, fps: int = 20, width: int = 1280, height: int = 704):
        self.hwnd = hwnd
        self.output_mp4 = Path(output_mp4).resolve()
        self.fps = fps
        self.width = width
        self.height = height
        self.stopping = threading.Event()
        self.frames_written = 0
        self.error = None
        self._latest_bgr = np.full((self.height, self.width, 3), 245, dtype=np.uint8)
        self._lock = threading.Lock()
        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
        self.proc = subprocess.Popen(
            [
                ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
                "-f", "rawvideo", "-pix_fmt", "bgr24",
                "-video_size", f"{width}x{height}",
                "-framerate", str(fps),
                "-i", "pipe:0", "-an",
                "-c:v", "libx264", "-preset", "ultrafast",
                "-crf", "20", "-pix_fmt", "yuv420p",
                "-movflags", "+faststart",
                str(self.output_mp4),
            ],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.cap_thread = threading.Thread(target=self._capture_loop, daemon=True)
        self.write_thread = threading.Thread(target=self._write_loop, daemon=True)
        self.cap_thread.start()
        self.write_thread.start()

    def _capture_hwnd_bgr(self):
        import win32gui
        import win32ui

        if not win32gui.IsWindow(self.hwnd):
            return None
        left, top, right, bot = win32gui.GetWindowRect(self.hwnd)
        w, h = right - left, bot - top
        if w <= 32 or h <= 32:
            return None
        hwnd_dc = win32gui.GetWindowDC(self.hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)
        # PW_RENDERFULLCONTENT = 2 captures DirectComposition / Office hardware-accelerated surfaces
        ctypes.windll.user32.PrintWindow(self.hwnd, save_dc.GetSafeHdc(), 2)
        bmp_info = bmp.GetInfo()
        bmp_bits = bmp.GetBitmapBits(True)
        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwnd_dc)

        img = np.frombuffer(bmp_bits, dtype=np.uint8).reshape((bmp_info["bmHeight"], bmp_info["bmWidth"], 4))
        bgr = img[:, :, :3]
        if bgr.shape[1] != self.width or bgr.shape[0] != self.height:
            bgr = cv2.resize(bgr, (self.width, self.height), interpolation=cv2.INTER_AREA)
        return bgr

    def _capture_loop(self):
        while not self.stopping.is_set():
            try:
                frame = self._capture_hwnd_bgr()
                if frame is not None and frame.std() > 2.0:
                    with self._lock:
                        self._latest_bgr = frame
            except Exception as exc:
                self.error = exc
            time.sleep(0.01)

    def _write_loop(self):
        interval = 1.0 / self.fps
        next_t = time.perf_counter()
        while not self.stopping.is_set():
            try:
                with self._lock:
                    payload = self._latest_bgr.tobytes()
                self.proc.stdin.write(payload)
                self.frames_written += 1
            except Exception as exc:
                self.error = exc
                break
            next_t += interval
            sleep_s = next_t - time.perf_counter()
            if sleep_s > 0:
                time.sleep(sleep_s)

    def stop(self):
        self.stopping.set()
        self.write_thread.join(timeout=5)
        self.cap_thread.join(timeout=3)
        try:
            self.proc.stdin.close()
        except Exception:
            pass
        self.proc.wait(timeout=15)
        return self.frames_written



def _find_ppt_hwnd(deck_name_stem: str) -> int:
    import win32gui

    found = []

    def _cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if deck_name_stem in title and ("PowerPoint" in title or ".pptx" in title or " - " in title):
                found.append(hwnd)

    win32gui.EnumWindows(_cb, None)
    return found[0] if found else 0


def run_live_playback(
    source_pptx: str | Path,
    output_mp4: str | Path | None = None,
    mode: str = "sequential",
    step_ms: int = 32,
    record: bool = True,
    keep_open: bool = False,
) -> dict:
    """Open PowerPoint in normal edit mode, reveal all elements live, and optionally record MP4.

    Parameters
    ----------
    source_pptx : Path to the complete editable PPTX.
    output_mp4  : Target MP4 path when `record=True`.
    mode        : 'sequential' (matches the 15s PDF reference video) or 'swarm' (10-cursor concurrent).
    step_ms     : Delay in milliseconds between element reveals in sequential mode (~32ms => ~15s for 16 slides).
    record      : Whether to record the live PowerPoint window into MP4.
    keep_open   : Keep the PowerPoint window open after completion.
    """
    import win32com.client
    import win32con
    import win32gui

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

    source_pptx = Path(source_pptx).resolve()
    work_dir = source_pptx.parent / "_playback_work"
    work_dir.mkdir(parents=True, exist_ok=True)
    # Name the working copy '成品ppt.pptx' so the PowerPoint title bar reads '成品ppt.pptx - PowerPoint'
    # exactly like the reference video!
    stage_dir = work_dir / f"run_{int(time.time())}"
    stage_dir.mkdir(parents=True, exist_ok=True)
    staged_pptx = stage_dir / "成品ppt.pptx"
    targets = prepare_hidden_copy(source_pptx, staged_pptx)

    if mode == "swarm":
        return _run_swarm_mode(source_pptx, staged_pptx, targets, output_mp4, record)

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = -1
    deck = app.Presentations.Open(str(staged_pptx), 0, 0, -1)
    win = deck.Windows.Item(1)
    win.WindowState = 3  # ppWindowMaximized
    win.ViewType = 9     # ppViewNormal = 9 (shows left thumbnail pane + main slide canvas!)
    try:
        win.SplitHorizontal = 15  # Expand left thumbnail strip (15% width) matching reference video
    except Exception:
        pass
    win.Activate()

    hwnd = 0
    for _ in range(30):
        hwnd = _find_ppt_hwnd("成品ppt")
        if hwnd:
            break
        time.sleep(0.1)

    if hwnd:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
    time.sleep(0.45)

    recorder = None
    if record and output_mp4 and hwnd:
        recorder = Win32WindowRecorder(hwnd, output_mp4, fps=30, width=1280, height=704)
        time.sleep(0.25)

    t0 = time.perf_counter()
    revealed_count = 0
    delay_s = max(0.005, step_ms / 1000.0)

    try:
        total_slides = deck.Slides.Count
        for s_idx in range(1, total_slides + 1):
            slide = deck.Slides.Item(s_idx)
            sid_key = str(slide.SlideID)
            shape_ids = targets.get(sid_key, [])
            if not shape_ids:
                continue
            win.View.GotoSlide(s_idx)
            shape_map = {}
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes.Item(i)
                shape_map[int(shp.Id)] = shp
            for target_id in shape_ids:
                shp = shape_map.get(int(target_id))
                if shp is not None:
                    shp.Visible = -1
                    revealed_count += 1
                    time.sleep(delay_s)
            # Brief pause at the end of each slide so the left thumbnail pane updates cleanly
            time.sleep(0.08)
        # Hold on the final slide briefly so the viewer sees the completed 16-slide deck
        time.sleep(0.5)
        deck.Save()
    finally:
        frames = recorder.stop() if recorder else 0
        if not keep_open:
            try:
                deck.Close()
            except Exception:
                pass

    elapsed = round(time.perf_counter() - t0, 2)
    return {
        "source_pptx": str(source_pptx),
        "staged_pptx": str(staged_pptx),
        "output_mp4": str(Path(output_mp4).resolve()) if (record and output_mp4) else None,
        "slides": len(targets),
        "revealed_elements": revealed_count,
        "frames_recorded": frames,
        "elapsed_seconds": elapsed,
        "mode": mode,
    }


def _run_swarm_mode(source_pptx: Path, staged_pptx: Path, targets: dict,
                    output_mp4: str | Path | None, record: bool) -> dict:
    """Invoke the multi-cursor concurrent swarm playback engine with optional window recording."""
    from .swarm_runner import execute_swarm_playback
    res = execute_swarm_playback(staged_pptx, targets, output_mp4=output_mp4, record=record)
    res["source_pptx"] = str(source_pptx)
    return res
