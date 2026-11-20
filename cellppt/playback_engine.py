# ==============================================================================
# 【文件作用】“桌面 PowerPoint 逐元素动态生长回放 + 窗口自动录屏引擎” (cellppt/playback_engine.py)
# 这是复现教程 PDF 内嵌视频特效的核心文件。工作原理分三步：
#   1. prepare_hidden_copy : 先把做好的 PPTX 复制一份，钻进底层 XML 把每一页上的所有形状按“从上往下、从左往右”排好队，并全部加上隐藏标记 (hidden="1")。
#   2. run_live_playback   : 启动桌面版 Microsoft PowerPoint 打开这份“全隐藏副本”（刚打开时每页看起来都是空白的），然后每隔二十几毫秒把一个元素的开关打开 (Visible = -1)，实现逐元素丝滑生长。
#   3. Win32WindowRecorder : 在生长的同时，针对 PowerPoint 窗口抓图并通过 FFmpeg 实时合成 1280x704 的 MP4 视频。
# ==============================================================================

from copy import copy           # 导入浅拷贝工具，用于复制 ZIP 压缩包条目信息
import ctypes                   # 导入 Windows C 语言动态链接库调用工具（调 PrintWindow 和 DPI 感知）
from ctypes import wintypes     # 导入 Windows 基础数据类型
from hashlib import sha256      # 导入 SHA256 哈希计算工具
import json                     # 导入 JSON 读写模块
from pathlib import Path        # 导入文件路径处理模块
import shutil                   # 导入文件复制与清理模块
import subprocess               # 导入子进程管理模块，用于启动后台 FFmpeg 编码器
import threading                # 导入多线程模块，用于后台抓图线程与固定帧率写视频线程
import time                     # 导入高精度时钟与延时模块
from zipfile import ZipFile     # 导入 ZIP 解压/打包模块（因为 .pptx 本质上是一个 ZIP 压缩包）

import cv2                      # 导入 OpenCV 图像处理库，用于缩放抓屏画面
import imageio_ffmpeg           # 导入自带的静态 FFmpeg 可执行程序获取器
from lxml import etree          # 导入 XML 解析与修改库
import numpy as np              # 导入数值数组库，用于图像像素矩阵处理
from pptx import Presentation   # 导入 python-pptx 库

# PowerPoint 底层 PresentationML 与 DrawingML 的 XML 命名空间常量
NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
# 需要被隐藏并逐个显现的 5 类顶层幻灯片形状标签：
#   sp=普通形状/文本框, pic=图片, graphicFrame=图表/表格, grpSp=组合形状, cxnSp=连接线/箭头
SHAPES = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}


def _p(name: str) -> str:
    # 【函数作用】给 XML 节点名拼上 PowerPoint PresentationML 的命名空间前缀
    return f"{{{NS}}}{name}"


def reading_order(slide_xml, slide=None):
    # 【函数作用】给一页幻灯片上的所有元素按“人类视觉阅读顺序”排队：
    # 先按高度（top）把同一水平线附近的元素归为同一行，同一行内再按从左到右（left）和图层先后顺序（index）排序。
    # 这样回放时就会严格按照“先出顶部标题栏 -> 再出要点 -> 再从左到右出卡片 -> 最后出底栏”的自然顺序生长。
    geometry = {str(s.shape_id): s for s in slide.shapes} if slide is not None else {}
    tree = slide_xml.find(f"{_p('cSld')}/{_p('spTree')}")  # 找到幻灯片 XML 里的形状树 <p:spTree>
    positioned = []
    # 遍历该页上的每一个 XML 形状节点
    for index, element in enumerate(tree):
        if not isinstance(element.tag, str):
            continue
        qname = etree.QName(element)
        # 如果不是上述 5 类可视形状（比如是母版占位信息），则跳过不处理
        if qname.namespace != NS or qname.localname not in SHAPES:
            continue
        info = element[0].find(_p("cNvPr"))  # 找到存放形状 ID 和隐藏状态的 <p:cNvPr> 节点
        if info is None or not info.get("id") or info.get("hidden") in {"1", "true"}:
            continue
        shape = geometry.get(info.get("id"))
        # 读取该形状在页面上的左上角横坐标 left 和纵坐标 top（单位：EMU）
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
    # 第一步：先按纵坐标 top 从上到下粗排
    remaining = sorted(positioned, key=lambda item: item[:3])
    result = []
    while remaining:
        row_top = remaining[0][0]  # 当前这一行的基准顶部高度
        row, later = [], []
        for item in remaining:
            # 【微调点】行高容差：6 * 12700 EMU（即 6 磅高度差以内的元素，都视为处在同一视觉水平行）
            # 如果你想让上下稍微错开一点的两个框也被归为同一行从左往右出，可以把 6 调大到 12
            (row if item[0] <= row_top + 6 * 12700 else later).append(item)
        # 在同一水平行内，按横坐标 left(item[1]) 从左到右排；如果横坐标相同（比如底栏背景色块和底栏文字重叠），
        # 按原始图层顺序 index(item[2]) 排，确保背景色块永远先于上面的文字出现！
        result.extend(item[3] for item in sorted(row, key=lambda item: (item[1], item[2], item[0])))
        remaining = later
    return result


def prepare_hidden_copy(source_pptx: str | Path, hidden_pptx: str | Path) -> dict:
    # 【函数作用】生成“预隐藏回放副本”：
    # 在不修改原始《成品ppt.pptx》的前提下，复制出一份临时 PPTX，把里面每一页的所有元素都加上 hidden="1" 隐藏属性，
    # 并返回每一页按阅读顺序排好队的元素 ID 列表字典：{ "幻灯片SlideID": [形状ID1, 形状ID2, ...] }
    source_pptx = Path(source_pptx).resolve()
    hidden_pptx = Path(hidden_pptx).resolve()
    hidden_pptx.parent.mkdir(parents=True, exist_ok=True)
    if hidden_pptx.exists():
        hidden_pptx.unlink()  # 如果旧的临时副本存在，先删掉

    slides = Presentation(str(source_pptx)).slides
    # 建立 ZIP 包内路径（如 ppt/slides/slide1.xml）到 Slide 对象的映射
    parts = {str(slide.part.partname).lstrip("/"): (idx, slide) for idx, slide in enumerate(slides, 1)}
    targets = {}
    parser = etree.XMLParser(resolve_entities=False, no_network=True)

    # 打开原始 PPTX 压缩包读取，同时把修改了 hidden="1" 的 XML 写进新的临时 PPTX 压缩包
    with ZipFile(source_pptx, "r") as src_zip, ZipFile(hidden_pptx, "w") as dst_zip:
        for member in src_zip.infolist():
            payload = src_zip.read(member)
            if member.filename in parts:
                idx, slide = parts[member.filename]
                xml = etree.fromstring(payload, parser)
                ids = []
                # 按上面算出的视觉阅读顺序遍历每个形状，把它的 hidden 属性设为 "1"（在 PowerPoint 编辑窗口中隐藏）
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
    # 【类作用】“PowerPoint 专属窗口录屏机”
    # 只针对打开的那个 PowerPoint 窗口抓图（包含标题栏、顶部功能区、左侧缩略图栏、中央画布和底部状态栏），
    # 不录桌面其它无关区域，并通过独立写帧时钟线程 + FFmpeg 编码成流畅的 MP4 视频。

    def __init__(self, hwnd: int, output_mp4: str | Path, fps: int = 30, width: int = 1280, height: int = 704):
        # 【函数作用】初始化录屏器并启动后台抓图线程与 FFmpeg 视频编码进程
        # 【微调点】：
        #   fps    : 录制视频的帧率（默认 30 帧/秒，与参考视频完全一致）
        #   width  : 录制视频的画面宽度（默认 1280 像素；想要 1080P 全高清可改为 1920）
        #   height : 录制视频的画面高度（默认 704 像素；想要 1080P 全高清可改为 1080）
        self.hwnd = hwnd
        self.output_mp4 = Path(output_mp4).resolve()
        self.fps = fps
        self.width = width
        self.height = height
        self.stopping = threading.Event()                                              # 停止录屏信号旗
        self.frames_written = 0                                                        # 已写入的视频总帧数计数器
        self.error = None                                                              # 记录录制过程中的异常
        self._latest_bgr = np.full((self.height, self.width, 3), 245, dtype=np.uint8)  # 最新抓取到的窗口画面缓冲区
        self._lock = threading.Lock()                                                  # 保护画面缓冲区的线程锁
        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()                                       # 获取内置 FFmpeg 路径
        # 启动 FFmpeg 子进程，通过标准输入管道 (pipe:0) 接收原始 BGR 画面并实时编码为 H.264 MP4
        self.proc = subprocess.Popen(
            [
                ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
                "-f", "rawvideo", "-pix_fmt", "bgr24",
                "-video_size", f"{width}x{height}",
                "-framerate", str(fps),
                "-i", "pipe:0", "-an",
                "-c:v", "libx264", "-preset", "ultrafast",
                "-crf", "20", "-pix_fmt", "yuv420p",  # 【微调点】-crf 20 控制画质（数字越小画质越高，18~23 最合适）
                "-movflags", "+faststart",
                str(self.output_mp4),
            ],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        # 启动两个独立后台线程：一个负责不停抓图，另一个负责按固定时钟节拍把图写进视频
        self.cap_thread = threading.Thread(target=self._capture_loop, daemon=True)
        self.write_thread = threading.Thread(target=self._write_loop, daemon=True)
        self.cap_thread.start()
        self.write_thread.start()

    def _capture_hwnd_bgr(self):
        # 【函数作用】调用 Windows 底层的 PrintWindow 接口，把当前 PowerPoint 窗口画面抓成一张 BGR 图像矩阵
        import win32gui
        import win32ui

        if not win32gui.IsWindow(self.hwnd):
            return None
        left, top, right, bot = win32gui.GetWindowRect(self.hwnd)  # 获取 PowerPoint 窗口在屏幕上的实际宽高
        w, h = right - left, bot - top
        if w <= 32 or h <= 32:
            return None
        hwnd_dc = win32gui.GetWindowDC(self.hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)
        # 最后一个参数传 2 (PW_RENDERFULLCONTENT)，专门用来抓取 Office 硬件加速渲染窗口的完整画面
        ctypes.windll.user32.PrintWindow(self.hwnd, save_dc.GetSafeHdc(), 2)
        bmp_info = bmp.GetInfo()
        bmp_bits = bmp.GetBitmapBits(True)
        # 及时释放 Windows GDI 句柄，防止内存泄漏
        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwnd_dc)

        # 把抓取到的字节流转成 OpenCV 的 BGR 三通道图像矩阵，并缩放到目标视频分辨率 (width x height)
        img = np.frombuffer(bmp_bits, dtype=np.uint8).reshape((bmp_info["bmHeight"], bmp_info["bmWidth"], 4))
        bgr = img[:, :, :3]
        if bgr.shape[1] != self.width or bgr.shape[0] != self.height:
            bgr = cv2.resize(bgr, (self.width, self.height), interpolation=cv2.INTER_AREA)
        return bgr

    def _capture_loop(self):
        # 【函数作用】后台持续抓图线程：不停抓取 PowerPoint 窗口最新画面，更新到 `_latest_bgr` 缓冲区
        while not self.stopping.is_set():
            try:
                frame = self._capture_hwnd_bgr()
                # frame.std() > 2.0 用于过滤掉偶尔抓到的纯黑/纯白无效帧
                if frame is not None and frame.std() > 2.0:
                    with self._lock:
                        self._latest_bgr = frame
            except Exception as exc:
                self.error = exc
            time.sleep(0.01)

    def _write_loop(self):
        # 【函数作用】后台固定时钟写帧线程：严格按照 1/fps 秒的间隔把当前画面写入 FFmpeg，确保录出来的视频时长与真实耗时 1:1 完全同步
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
        # 【函数作用】结束录屏：通知两个后台线程停止，关闭管道并等待 FFmpeg 完成 MP4 文件封尾
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
    # 【函数作用】遍历 Windows 当前所有可见窗口，找到窗口标题里包含指定 PPT 文件名的那个 PowerPoint 窗口句柄 (hwnd)
    import win32gui

    found = []

    def _cb(hwnd, _):
        # 【函数作用】EnumWindows 的枚举回调函数：检查窗口是否可见且标题匹配
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
    step_ms: int = 25,
    record: bool = True,
    keep_open: bool = False,
) -> dict:
    # 【函数作用】驱动桌面 PowerPoint 进行“逐元素生长”的核心总控函数
    # 【核心微调参数】：
    #   mode      : "sequential" 为视频原版顺序生长；"swarm" 为多光标并发打字+贴图特效
    #   step_ms   : 每个元素跳出来的时间间隔（默认 25 毫秒，16 页 375 个元素刚好播 15 秒；想慢一点看清细节可调成 80~150）
    #   record    : True=边播边录 MP4 视频；False=只弹窗演示不录视频
    #   keep_open : True=播完后保留 PowerPoint 窗口不关闭；False=播完自动关闭窗口
    import win32com.client
    import win32con
    import win32gui

    # 开启 Windows 高分屏 DPI 感知，防止在 125%/150% 缩放的屏幕上抓屏画面被裁切
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

    source_pptx = Path(source_pptx).resolve()
    work_dir = source_pptx.parent / "_playback_work"
    work_dir.mkdir(parents=True, exist_ok=True)
    # 故意把临时回放副本命名为“成品ppt.pptx”，这样 PowerPoint 顶部标题栏显示的就是“成品ppt - Microsoft PowerPoint”，与原视频一模一样
    stage_dir = work_dir / f"run_{int(time.time())}"
    stage_dir.mkdir(parents=True, exist_ok=True)
    staged_pptx = stage_dir / "成品ppt.pptx"
    # 第一步：生成所有元素都被隐藏的隔离回放副本，并拿到每页的元素出现顺序表 targets
    targets = prepare_hidden_copy(source_pptx, staged_pptx)

    # 如果用户选的是 swarm 多光标并发模式，转交给 _run_swarm_mode 执行
    if mode == "swarm":
        return _run_swarm_mode(source_pptx, staged_pptx, targets, output_mp4, record)

    # 第二步：通过 COM 启动桌面版 PowerPoint 并打开隐藏副本
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = -1                                          # 让 PowerPoint 主程序窗口可见
    deck = app.Presentations.Open(str(staged_pptx), 0, 0, -1) # 以正常有窗口模式打开演示文稿
    win = deck.Windows.Item(1)
    try:
        win.WindowState = 3                                   # 3 = ppWindowMaximized（将 PowerPoint 窗口最大化）
    except Exception:
        pass
    try:
        # 【关键设置】ViewType = 9 代表 ppViewNormal（带左侧缩略图栏 + 中央主画布的标准三窗格编辑视图）
        win.ViewType = 9
        # 【微调点】SplitHorizontal = 15 表示左侧缩略图栏占窗口宽度的 15%
        win.SplitHorizontal = 15
    except Exception:
        pass
    try:
        win.Activate()
    except Exception:
        pass

    # 查找刚刚打开的 PowerPoint 窗口句柄并将其置顶激活
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
    time.sleep(0.45)  # 等待窗口最大化动画彻底稳定后再开始录屏

    # 第三步：如果开启了录屏，启动窗口录屏机（30fps，1280x704 分辨率）
    recorder = None
    if record and output_mp4 and hwnd:
        recorder = Win32WindowRecorder(hwnd, output_mp4, fps=30, width=1280, height=704)
        time.sleep(0.25)

    t0 = time.perf_counter()
    revealed_count = 0
    delay_s = max(0.005, step_ms / 1000.0)  # 将毫秒间隔转换为秒

    # 第四步：从第 1 页到最后一页，按阅读顺序把隐藏的元素一个个点亮 (shp.Visible = -1)
    try:
        total_slides = deck.Slides.Count
        for s_idx in range(1, total_slides + 1):
            slide = deck.Slides.Item(s_idx)
            sid_key = str(slide.SlideID)
            shape_ids = targets.get(sid_key, [])
            if not shape_ids:
                continue
            win.View.GotoSlide(s_idx)  # 切换到第 s_idx 页
            # 建立当前页所有形状 ID 到 COM Shape 对象的字典，加快查找速度
            shape_map = {}
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes.Item(i)
                shape_map[int(shp.Id)] = shp
            # 按视觉阅读顺序依次点亮每一个元素
            for target_id in shape_ids:
                shp = shape_map.get(int(target_id))
                if shp is not None:
                    shp.Visible = -1       # -1 在 Office COM 中代表 msoTrue（设为可见）
                    revealed_count += 1
                    time.sleep(delay_s)    # 每点亮一个元素暂停 delay_s 秒
            # 【微调点】每页所有元素全部出来后，短暂停留 0.08 秒，让左侧缩略图栏刷新出该页完整画面
            time.sleep(0.08)
        # 【微调点】全部 16 页播完后，在最后一页多停留 0.5 秒再结束录屏
        time.sleep(0.5)
        deck.Save()  # 保存全部元素已恢复可见的副本
    finally:
        # 第五步：停止录屏并根据 keep_open 决定是否关闭 PowerPoint 窗口
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
    # 【函数作用】当用户选择 mode="swarm" 时，调用 swarm_runner.py 里的多光标并发打字+贴图特效引擎
    from .swarm_runner import execute_swarm_playback
    res = execute_swarm_playback(staged_pptx, targets, output_mp4=output_mp4, record=record)
    res["source_pptx"] = str(source_pptx)
    return res
