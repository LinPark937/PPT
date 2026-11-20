# ==============================================================================
# 【文件作用】`cellppt` 核心功能包的对外导出窗口 (cellppt/__init__.py)
# 把包里最常用的几个核心函数汇总暴露出去，方便外部脚本直接 `from cellppt import ...` 调用。
# ==============================================================================

from .nature2026_deck import generate_companion_files, generate_nature2026_pptx  # 导入 16 页示例 PPT 及配套审阅文档生成函数
from .paper_to_ppt import build_pptx_from_pdf                                    # 导入通用 PDF 论文转 PPTX 函数
from .playback_engine import prepare_hidden_copy, run_live_playback              # 导入元素预隐藏与桌面 PowerPoint 实时回放录屏函数
from .theme import THEMES, get_theme                                             # 导入配色主题字典与获取函数

__version__ = "2.0.0"  # 当前产品版本号

# 定义 `from cellppt import *` 时对外公开的符号列表
__all__ = [
    "THEMES",
    "get_theme",
    "generate_nature2026_pptx",
    "generate_companion_files",
    "build_pptx_from_pdf",
    "prepare_hidden_copy",
    "run_live_playback",
]
