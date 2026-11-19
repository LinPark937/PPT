# CellPPT Studio Pro v2.0 — 科研 PPT 自动生成与 PowerPoint 实时逐元素输出系统

本项目完整解析并本地化复现了 `cell_ppt`（`yrui-cmd/cell_ppt`）及其教程 PDF 内嵌演示视频的全套效果，移除全部远程付费 API Key 依赖，打包为开箱即用的本地化产品。

---

## 一、原项目与 PDF 内嵌视频全面逆向解析

### 1. PDF 内嵌视频（RichMedia Stream `1.mp4`）揭秘
- **物理流信息**：位于教程 PDF 第 1 页 `RichMedia` 注释（xref 79），分辨率 `1280×704`，帧率 `30 fps`，总帧数 `458 帧`（时长 `15.27 秒`）。
- **演示文献**：*A two-dimensional reticular triaxial molecularly woven fabric*（Wang 等，*Nature* 2026，DOI: `10.1038/s41586-026-11218-8`）。
- **核心视觉机制（“逐元素生长”的真相）**：
  - 视频展示的**并非在空白 PPT 中实时调用插入接口逐个新建形状**（COM 逐个新建复杂形状会导致版式重排与严重闪烁），而是**“预构建完整可编辑 PPTX + 预隐藏 XML 节点 + PowerPoint 编辑窗口实时逐元素显现”**：
    1. **预构建阶段**：先使用 `python-pptx` 与 DrawingML 生成全部 16 页、375+ 个原生可编辑矢量形状、富文本框、化学键与科研图表的《成品ppt.pptx》。
    2. **预隐藏阶段（`prepare_hidden_copy`）**：解析每页 XML（`p:cSld/p:spTree`），按视觉阅读顺序（`top` 按行分桶容差 `6 × 12700` EMU，同行按 `left` 从左到右及 Z-order 排序）将所有顶层形状（`sp`, `pic`, `graphicFrame`, `grpSp`, `cxnSp`）标记为 `p:cNvPr hidden="1"`，生成隔离回放副本。
    3. **PowerPoint 桌面编辑窗口逐元素输出（`run_live_playback`）**：通过 COM 接口启动桌面 Microsoft PowerPoint 正常编辑视图（最大化窗口，保留完整标题栏、Ribbon 功能区、左侧缩略图栏、中央画布与底部状态栏），从第 1 页至第 16 页依次跳转，并按视觉阅读顺序将隐藏形状设回 `shape.Visible = -1`，同时通过窗口级抓屏编码器录制成 MP4。

---

## 二、产品架构与目录说明

```text
e:\AGP_WORK\科研ppt_g\
├── cellppt_studio.py            # 统一命令行与产品主入口 (demo / replay / from-pdf)
├── run_demo.bat                 # Windows 一键双击运行脚本
├── cellppt/
│   ├── __init__.py              # 包导出接口
│   ├── theme.py                 # 4 套顶级期刊配色主题 (nature_purple, cell_blue, science_red, forest_green)
│   ├── builder.py               # 原生可编辑 DrawingML 构建器 (富文本、上下标、渐变、连接符、无损插图)
│   ├── layouts_basic.py         # 封面、目录、章节过渡页、致谢页与统一编号标题栏
│   ├── blocks.py                # 列表项、高亮横幅、PowerPoint 原生可编辑图表 (柱状/折线/圆环/散点)
│   ├── nature2026_data.py       # Nature 2026 演示文稿元数据与目录结构
│   ├── slides_1_to_6.py         # 第 1~6 页高精度原生构建逻辑
│   ├── slides_7_to_11.py        # 第 7~11 页高精度原生构建逻辑
│   ├── slides_12_to_16.py       # 第 12~16 页高精度原生构建逻辑
│   ├── nature2026_deck.py       # 16 页《成品ppt.pptx》、《逐页安排.pptx》、《问答与原始资料.pdf》组装器
│   ├── paper_to_ppt.py          # 通用科研论文 PDF 自动解析提取图文转 PPTX 引擎
│   ├── playback_engine.py       # 视觉阅读顺序预隐藏 + PowerPoint 实时逐元素显现 + 1:1 窗口录屏引擎
│   └── swarm_runner.py          # 升级版多光标并发打字 + 拼音气泡 + Ctrl+V 贴图特效引擎
├── assets/
│   └── nature2026_figures/      # 38 张高清科研子图素材
└── output/
    ├── 成品ppt.pptx             # 生成的 16 页全可编辑成品 PPTX (375 个原生形状)
    ├── 逐页安排.pptx            # 16 页逐页规划审阅稿 PPTX
    ├── 问答与原始资料.pdf        # 交接审阅 PDF
    └── ppt.mp4                  # 桌面 PowerPoint 窗口逐元素动态输出实录视频
```

---

## 三、快速使用指南

### 1. 一键复现 PDF 视频中的 16 页 PPTX + 桌面实时逐元素输出与录屏
```powershell
.\.venv\Scripts\python.exe cellppt_studio.py demo --record yes --mode sequential --step-ms 25
```
或直接双击根目录下的 `run_demo.bat`。

### 2. 对任意现有 PPTX 执行 PowerPoint 逐元素生长输出与录屏
- **视频原版顺序生长模式 (`sequential`)**：
  ```powershell
  .\.venv\Scripts\python.exe cellppt_studio.py replay --pptx output\成品ppt.pptx --mode sequential --record yes
  ```
- **多光标并发 + 拼音气泡模式 (`swarm`)**：
  ```powershell
  .\.venv\Scripts\python.exe cellppt_studio.py replay --pptx output\成品ppt.pptx --mode swarm --record yes
  ```

### 3. 从任意科研论文 PDF 自动生成可编辑汇报 PPTX
```powershell
.\.venv\Scripts\python.exe cellppt_studio.py from-pdf --pdf your_paper.pdf --output output\my_report.pptx --theme cell_blue --play yes --record yes
```
