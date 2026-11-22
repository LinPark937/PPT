# ==============================================================================
# 【文件作用】v0.3 同源演讲逐字稿模板库 (cellppt/speech_templates.py)
# 与 PPT 每一页的版式、图表、参考文献引用严格一一对应，包含承上引入、指图引导、开口讲稿、重点强调与翻页衔接。
# ==============================================================================

WIFI_SPEECH_SCRIPTS = {
    1: {
        "estimated_seconds": 35,
        "transition_in": "各位评审老师好！我是计算机技术（感知智能方向）的硕士研究生王树坤，指导教师是肖航老师。",
        "figure_guide": "请各位老师看封面，我本次硕士学位论文开题报告的题目是《基于域自适应的 Wi-Fi 手势识别泛化方法研究》。",
        "main_speech": "本研究面向智能家居与自然人机交互场景，重点探讨如何利用商用 Wi-Fi 设备的信道状态信息（CSI），解决手势识别模型在跨用户、跨位置与跨环境下的性能衰减问题。",
        "key_emphasis": "核心目标是在 Widar3.0 物理特征基线上引入域对抗迁移学习，实现从实验室走向真实场景的跨域泛化。",
        "transition_out": "下面我将从以下六个部分向各位老师作详细汇报。",
    },
    2: {
        "estimated_seconds": 30,
        "transition_in": "首先请看汇报目录。",
        "figure_guide": "本次汇报共分为六个章节：01 选题依据与跨域痛点、02 国内外现状与演进、03 研究方案与核心架构、04 研究重点难点与创新、05 已有研究基础与实验、06 问题对策与研究计划。",
        "main_speech": "我们将遵循“发现跨域痛点 → 梳理技术演进 → 提出 BVP+DANN 混合方案 → 展示已完成基线实验与系统原型 → 规划后续攻关路线”的逻辑层层展开。",
        "key_emphasis": "首先进入第一部分：选题依据与跨域痛点。",
        "transition_out": "请看第 3 页。",
    },
    3: {
        "estimated_seconds": 75,
        "transition_in": "为什么我们要研究基于 Wi-Fi 的跨域手势识别？",
        "figure_guide": "请各位老师先看左侧卡片（基于参考文献 Ref.[1-5] 整理的菲涅尔区多径感知原理图）。",
        "main_speech": "传统手势识别主要有三类：基于摄像头的视觉方案受光照限制且有隐私顾虑；基于可穿戴传感器的方案需要用户随身佩戴设备，侵入感强。而 Wi-Fi 无线感知利用室内无处不在的射频信号，当人手在发射端 Tx 与接收端 Rx 构成的菲涅尔区内挥动时，会引起动态反射径的多普勒频移（DFS）。清华大学团队提出的 Widar3.0（Ref.[5]）通过提取体坐标速度剖面（BVP），剥离了静态背景。然而请看右侧卡片的四条胶囊痛点：一旦训练集（源域）与测试集（目标域）在用户体型、所处位置或房间环境上发生变化，就会产生严重的“域漂移（Domain Shift）”，导致识别准确率断崖式下降。",
        "key_emphasis": "因此，在复现 Widar3.0 基线的基础上，引入域自适应（Domain Adaptation）消除跨域分布偏移，是推动 Wi-Fi 感知走向实用化的关键一步。",
        "transition_out": "那么国内外学者目前是如何应对这一挑战的呢？请看第 4 页。",
    },
    4: {
        "estimated_seconds": 70,
        "transition_in": "纵观国内外研究现状，Wi-Fi 手势识别大致经历了三个阶段的演进。",
        "figure_guide": "请看页面上的三列递进卡片：演进一、演进二和演进三。",
        "main_speech": "第一阶段是物理模型驱动特征，以 Widar3.0 的 BVP 特征为代表，物理可解释性强，但手工建模极度依赖先验知识；第二阶段是端到端深度学习，学者利用 CNN、RNN、3D-CNN 直接从 CSI 学习时空特征，同分布下精度极高，但严重依赖海量标注数据，换个环境性能就急剧崩塌；第三阶段则是近年来兴起的跨域泛化研究，主要包括以 GAN 和域判别器为代表的域自适应方法，以及多模态注意力机制。但现有域自适应多集中在图像与语音领域，针对 Wi-Fi 特有的 BVP 物理特征如何做深度对抗适配，研究尚不深入。",
        "key_emphasis": "正因如此，本课题的研究定位非常明确：将物理模型稳健的 BVP 特征与深度学习强大的域对抗迁移（DANN）深度融合。",
        "transition_out": "接下来汇报第二部分：我们的具体研究方案与技术路线，请看第 5 页。",
    },
    5: {
        "estimated_seconds": 80,
        "transition_in": "本课题的研究方案第一步，是对经典基线 Widar3.0 系统的全链路复现。",
        "figure_guide": "请各位老师看右侧的《图一：Widar3.0 基线系统信号处理全流程》以及中间的四步流程图。",
        "main_speech": "在硬件端，我们基于商用 Intel 5300 网卡和 Linux CSI Tool 采集细粒度幅值与相位数据（即图一顶部的 Data Acquisition）；随后进入中间橙色虚线框的 BVP Generation 阶段：通过相位校准、时频分析提取多普勒频移剖面（DFS Profile），再利用压缩感知（Compressed Sensing）估计出与环境无关的体坐标速度剖面 BVP；最后进入底部紫色虚线框的 Gesture Recognition 阶段，将归一化后的 BVP 序列送入 CNN 提取空间分布特征、结合 GRU 建模时序动作轨迹，最终输出手势类别。",
        "key_emphasis": "打通这条从原始 CSI 到 BVP 再到 CNN-GRU 的完整流水线，为我们定量分析跨域性能衰减并接入域对抗模块奠定了坚实基线。",
        "transition_out": "在基线之上，我们如何引入域自适应核心创新？请看第 6 页。",
    },
    6: {
        "estimated_seconds": 85,
        "transition_in": "第 6 页展示了本课题最核心的算法创新：基于 BVP 与 DANN 的域对抗泛化架构。",
        "figure_guide": "请各位老师重点对比左侧《图二》中的蓝色基线分支（Widar3.0 baseline）与右侧橙色增强分支（Enhanced method）。",
        "main_speech": "在传统蓝色基线中，分类器学到的特征往往混杂着特定用户和特定房间的干扰印记。为此，我们在右侧橙色分支中设计了三层协同机制（请对照右侧三张卡片）：第一，特征提取器以物理 BVP 时空序列为输入，提取高层语义特征；第二，也是最关键的一步，我们在特征提取器与域判别器之间接入了梯度反转层（Gradient Reversal Layer, GRL），它在前向传播时不改变数据，但在反向传播时将梯度乘以负系数（-λ）；第三，主任务手势分类器努力分清“是什么手势”，辅助任务域判别器努力分清“来自哪个环境”，而 GRL 迫使特征提取器去欺骗域判别器。",
        "key_emphasis": "当这双分支博弈达到纳什均衡时，网络学到的就是既抹除了环境/用户差异、又保留了手势判别力的“域不变特征”。",
        "transition_out": "围绕这一核心架构，我们需要攻克哪些研究重点与难点？请看第 7 页。",
    },
}


WIFI_SPEECH_SCRIPTS_PART2 = {
    7: {
        "estimated_seconds": 75,
        "transition_in": "第 7 页梳理了本课题的两大研究重点与三项关键技术难点。",
        "figure_guide": "请看上半部分的两张重点卡片与下半部分夹在中间的三列难点。",
        "main_speech": "研究重点一是设计适配 BVP 时空特性的 DANN 混合网络结构；重点二是通过消融实验探索 GRL 在哪一层级对齐效果最优。而在实际攻关中，我们面临三大难点：难点一是高维复杂场景下的特征解耦，多径效应与用户体型非线性耦合，剥离难度大；难点二是对抗训练的纳什均衡收敛稳定性，生成与判别博弈极易出现梯度消失或负迁移；难点三是商用网卡低信噪比处理，实际环境动态噪声会污染 BVP 纯净度。",
        "key_emphasis": "针对这三大难点，我们将在后续通过“物理滤波提纯 + 预训练渐进对抗 + 动态损失权重”三位一体加以解决。",
        "transition_out": "接下来汇报本课题的特色与创新之处，请看第 8 页。",
    },
    8: {
        "estimated_seconds": 70,
        "transition_in": "第 8 页总结了本研究在方法、实验与系统三个维度的特色与创新。",
        "figure_guide": "请各位老师先看左起第一张卡片（基于参考文献 Ref.[7,9] 合成的域对齐前后特征空间对比图）。",
        "main_speech": "如左一图所示，对齐前源域（紫点）与目标域（红块）分布偏离，原有分类超平面失效；引入 DANN 对抗对齐后，两域特征紧致重合。以此为核心，本课题具备三大创新：第一是方法创新，构建“物理 BVP 特征 + 对抗迁移”混合驱动范式；第二是实验创新，针对现有文献跨域定义模糊的问题，建立涵盖跨用户、跨位置、跨房间的多维细粒度评估基准；第三是系统特色与软著产出，仅需目标域无标签数据即可完成适配，支撑“一次训练、处处可用”的低成本部署。",
        "key_emphasis": "这套方案不仅具备理论深度，更为申请软件著作权和实际智能家居落地提供了完整支撑。",
        "transition_out": "下面向各位老师汇报第三大部分：我们已经完成的研究基础与实测结果，请看第 9 页。",
    },
    9: {
        "estimated_seconds": 85,
        "transition_in": "目前，我们已经完成了硬件采集平台搭建与 Widar3.0 基线算法的全部复现工作。",
        "figure_guide": "请各位老师对比左侧的准确率柱状图与右侧《图四：CNN-GRU 基线 30 轮训练实录与混淆矩阵》。",
        "main_speech": "在硬件层面，课题组已攻克 Intel 5300 网卡底层驱动与内核编译难关，成功解析出原始 CSI 幅相数据包。在算法层面，我们基于 Python 和 TensorFlow/Keras 跑通了从 BVP 输入到 CNN-GRU 分类的全流程。请看左侧柱状图与右侧图四终端实测数据：经过 30 轮（Epochs）训练后，训练集准确率达到 91.89%，验证集准确率达到 85.47%（混淆矩阵对角线清晰）。同时我们也观察到，验证集 Loss 在第 11 轮后出现明显波动并进入平台期，存在约 6.42% 的泛化性能缺口。",
        "key_emphasis": "这一实测结果既证实了 BVP 特征的有效性（85.47% 高基线），又直观暴露了模型过拟合与域泛化瓶颈，为下一步引入 DANN 提供了最直接的实验依据！",
        "transition_out": "除了算法复现，我们还同步完成了应用系统原型的开发，请看第 10 页。",
    },
    10: {
        "estimated_seconds": 70,
        "transition_in": "第 10 页展示的是我们已经开发完成的“WiFi 信号行为识别系统”前端原型。",
        "figure_guide": "请各位老师看右侧《图三：已开发的 WiFi 信号行为识别系统 UI 界面概览》以及左侧的三层架构。",
        "main_speech": "为了探寻算法的工程实用性，我们采用前后端分离架构开发了可视化监测平台。如图三所示，系统已集成用户登录鉴权、实时信号强度监测趋势图（如当前 -68 dBm 实时波形）、在线设备管理以及右侧的行为分类统计圆环图。目前前端软件框架已全部就绪，并在后端预留了标准化的 Python 模型推理接口（如中间流程图所示：CSI 实时抓包 → Socket 传输 → Python 推理 → Web 面板呈现）。",
        "key_emphasis": "待后续域自适应模型调优与量化完成后，即可直接装载进该系统实现实时手势识别闭环。",
        "transition_out": "最后，针对当前现存问题与后续时间安排，请看第 11 页。",
    },
    11: {
        "estimated_seconds": 75,
        "transition_in": "在已有研究基础上，我们也清醒地梳理了现存的三个问题，并制定了针对性的解决途径与五阶段进度计划。",
        "figure_guide": "请看上半部分左列“现存瓶颈问题”与右列“针对性解决途径”的一一对应关系，以及下半部分的五阶段时间轴。",
        "main_speech": "针对问题一（复杂动态跨域数据集构建），我们将制定严格的控制变量采集协议，并在预处理端引入自适应滤波与子载波优选算法；针对问题二（域对抗训练中的负迁移与收敛平衡），我们采用“先预训练特征提取器、再对抗微调”的两阶段策略，并设计动态权重因子平衡分类与对抗损失；针对问题三（从离线模型到实时系统的延迟瓶颈），我们将利用 Socket/RESTful API 打通数据流，并采用 NumPy 矩阵加速与模型量化（Quantization）满足毫秒级响应。下方时间轴明确了从 2025 年 11 月至 2027 年 3 月的五阶段推进节点。",
        "key_emphasis": "综合来看，本课题理论依据充分、基线实验（85.47%）与 UI 原型均已跑通、硬件成本可控，整体计划切实可行。",
        "transition_out": "我的开题报告核心内容汇报完毕，请翻至最后一页。",
    },
    12: {
        "estimated_seconds": 20,
        "transition_in": "以上就是我关于《基于域自适应的 Wi-Fi 手势识别泛化方法研究》开题报告的全部汇报内容。",
        "figure_guide": "感谢各位评审老师的耐心聆听！",
        "main_speech": "本课题在实验数据集协议设计与域对抗损失函数调优方面还有许多需要进一步完善的地方，非常期待并恳请各位评审老师提出宝贵的批评与指导意见！",
        "key_emphasis": "谢谢各位老师！",
        "transition_out": "（汇报结束，准备记录并回答评审老师提问）",
    },
}

WIFI_SPEECH_SCRIPTS.update(WIFI_SPEECH_SCRIPTS_PART2)


def attach_scripts_to_plan(meta: dict, agenda_items: list, slides: list,
                           refs: list, is_wifi: bool = True) -> dict:
    # 【函数作用】将每一页的同源演讲逐字稿绑死到 plan 蓝图中，确保 PPT 画面、PPT 备注栏与独立逐字稿文档 100% 同步
    all_scripts = []
    # 第 1 页封面逐字稿
    s1 = WIFI_SPEECH_SCRIPTS[1] if is_wifi else {
        "estimated_seconds": 35,
        "transition_in": f"各位老师、同学好！今天由我为大家汇报《{meta.get('title', '')}》。",
        "figure_guide": "请看封面课题信息。",
        "main_speech": "本次汇报将围绕核心研究问题、方法架构设计、实验数据验证与未来工作计划展开。",
        "key_emphasis": "下面正式开始汇报。",
        "transition_out": "首先请看汇报目录。",
    }
    all_scripts.append({"slide_no": 1, "title": f"封面：{meta.get('title', '')}", "layout": "cover", **s1})

    # 第 2 页目录逐字稿
    s2 = WIFI_SPEECH_SCRIPTS[2] if is_wifi else {
        "estimated_seconds": 30,
        "transition_in": "请看汇报目录。",
        "figure_guide": f"本次汇报涵盖：{'、'.join(agenda_items)}。",
        "main_speech": "我们将按循序渐进的科研逻辑逐一展开。",
        "key_emphasis": "首先进入第一部分。",
        "transition_out": "请看第 3 页。",
    }
    all_scripts.append({"slide_no": 2, "title": "汇报目录", "layout": "agenda", **s2})

    # 第 3 ~ N-1 页正文逐字稿
    for idx, sl in enumerate(slides, 3):
        if is_wifi and idx in WIFI_SPEECH_SCRIPTS:
            sc = WIFI_SPEECH_SCRIPTS[idx]
        else:
            b_txt = "；".join(sl.get("bullets", []))
            sc = sl.get("speaker_script") or {
                "estimated_seconds": 65,
                "transition_in": f"接下来我们看到第 {idx} 页：{sl.get('title', '')}。",
                "figure_guide": "请大家结合页面上的图表与核心卡片来看。",
                "main_speech": f"本页核心阐述的内容包括：{b_txt}。通过图表数据与理论机制的相互印证，支撑本章节的核心论点。",
                "key_emphasis": sl.get("banner", "该结果进一步验证了方案的有效性。"),
                "transition_out": "请看下一页。",
            }
        sl["speaker_script"] = sc
        all_scripts.append({"slide_no": idx, "title": sl.get("title", ""), "layout": sl.get("layout", ""), **sc})

    # 最后一页致谢逐字稿
    last_no = len(slides) + 3
    s_last = WIFI_SPEECH_SCRIPTS.get(12) if is_wifi else {
        "estimated_seconds": 20,
        "transition_in": "以上就是本次汇报的全部内容。",
        "figure_guide": "感谢各位老师的聆听！",
        "main_speech": "恳请各位老师批评指正，欢迎交流讨论。",
        "key_emphasis": "谢谢大家！",
        "transition_out": "（汇报结束）",
    }
    all_scripts.append({"slide_no": last_no, "title": "致谢与交流讨论", "layout": "thanks", **s_last})

    return {
        "meta": meta,
        "agenda": agenda_items,
        "slides": slides,
        "references": refs,
        "cover_script": s1,
        "agenda_script": s2,
        "thanks_script": s_last,
        "all_scripts": all_scripts,
    }
