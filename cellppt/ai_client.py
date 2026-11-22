# ==============================================================================
# 【文件作用】v0.3 AI 大模型统一调用底座 (cellppt/ai_client.py)
# 支持接入 DeepSeek、通义千问(Qwen)、Kimi、智谱GLM、OpenAI、Gemini 等任意大模型 API，
# 负责做三件事：
#   1. 通读学生的论文/开题报告全文，规划一步步递进的 10~15 页叙事剧本
#   2. 如果学生文档里没有参考文献，自动帮他生成该领域的权威顶刊参考文献并提炼检索词
#   3. 输出严格符合 8 大百变版式格式要求、且自带【每页同源演讲逐字稿】的 JSON 蓝图
# ==============================================================================

import json                      # 导入 JSON 解析与序列化模块
import os                        # 导入环境变量读取模块
import re                        # 导入正则表达式模块
from pathlib import Path         # 导入文件路径模块
import urllib.request            # 导入标准 HTTP 请求库（零额外依赖即可调用各大 AI 接口）
import urllib.error              # 导入 HTTP 异常处理模块

# 默认配置文件路径（位于项目根目录下的 ai_config.json）
CONFIG_PATH = Path(__file__).resolve().parent.parent / "ai_config.json"

# 常见国内与国际大模型 API 的默认 Base URL 映射表（方便学生只填 provider 名字就能用）
PROVIDER_URLS = {
    "deepseek": "https://api.deepseek.com/v1/chat/completions",
    "qwen": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
    "kimi": "https://api.moonshot.cn/v1/chat/completions",
    "glm": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
    "siliconflow": "https://api.siliconflow.cn/v1/chat/completions",
    "openai": "https://api.openai.com/v1/chat/completions",
}

# 默认推荐模型名称映射表
PROVIDER_MODELS = {
    "deepseek": "deepseek-chat",
    "qwen": "qwen-plus",
    "kimi": "moonshot-v1-32k",
    "glm": "glm-4-plus",
    "siliconflow": "deepseek-ai/DeepSeek-V3",
    "openai": "gpt-4o-mini",
}


def load_ai_config() -> dict:
    # 【函数作用】读取 AI 接口配置（优先读环境变量，其次读项目根目录下的 ai_config.json）
    # 后期微调：如果你想固定用自己的 DeepSeek 或通义千问 Key，只需在根目录 ai_config.json 填入 api_key 即可
    cfg = {
        "provider": os.environ.get("CELLPPT_PROVIDER", "deepseek"),
        "api_key": (
            os.environ.get("CELLPPT_API_KEY")
            or os.environ.get("DEEPSEEK_API_KEY")
            or os.environ.get("DASHSCOPE_API_KEY")
            or os.environ.get("OPENAI_API_KEY")
            or ""
        ),
        "api_url": os.environ.get("CELLPPT_API_URL", ""),
        "model": os.environ.get("CELLPPT_MODEL", ""),
        "temperature": 0.3,  # 【微调点】生成温度 0.3，保证学术汇报内容严谨、JSON 格式稳定
    }
    if CONFIG_PATH.is_file():
        try:
            file_cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            for k in ("provider", "api_key", "api_url", "model", "temperature"):
                if file_cfg.get(k):
                    cfg[k] = file_cfg[k]
        except Exception:
            pass
    prov = cfg["provider"].lower()
    if not cfg["api_url"]:
        cfg["api_url"] = PROVIDER_URLS.get(prov, PROVIDER_URLS["deepseek"])
    if not cfg["model"]:
        cfg["model"] = PROVIDER_MODELS.get(prov, "deepseek-chat")
    return cfg


def ensure_config_template() -> Path:
    # 【函数作用】如果项目根目录下还没有 ai_config.json，自动生成一个带中文说明的配置模板文件
    if not CONFIG_PATH.exists():
        template = {
            "_comment": "填入任意一家大模型的 api_key 即可开启实时 AI 深度规划（支持 deepseek / qwen / kimi / glm / siliconflow / openai）",
            "provider": "deepseek",
            "api_key": "",
            "api_url": "https://api.deepseek.com/v1/chat/completions",
            "model": "deepseek-chat",
            "temperature": 0.3,
        }
        CONFIG_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
    return CONFIG_PATH


def call_llm_json(system_prompt: str, user_prompt: str, timeout: int = 90) -> dict | None:
    # 【函数作用】向大模型发送请求并强制解析返回 JSON 字典；如果未配置 Key 或网络超时则返回 None（自动触发本地智能规划回退）
    ensure_config_template()
    cfg = load_ai_config()
    api_key = cfg.get("api_key", "").strip()
    if not api_key:
        return None  # 未配置 API Key 时平滑交由本地智能规划引擎处理

    payload = {
        "model": cfg["model"],
        "temperature": float(cfg.get("temperature", 0.3)),
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    }
    req = urllib.request.Request(
        cfg["api_url"],
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            content = body["choices"][0]["message"]["content"]
            # 剥离可能包裹在 JSON 外面的 ```json ... ``` 代码块标记
            content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content.strip(), flags=re.M)
            return json.loads(content)
    except Exception as exc:
        print(f"[AI 提示] 在线大模型调用未生效 ({exc})，已自动切换至本地深度语义规划引擎。")
        return None


def build_planner_system_prompt() -> str:
    # 【函数作用】构造给 AI 总导演的系统级指令：严格规定 8 种版式 Schema、四字红字提词原则、配图需求单与同源逐字稿格式
    return (
        "你是一位顶级学术汇报 PPT 总导演。请将用户的论文或开题报告规划为一套逻辑丝滑、图文呼应的 JSON 幻灯片蓝图。\n"
        "必须遵守以下 4 条铁律：\n"
        "1. 【版式百变】：正文页必须从以下 8 种版式中轮换选择，相邻两页严禁使用相同版式：\n"
        "   dual_compare, three_tier_cards, tri_zone_flow, arch_innovation, focus_challenges, four_grid_metrics, data_chart_and_shot, roadmap_solutions。\n"
        "2. 【极简提词原则】：页面上的文字是给汇报人提词、给听众看图抓重点的锚点，每条要点必须采用“四字核心词：15~25字高密度结论（核心词用 **包裹** 标红）”。\n"
        "3. 【配图需求单】：如果某页缺少原始图片，必须在该页 JSON 中填写 image_query 字段（包含英文学术检索词 search_keywords、关联参考文献编号 ref_hint、以及期望展示的科学原理 diagram_topic）。\n"
        "4. 【同源演讲逐字稿】：每一页必须包含 speaker_script 字段，内含 estimated_seconds(预计秒数)、transition_in(承上引入词)、main_speech(开口逐字稿，必须明确指图讲解)、key_emphasis(重点强调句)、transition_out(启下翻页词)。\n"
        "只输出合法的 JSON 对象，不要输出任何额外解释。"
    )
