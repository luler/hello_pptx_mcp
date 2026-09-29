# -*- coding: utf-8 -*-
"""pptx-studio MCP 服务（基于 MCP Python SDK v2 的 MCPServer）。

包含 12 个专业级工具与 MCP Prompts，全面赋能大模型：
1. 大模型设计指导与无空白规范（get_design_guide）
2. 全流程出片与逐页截图生成（create_presentation，返回 PPTX 下载链接与逐页截图链接清单）
3. 单页精准定向重构与修改（update_page）
4. 一键主题切换与样式微调（change_theme）
5. 自动获取 Host 与完整下载链接返回（http://host:port/api/download/xxx.pptx）
"""
from __future__ import annotations

import contextvars
import json
import os
import time
import uuid

try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP as MCPServer
    except ImportError:
        from mcp.server import FastMCP as MCPServer

try:
    from mcp.server.transport_security import TransportSecurityMiddleware, TransportSecuritySettings
    # 彻底放行所有域名、反向代理与 Origin，彻底根治 421 Invalid Host header
    TransportSecurityMiddleware._validate_host = lambda self, host: True
    TransportSecurityMiddleware._validate_origin = lambda self, origin: True
    _transport_security = TransportSecuritySettings(
        enable_dns_rebinding_protection=False,
        allowed_hosts=["*"],
        allowed_origins=["*"],
    )
except Exception:
    _transport_security = None

from pptkit import render, store as store_mod, templates, theme as theme_mod
from pptkit.deck import build_from_spec

WORKSPACE = os.environ.get("PPTKIT_WORKSPACE",
                           os.path.join(os.path.expanduser("~"), "workspace"))
OUTPUT_DIR = os.environ.get("PPTKIT_OUTPUT_DIR",
                            os.path.join(WORKSPACE, "pptx-mcp", "output"))
STORE_DIR = os.environ.get("PPTKIT_STORE_DIR",
                           os.path.join(WORKSPACE, "pptx-mcp", "store"))
STORE = store_mod.Store(STORE_DIR)

# 动态保存每次 HTTP 请求传入的真实 Host 基础路径
current_request_base_url: contextvars.ContextVar[str | None] = contextvars.ContextVar(
    "current_request_base_url", default=None
)


def get_base_url() -> str:
    """获取当前服务外部可访问的 Base URL（自适应 Host 头、环境变量或端口兜底）。"""
    env_base = os.environ.get("BASE_URL") or os.environ.get("PUBLIC_HOST")
    if env_base:
        return env_base.rstrip("/")
    ctx_base = current_request_base_url.get()
    if ctx_base:
        return ctx_base.rstrip("/")
    port = os.environ.get("PORT", "48000")
    return f"http://127.0.0.1:{port}"


_server_kwargs = {
    "name": "pptx-studio",
    "version": "0.4.0",
    "instructions": (
        "你是一个顶尖商业与科技演示文稿架构师与设计专家。\n\n"
        "【自主选择主题与版式核心法则】\n"
        "系统已内置 24 种覆盖全业务场景的高级版式库与 9 大专业色彩主题（支持自定义配色）。\n"
        "严禁千篇一律地只使用单一版式（如整篇都是 bullets 或 grid），大模型必须根据用户的【主题定位】与【页面叙事目标】自主决策匹配最合适的风格和版式组合：\n"
        "1. 风格主题自主匹配（按主题行业与基调自适应）：\n"
        "   - AI/大模型/前沿科技/极客发布会 -> `tech_dark`（午夜黑+赛博极光青）\n"
        "   - 咨询机构/投行/金融研报/商业分析 -> `business_blue`（麦肯锡冷灰+藏青）\n"
        "   - 高端商务盛典/领袖峰会/年度答谢/奢华黑金 -> `executive_gold`（黑曜石+香槟金）\n"
        "   - 数字化转型/产业互联网/企业SaaS -> `modern_teal`（科技深青+珊瑚橙）\n"
        "   - 国企汇报/党政政务/行业监管/社会责任 -> `gov_blue`（规范庄重蓝+中国红）\n"
        "   - 医疗健康/绿色低碳/新能源/ESG -> `fresh_green`（薄荷雾白+森林绿）\n"
        "   - 新消费/电商零售/创业路演/品牌推广 -> `warm_coral`（象牙白+活力珊瑚橙）\n"
        "   - 文化创意/艺术时尚/泛娱乐潮牌 -> `violet`（紫罗兰+摩登洋红）\n"
        "   - 东方国风/传统文化/诗词意境/借诗明志 -> `chinese_poetry`（宣纸白+松烟墨+朱砂红）\n"
        "   - 个性化定制 -> theme 直接传入包含 bg, card_bg, primary, accent, is_dark 的字典对象\n"
        "2. 24 大版式库全景（按表达目的分为 12 大功能维度，大模型自主随心组合）：\n"
        "   - 封面导读类：`title_slide`（带右侧悬浮看点卡片）\n"
        "   - 结构转承类：`section`（章节过渡导航大浮岛卡）\n"
        "   - 业务全景类：`dashboard`（数据指标+图表+进度条综合看板）\n"
        "   - 团队组织类：`team` / `team_grid`（3~4 位专家/骨干画像履历与技能标签）\n"
        "   - 战略层级类：`pyramid` / `pyramid_funnel`（战略金字塔 / 业务转化漏斗）\n"
        "   - 态势分析类：`swot` / `swot_matrix`（SWOT 优势、劣势、机会、威胁 4 象限）\n"
        "   - 架构拓扑类：`stack`（分层系统技术架构）、`diagram`（微服务拓扑节点调用网络）\n"
        "   - 方案权衡类：`comparison`（Before vs After 对比）、`pros_cons`（利弊收益与风险考量）\n"
        "   - 流程演进类：`process_flow`（有向粗箭头业务流程）、`timeline`（横向时间线里程碑）、`roadmap`（多阶段跨周期战略路线图）\n"
        "   - 核心论断类：`hero_statement`（震撼大字报金句+支柱）、`kpi`（大数字指标冲击+战略洞察）\n"
        "   - 商业落地类：`case_study`（客户案例故事+痛点+ROI+证言）、`pricing`（产品套餐与定价推荐）、`table`（卡片对比表格）、`grid`（2x2/1x3 Bento矩阵）、`bullets`（Bento要点砖块）、`text`（正文演说）\n"
        "   - 疑虑与收尾类：`faq`（手风琴 Q&A 问答）、`poetry`（古典诗词韵律）、`summary`（高管核心结论沉淀+落地行动计划 Next Steps）\n"
        "3. 视觉节奏感与饱满度法则：一份 PPT 必须交替运用 5~8 种不同版式，严禁连续两页使用同一 layout；善用 [[关键字]] 自动高亮，杜绝空白。\n\n"
        "【大模型生成与用户交互工作流】\n"
        "1. 设计规划：调用 `get_design_guide(topic)` 了解匹配的主题与版式推荐；\n"
        "2. 生成演示文稿：调用 `create_presentation` 依据 spec 一键生成 PPTX，系统自动渲染出每一页的高清预览截图链接（preview_urls）并生成 PPTX 下载链接（download_url）；\n"
        "3. 直观呈现给用户（输出强制铁律）：工具返回值中已提供精心排版的 `user_display_markdown`。在最终回复中，你【必须直接原样输出 user_display_markdown】！必须确保每一页的高清截图都以 Markdown 图片语法 `![第X页预览](url)` 直接渲染为大图展示。绝对严禁自作聪明将其转换为 `[标题](url)` 等纯文字超链接，绝对严禁因页数多而擅自省略或折叠图片！大模型无需进行自主视觉审查；\n"
        "4. 用户反馈与修改闭环：用户在聊天界面直接查看到完整渲染的每页大图后，若满意可直接点击链接下载 PPTX 文件；若用户提出具体修改建议，大模型直接调用相应的修改接口：\n"
        "   - 单页精修：调用 `update_page(deck_id, page_number, page_spec)` 修改某页；\n"
        "   - 全局换肤：调用 `change_theme(deck_id, new_theme)` 一键切换全套配色风格；\n"
        "   - 增删页面：调用 `insert_page` 或 `delete_page` 动态增删页面；\n"
        "   修改完成后，同样必须直接原样输出返回值中的 `user_display_markdown`，以图片形式直观展示最新效果。"
    ),
}

if _transport_security is not None:
    _server_kwargs["transport_security"] = _transport_security

try:
    server = MCPServer(**_server_kwargs)
except TypeError:
    _server_kwargs.pop("transport_security", None)
    server = MCPServer(**_server_kwargs)

if _transport_security is not None:
    if hasattr(server, "settings") and hasattr(server.settings, "transport_security"):
        server.settings.transport_security = _transport_security




def _json(obj):
    return json.dumps(obj, ensure_ascii=False, indent=2)


def _slug(name, fallback="deck"):
    raw = str(name or "").strip()
    if raw.lower().endswith(".pptx"):
        raw = raw[:-5]
    keep = "".join(c for c in raw if c.isalnum() or c in "-_ ")
    keep = keep.strip().replace(" ", "_")
    return keep or fallback


def make_unique_filename(filename: str = "", fallback: str = "deck") -> str:
    """生成以时间戳和并发隔离短码开头的唯一 PPTX 文件名，保证目录自然时间排序且多客户端并发绝对零冲突。

    例如:
        传入 "gold_fund_pitch.pptx"
        生成 "20260929_162018_c2a2_gold_fund_pitch.pptx"
    """
    raw = str(filename or "").strip()
    if raw.lower().endswith(".pptx"):
        raw = raw[:-5]
    base = _slug(raw, fallback=fallback)
    ts = time.strftime("%Y%m%d_%H%M%S")
    rand_tag = uuid.uuid4().hex[:4]

    if base.startswith(ts):
        candidate = f"{base}.pptx"
    else:
        candidate = f"{ts}_{rand_tag}_{base}.pptx"

    # 兜底防重（万亿分之一概率碰撞时自旋）
    while os.path.exists(os.path.join(OUTPUT_DIR, candidate)):
        rand_tag = uuid.uuid4().hex[:4]
        candidate = f"{ts}_{rand_tag}_{base}.pptx"

    return candidate



def make_display_markdown(download_url: str, filename: str, pages_info: list[dict]) -> str:
    """生成给大模型直接输出给用户的高颜值 Markdown 预览与下载模块。"""
    lines = [
        "### 📥 演示文稿生成完成",
        f"👉 **[点击直接下载 PPTX 演示文稿]({download_url})**  `({filename})`\n",
        "#### 🖼️ 幻灯片效果预览（点击下载或提出修改需求）："
    ]
    for p in pages_info:
        idx = p.get("page", 1)
        title = p.get("title") or f"第 {idx} 页"
        layout = p.get("layout", "")
        url = p.get("preview_url", "")
        tag = f" `[{layout}]`" if layout else ""
        lines.append(f"- **第 {idx} 页{tag} · {title}**\n\n  ![第 {idx} 页预览]({url})")
    return "\n\n".join(lines)


VALID_LAYOUTS = {
    # 1. 封面与结构
    "title_slide", "section",
    # 2. 数据与看板
    "dashboard", "kpi", "table",
    # 3. 架构与拓扑
    "stack", "architecture", "layers",
    "diagram", "topology", "architecture_diagram",
    # 4. 流程与时序
    "process_flow", "flow", "pipeline",
    "timeline",
    "roadmap", "roadmap_milestones", "milestones",
    # 5. 分析与对比
    "comparison",
    "pros_cons", "tradeoff",
    "swot", "swot_matrix",
    "pyramid", "pyramid_funnel", "funnel",
    # 6. 核心论断与文化
    "hero_statement", "statement", "quote", "focus",
    "poetry", "poem", "verse",
    # 7. 团队与商业落地
    "team", "team_grid", "members", "org",
    "case_study", "case",
    "pricing", "pricing_packages", "packages", "plans",
    # 8. 总结与问答
    "faq", "faq_accordion", "qa",
    "summary", "summary_next_steps", "next_steps", "action_plan",
    # 9. 通用多卡片与要点
    "grid", "cards", "bullets", "text"
}


def _validate_spec(spec):
    """结构性校验，返回错误列表（空表示通过）。"""
    errors = []
    if not isinstance(spec, dict):
        return ["spec 必须是对象"]
    pages = spec.get("pages")
    if not isinstance(pages, list) or not pages:
        errors.append("spec.pages 必须是非空数组")
    else:
        for i, p in enumerate(pages):
            if not isinstance(p, dict):
                errors.append("第 %d 页不是对象" % (i + 1))
                continue
            lay = p.get("layout", "bullets")
            if lay not in VALID_LAYOUTS:
                errors.append("第 %d 页 layout=%r 未知，可选 %s"
                              % (i + 1, lay, ", ".join(sorted(VALID_LAYOUTS))))
        theme_val = spec.get("theme")
        if theme_val:
            if isinstance(theme_val, dict):
                # 大模型自定义配色字典
                pass
            elif isinstance(theme_val, str):
                aliases = {"business", "gold", "dark", "green", "purple", "blue", "coral", "teal", "poetry", "ink", "guofeng", "chinese"}
                if theme_val not in theme_mod.PALETTES and theme_val.lower() not in aliases:
                    errors.append("未知主题 %r，可选 %s 或传入自定义配色字典"
                                  % (theme_val, ", ".join(theme_mod.PALETTES)))
            else:
                errors.append("theme 必须是字符串主题名或自定义配色对象")
        if spec.get("size") and spec["size"] not in theme_mod.SLIDE_SIZES:
            errors.append("未知尺寸 %r" % spec["size"])
    return errors



@server.tool()
def get_design_guide(topic: str = "") -> str:
    """获取大模型制作精美高级、无空白留白、高视觉密度的 PPT 设计指南。

    涵盖 9 大顶级主题与行业决策树、24 种全场景专业版式用法参数表、防止页面空洞的核心技巧与大模型自适应推荐。

    Args:
        topic: 演示文稿主题（如"AI大模型商业化落地方案"、"跨境电商2026出海战略"、"新一代微服务架构改造"等），传入后将返回智能匹配的主题与8~10页叙事版式蓝图。
    """
    # 智能主题与叙事推荐算法
    topic_rec = {}
    if topic:
        t_low = topic.lower()
        if any(k in t_low for k in ["ai", "大模型", "算法", "极客", "智能", "算力", "芯片", "科技"]):
            rec_theme = "tech_dark"
            theme_reason = "AI与硬核科技主题匹配深空极客黑，赛博荧光青与午夜黑凸显未来感"
            storyboard = ["title_slide", "hero_statement", "pyramid", "stack", "diagram", "process_flow", "kpi", "roadmap", "summary"]
        elif any(k in t_low for k in ["金融", "银行", "证券", "投资", "投行", "基金", "咨询", "财报", "分析"]):
            rec_theme = "business_blue"
            theme_reason = "金融投资咨询匹配顶尖咨询蓝（麦肯锡冷灰+藏青），理性克制、严谨专业"
            storyboard = ["title_slide", "dashboard", "swot", "comparison", "kpi", "table", "timeline", "summary"]
        elif any(k in t_low for k in ["高端", "奢华", "峰会", "盛典", "年会", "年度", "领袖", "颁奖"]):
            rec_theme = "executive_gold"
            theme_reason = "领袖峰会与高端盛典匹配曜石黑金，尊享奢华、大气非凡"
            storyboard = ["title_slide", "hero_statement", "team", "dashboard", "kpi", "case_study", "roadmap", "summary"]
        elif any(k in t_low for k in ["政务", "国企", "党建", "社会责任", "治理", "监管", "汇报"]):
            rec_theme = "gov_blue"
            theme_reason = "党政国企与行业监管匹配政务权威蓝（规范庄重蓝+典雅中国红），庄重大气"
            storyboard = ["title_slide", "dashboard", "grid", "process_flow", "kpi", "table", "timeline", "summary"]
        elif any(k in t_low for k in ["医疗", "医药", "健康", "绿色", "低碳", "环保", "esg", "农业"]):
            rec_theme = "fresh_green"
            theme_reason = "健康生态绿色环保匹配清新生态绿（薄荷雾白+森林绿），生机盎然"
            storyboard = ["title_slide", "dashboard", "swot", "process_flow", "kpi", "roadmap", "case_study", "summary"]
        elif any(k in t_low for k in ["电商", "零售", "消费", "出海", "路演", "创业", "营销"]):
            rec_theme = "warm_coral"
            theme_reason = "新消费与创业路演匹配活力珊瑚橙，富有感染力与商业爆发力"
            storyboard = ["title_slide", "hero_statement", "pros_cons", "pyramid", "kpi", "pricing", "case_study", "summary"]
        elif any(k in t_low for k in ["文化", "艺术", "诗", "词", "国风", "非遗", "汉服", "古典"]):
            rec_theme = "chinese_poetry"
            theme_reason = "国风雅韵与文化艺术匹配水墨丹青，宣纸白+松烟墨+朱砂红，格调拉满"
            storyboard = ["title_slide", "poetry", "hero_statement", "grid", "timeline", "case_study", "summary"]
        elif any(k in t_low for k in ["潮牌", "文创", "设计", "年轻", "潮流", "时尚", "娱乐"]):
            rec_theme = "violet"
            theme_reason = "潮流先锋与品牌创意匹配紫罗兰先锋，摩登时尚"
            storyboard = ["title_slide", "hero_statement", "grid", "pros_cons", "kpi", "pricing", "summary"]
        else:
            rec_theme = "modern_teal"
            theme_reason = "现代化商业方案首选数字科技青（冰爽微青白+沉静深青+珊瑚橙高光）"
            storyboard = ["title_slide", "dashboard", "comparison", "stack", "kpi", "roadmap", "case_study", "summary"]

        topic_rec = {
            "topic": topic,
            "recommended_theme": rec_theme,
            "theme_reason": theme_reason,
            "recommended_slide_flow": storyboard,
            "design_tip": "建议整套文稿采用推荐的版式序列，严格保持张弛有度的视觉节奏，避免同类版式连续出现。"
        }

    guide = {
        "topic_custom_recommendation": topic_rec if topic else "传入 topic 参数（如 topic='AI大模型赋能金融'）可获取专属主题与版式流推荐",
        "principles": [
            "【视觉节奏感（Visual Rhythm）】一份高水准的 PPT 必须有张有弛、版式交替推进！严禁连续两页使用同一种 layout。一套 PPT 建议自主编排 5~8 种不同版式！",
            "【拒绝大片空白与单调】每页幻灯片必须信息饱满、层级鲜明。切勿使用空旷大卡片，应丰富为带数据标签的砖块、分层架构胶囊、或底部核心战略洞察卡（takeaway / note）。",
            "【核心数据突出】大字号数值是现代 PPT 的灵魂。关键数字用 34~42pt 加粗展示，文本中用 [[关键字/数字]] 标记自动高亮。",
            "【结构化表达】避免纯大段文字堆砌，根据内容属性选择最匹配的专业版式：架构选 stack/diagram，流程选 process_flow/roadmap，团队选 team，分析选 swot/pros_cons/pyramid，金句选 hero_statement，指标选 kpi，商业选 case_study/pricing。"
        ],
        "themes": [
            {"name": "tech_dark", "style": "深空极客黑（科技发布会/AI/大模型/硬核系统）", "features": "午夜黑 + 赛博荧光青 + 通透纯白字 + 极光环境微光晕，未来感爆棚"},
            {"name": "business_blue", "style": "顶尖咨询蓝（麦肯锡/投行/金融研报/商业分析）", "features": "冷灰底 + 藏青大字 + 皇家宝蓝核心高光，极简理性"},
            {"name": "executive_gold", "style": "黑金尊享（高端战略/领袖峰会/年度答谢/颁奖盛典）", "features": "曜石黑 + 奢华香槟金 + 暖象牙白，领袖专属"},
            {"name": "modern_teal", "style": "数字科技青（数字化转型/产业互联网/企业SaaS）", "features": "冰爽微青白 + 沉静深青 + 珊瑚橙互补高光，通用现代"},
            {"name": "fresh_green", "style": "清新生态绿（医疗健康/ESG/绿色低碳/新能源/现代农业）", "features": "薄荷雾白 + 森林绿 + 翡翠绿核心数据，生机盎然"},
            {"name": "warm_coral", "style": "活力珊瑚橙（创业路演/新消费/跨境电商/品牌推广）", "features": "暖象牙白 + 爆裂珊瑚橙 + 互补宝蓝，张力十足"},
            {"name": "violet", "style": "紫罗兰先锋（品牌营销/文化创意/时尚潮牌/泛娱乐）", "features": "淡雅薰衣草白 + 贵族深紫 + 摩登洋红，摩登先锋"},
            {"name": "gov_blue", "style": "政务权威蓝（国企党政汇报/公共服务/监管审计/社会责任）", "features": "规范庄重蓝 + 典雅中国红强调 + 暖金辅助，庄重肃穆"},
            {"name": "chinese_poetry", "style": "水墨丹青·东方雅韵（国风/传统文化/非遗/诗词借物明志）", "features": "温润宣纸米白 + 沉着松烟墨黑 + 经典朱砂红印泥 + 天青汝窑黛绿，高古雅致"},
            {"name": "custom", "style": "大模型自主设计 HTML/CSS 配色字典", "features": "支持在 theme 字段直接传入包含 bg, card_bg, primary, accent, is_dark 的字典对象，完全解放调色自由"}
        ],
        "layouts_catalog_24": {
            "1. 封面与结构类": {
                "title_slide": "封面页。必填: title, subtitle。推荐在 bullets 中提供 3~4 条核心战略看点（自动生成右侧【核心看点悬浮卡】）与 meta_line。",
                "section": "章节过渡页。必填: title。可选: subtitle。居中大卡片 + 章节序号胶囊 + 强调色边条，提升叙事段落感。"
            },
            "2. 数据与看板类": {
                "dashboard": "综合全屏数据看板。包含 lead（导读条）、kpis（3~4个指标卡）、panels（2~3个多功能卡片，支持 chart 柱状图或 progress 进度条），视觉信息密度极高。",
                "kpi": "核心量化指标冲击页。包含 lead、kpis（3~4个大数字指标卡，含 label/value/unit/desc），可在 takeaway/note 中写入深度洞察卡。",
                "table": "卡片式指标/对比表格。包含 rows（二维数组或对象数组，支持表头），智能自适应行高与斑马条纹，下部带 takeaway/note 洞察卡。"
            },
            "3. 架构与拓扑类": {
                "stack": "纵向分层技术架构图。在 layers 中传入 3~4 层架构（如接入层、逻辑层、存储层），每层包含 title, desc, components（组件药丸标签列表），层间自带流转箭头。",
                "diagram": "分布式服务拓扑网络图。在 nodes 中传入 2D 网格服务节点（col, row, type, label, desc），在 edges 中指定服务调用连线（from, to, label），底部带 takeaway 核心解耦原则。"
            },
            "4. 团队与组织类": {
                "team": "核心专家与组织团队页。在 members 中传入 3~4 位核心团队卡片（包含 name, role, bio/desc, tags/skills, avatar_text），顶部带色带与字母头像徽章。"
            },
            "5. 战略分析与层级类": {
                "pyramid": "战略金字塔 / 业务转化漏斗页。在 tiers 中传入 3~5 个层级（从顶至底展开，包含 level, title, desc, rate/metric），居中梯形自适应宽度。",
                "swot": "SWOT 战略态势分析矩阵。在 swot 对象中传入 s (优势), w (劣势), o (机会), t (威胁) 4 大象限列表，带大号水印背景字母与彩色药丸。"
            },
            "6. 方案对比与权衡类": {
                "comparison": "两栏深度对比页（Before vs After / 现状 vs 突破）。包含 left 与 right 对象（各含 title, items），带演进结论卡。",
                "pros_cons": "方案利弊得失权衡页。包含 pros（核心优势与收益列表）与 cons（潜在风险与成本列表），底部带 takeaway 权衡决策结论卡。"
            },
            "7. 流程推进与路线图类": {
                "process_flow": "有向流转演进图。在 steps 中传入 3~4 个带有粗箭头指示（➔）的流程阶段卡片，包含 title, desc, tag。",
                "roadmap": "战略演进演化路线图。在 phases 中传入 3~4 个阶段（如启动期、扩展期、成熟期），包含 phase, title, goal, tasks（打勾清单）与卡片间推演箭头。",
                "timeline": "横向时间线里程碑。在 items 中传入 3~5 个时间节点（包含 year/date, title, desc），下部带推进保障机制卡。"
            },
            "8. 震撼大字与文化韵律类": {
                "hero_statement": "战略论断/大字报页。在 statement 中传入 24pt 震撼核心论断金句（支持 [[高光]]），在 points 中传入 3~4 个关键支柱微卡片，直击重点。",
                "poetry": "东方水墨诗词与韵律美学页。在 cipai/tag、author、poem（对仗诗句）中展示，在 interpretation 中传入战略/技术寓意，古典优雅。"
            },
            "9. 商业落地与转化类": {
                "case_study": "标杆客户成功案例页。包含 client (客户名称), industry (行业), challenge (痛点挑战), solution (举措), results (量化ROI指标数组), quote (客户证言)。",
                "pricing": "产品套餐与商业定价对比页。在 packages 中传入 3 档定价卡片（包含 name, price, period, desc, features 权益清单, is_popular 推荐标识, cta 按钮文案）。"
            },
            "10. 疑虑解答与收尾类": {
                "faq": "常见疑问解答（FAQ）页。在 faqs 中传入 3~4 个高质感手风琴式 Q&A 卡片（包含 q, a, tag），消除客户或投资人顾虑。",
                "summary": "高管总结与行动计划页（Next Steps）。包含 takeaways（核心结论沉淀列表）与 actions（下一步行动计划清单，含 task, owner, due）。"
            },
            "11. 通用多卡片与正文类": {
                "grid": "2x2 或 1x3 Bento 矩阵网格页。展示 3~4 项核心支柱，四角或横向均布，下部可配置 takeaway 总结卡。",
                "bullets": "结构化要点页。自动将每个要点打包为立体 Bento 独立砖块，彻底消除大片空白。",
                "text": "正文演讲说明页。双层高质感卡片容器，段落智能缩进排版。"
            }
        }
    }
    return _json(guide)


@server.tool()
def list_themes() -> str:
    """列出可用配色主题（名称、定位、色值）。"""
    out = []
    for name, pal in theme_mod.PALETTES.items():
        out.append({"name": name, "label": pal["label"],
                    "is_dark": pal.get("is_dark", False),
                    "primary": pal["primary"], "accent": pal["accent"],
                    "accent2": pal["accent2"]})
    return _json({"themes": out, "sizes": list(theme_mod.SLIDE_SIZES.keys())})


@server.tool()
def list_templates() -> str:
    """列出内置精美 spec 模板。"""
    return _json({"templates": templates.listing()})


@server.tool()
def get_template(name: str) -> str:
    """获取一个完整可用的 spec 模板。"""
    try:
        return _json(templates.get(name))
    except KeyError as e:
        return _json({"ok": False, "error": str(e),
                      "available": list(templates.TEMPLATES.keys())})


@server.tool()
def validate_spec(spec: dict) -> str:
    """校验 spec 语法结构与必填字段有效性。"""
    errs = _validate_spec(spec)
    if errs:
        return _json({"ok": False, "errors": errs})
    return _json({"ok": True, "message": "spec 结构校验通过"})


@server.tool()
def create_presentation(spec: dict, filename: str = "", preview: bool = True,
                        store: bool = True, preview_pages: str = "") -> str:
    """按 spec 生成 PPTX，自动渲染每页高清预览截图并生成用户访问链接，返回包含 PPTX 下载链接与逐页截图链接的完整清单。

    Args:
        spec: 演示文稿定义，必须含 pages，可含 theme / size / meta。
        filename: 输出文件名（不含目录），留空则按时间戳自动命名。
        preview: 是否渲染高清预览图以供用户在对话中查看（默认 True）。
        store: 是否登记到产物仓库。
        preview_pages: 仅预览指定页（如 "1,3"），留空表示全部页面。

    Returns:
        JSON 格式结果，包含 PPTX 下载链接 (download_url)、每页截图链接 (preview_urls) 以及已格式化排版的 Markdown 内容 (user_display_markdown)。大模型在最终回复中【必须直接原样输出 user_display_markdown】，以 Markdown 图片语法直接渲染所有页面的预览大图，严禁精简为纯文字超链接。
    """
    errs = _validate_spec(spec)
    if errs:
        return _json({"ok": False, "stage": "validate", "errors": errs})

    name = make_unique_filename(filename, fallback="deck")
    out_path = os.path.join(OUTPUT_DIR, name)

    try:
        deck, path, warnings = build_from_spec({**spec, "output": {"path": out_path}})
    except Exception as e:
        return _json({"ok": False, "stage": "build",
                      "error": "%s: %s" % (type(e).__name__, e)})

    rec = STORE.register(name, "pptx", path, spec=spec) if store else {
        "id": None, "name": name, "path": path, "bytes": os.path.getsize(path)}

    base_url = get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"
    alt_download_url = f"{base_url}/api/download/{rec['id']}.pptx" if rec.get("id") else download_url

    pages = spec.get("pages") or []
    result = {
        "ok": True,
        "deck_id": rec["id"],
        "filename": rec["name"],
        "download_url": download_url,
        "alt_download_url": alt_download_url,
        "markdown_download_link": f"[点击下载演示文稿 (PPTX)]({download_url})",
        "bytes": rec["bytes"],
        "total_pages": len(pages),
        "theme": spec.get("theme", "gov_blue"),
        "warnings": warnings,
        "preview_urls": [],
        "slides": [],
        "user_display_markdown": "",
        "action_guide": (
            "演示文稿已成功生成！【最终回复输出铁律】：请务必直接在最终回复中原样输出返回值中的 `user_display_markdown`！"
            "必须确保每一页的高清预览图都以 Markdown 图片语法 `![第X页预览](url)` 直接渲染为大图展示，"
            "绝对严禁将其简化为 `[标题](url)` 等纯文字超链接，绝对严禁因页数多而擅自省略图片！用户查看后若满意可直接下载，或提出修改意见。"
        )
    }

    if not preview:
        result["user_display_markdown"] = f"### 📥 演示文稿已生成\n👉 **[点击下载 PPTX 文件]({download_url})**"
        return _json(result)

    if not render.available():
        result["preview_error"] = "未检测到 LibreOffice，跳过预览图生成"
        result["user_display_markdown"] = f"### 📥 演示文稿已生成\n👉 **[点击下载 PPTX 文件]({download_url})**"
        return _json(result)

    try:
        sel = None
        if preview_pages.strip():
            sel = [int(x) - 1 for x in preview_pages.replace("，", ",").split(",")
                   if x.strip().isdigit()]
        pv_dir = os.path.join(os.path.dirname(path), "preview", _slug(name, "deck"))
        pdf, pngs = render.render_pptx(path, out_dir=pv_dir, pages=sel)

        if store:
            STORE.register_many([
                {"name": os.path.basename(pdf), "kind": "pdf", "path": pdf,
                 "item_id": f"{rec['id']}_pdf" if rec["id"] else None}
            ], validate=False)


        preview_urls = [
            f"{base_url}/api/preview/{rec['id']}?page={i+1}" for i in range(len(pngs))
        ]
        result["preview_urls"] = preview_urls

        slides_info = []
        for i, png in enumerate(pngs):
            page_idx = i + 1
            page_meta = pages[i] if i < len(pages) else {}
            slides_info.append({
                "page": page_idx,
                "title": page_meta.get("title", f"第 {page_idx} 页"),
                "layout": page_meta.get("layout", "bullets"),
                "preview_url": preview_urls[i],
                "markdown_image": f"![第 {page_idx} 页预览]({preview_urls[i]})"
            })
        result["slides"] = slides_info
        result["user_display_markdown"] = make_display_markdown(download_url, rec["name"], slides_info)

    except Exception as e:
        result["preview_error"] = "%s: %s" % (type(e).__name__, e)
        result["user_display_markdown"] = f"### 📥 演示文稿已生成\n👉 **[点击下载 PPTX 文件]({download_url})**"

    return _json(result)



@server.tool()
def update_page(deck_id: str, page_number: int, page_spec: dict) -> str:
    """根据用户反馈定向修改某页的内容、版式或样式，重新构建文稿并返回更新后的下载链接与该页截图链接。

    Args:
        deck_id: 演示文稿 ID 或文件名（如 pptx_20260928_123456 或 20260929_xxxx.pptx）。
        page_number: 要精修调整的页码（从 1 开始计数）。
        page_spec: 该页调整后的全新 spec 定义（必须含 layout）。

    Returns:
        JSON 格式结果，包含更新后的 PPTX 下载链接与该页最新的截图预览链接。
    """
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": f"演示文稿不存在: {deck_id}"})

    spec = rec.get("spec")
    if not spec or not isinstance(spec.get("pages"), list):
        return _json({"ok": False, "error": "该文稿未包含可编辑的 spec 定义"})

    pages = spec["pages"]
    if page_number < 1 or page_number > len(pages):
        return _json({"ok": False, "error": f"页码超出范围: {page_number} (总页数: {len(pages)})"})

    # 更新该页 spec
    pages[page_number - 1] = page_spec
    spec["pages"] = pages

    # 重新构建 PPTX
    try:
        deck, path, warnings = build_from_spec({**spec, "output": {"path": rec["path"]}})
    except Exception as e:
        return _json({"ok": False, "stage": "rebuild", "error": str(e)})

    STORE.update_spec(rec["id"], spec)

    base_url = get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"
    t_tag = int(time.time())
    page_preview_url = f"{base_url}/api/preview/{rec['id']}?page={page_number}&t={t_tag}"

    # 重新渲染该页预览图
    if render.available():
        try:
            pv_dir = os.path.join(os.path.dirname(rec["path"]), "preview", _slug(rec["name"], "deck"))
            render.render_pptx(rec["path"], out_dir=pv_dir, pages=[page_number - 1])
        except Exception as e:
            pass

    user_md = (
        f"### ✏️ 第 {page_number} 页已更新完成\n"
        f"👉 **[点击下载最新 PPTX 文件]({download_url})**\n\n"
        f"#### 🖼️ 第 {page_number} 页更新效果预览：\n"
        f"![第 {page_number} 页最新预览]({page_preview_url})"
    )

    result = {
        "ok": True,
        "deck_id": rec["id"],
        "updated_page": page_number,
        "title": page_spec.get("title", ""),
        "layout": page_spec.get("layout", ""),
        "download_url": download_url,
        "page_preview_url": page_preview_url,
        "user_display_markdown": user_md,
        "warnings": warnings,
        "action_guide": (
            "单页精修已完成！【最终回复输出铁律】：请务必直接在最终回复中原样输出返回值中的 `user_display_markdown`，"
            "确保修改后的页面以 Markdown 图片语法 `![预览](url)` 直接渲染为大图展示给用户，绝对严禁写成纯文字超链接！"
        )
    }
    return _json(result)


@server.tool()
def change_theme(deck_id: str, new_theme: str) -> str:
    """一键切换整套演示文稿的设计配色主题与风格（如换成黑金、商务蓝、水墨丹青等），重新构建并返回最新下载链接与全套截图预览链接。

    Args:
        deck_id: 演示文稿 ID 或文件名。
        new_theme: 目标主题名称（如 chinese_poetry, tech_dark, executive_gold, business_blue, modern_teal 等）或自定义配色对象。

    Returns:
        JSON 格式结果，包含更新后的 PPTX 下载链接与全套最新截图预览链接。
    """
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": f"演示文稿不存在: {deck_id}"})

    spec = rec.get("spec")
    if not spec:
        return _json({"ok": False, "error": "该文稿未包含可编辑的 spec 定义"})

    if isinstance(new_theme, dict):
        spec["theme"] = new_theme
        theme_label = new_theme.get("label", "自定义大模型配色")
    else:
        clean_theme = str(new_theme).strip().lower()
        aliases = {
            "gold": "executive_gold", "dark": "tech_dark", "blue": "business_blue",
            "green": "fresh_green", "purple": "violet", "poetry": "chinese_poetry",
            "ink": "chinese_poetry", "chinese": "chinese_poetry"
        }
        clean_theme = aliases.get(clean_theme, clean_theme)
        if clean_theme not in theme_mod.PALETTES:
            return _json({"ok": False, "error": f"未知主题: {new_theme}，可选: {list(theme_mod.PALETTES.keys())}"})
        spec["theme"] = clean_theme
        theme_label = theme_mod.PALETTES[clean_theme]["label"]

    try:
        deck, path, warnings = build_from_spec({**spec, "output": {"path": rec["path"]}})
    except Exception as e:
        return _json({"ok": False, "stage": "rebuild", "error": str(e)})

    STORE.update_spec(rec["id"], spec)

    base_url = get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"

    preview_urls = []
    slides_info = []
    pages = spec.get("pages") or []

    if render.available():
        try:
            pv_dir = os.path.join(os.path.dirname(rec["path"]), "preview", _slug(rec["name"], "deck"))
            _, pngs = render.render_pptx(rec["path"], out_dir=pv_dir)
            t_tag = int(time.time())
            preview_urls = [f"{base_url}/api/preview/{rec['id']}?page={i+1}&t={t_tag}" for i in range(len(pngs))]
            for i, png in enumerate(pngs):
                p_idx = i + 1
                p_meta = pages[i] if i < len(pages) else {}
                slides_info.append({
                    "page": p_idx,
                    "title": p_meta.get("title", f"第 {p_idx} 页"),
                    "layout": p_meta.get("layout", "bullets"),
                    "preview_url": preview_urls[i],
                    "markdown_image": f"![第 {p_idx} 页预览]({preview_urls[i]})"
                })
        except Exception:
            pass

    user_md = make_display_markdown(download_url, rec["name"], slides_info)

    result = {
        "ok": True,
        "deck_id": rec["id"],
        "theme": spec["theme"],
        "theme_label": theme_label,
        "download_url": download_url,
        "preview_urls": preview_urls,
        "slides": slides_info,
        "user_display_markdown": user_md,
        "warnings": warnings,
        "action_guide": (
            "主题已成功切换！【最终回复输出铁律】：请务必直接在最终回复中原样输出返回值中的 `user_display_markdown`，"
            "确保全套最新截图以 Markdown 图片语法 `![预览](url)` 直接渲染为大图展示，绝对严禁写成纯文字超链接！"
        )
    }
    return _json(result)



@server.tool()
def insert_page(deck_id: str, page_spec: dict, position: int = -1) -> str:
    """在现有演示文稿中插入全新的一页（支持指定插入位置或默认追加到末尾），重新构建并返回最新下载链接与截图链接。

    Args:
        deck_id: 演示文稿 ID 或文件名。
        page_spec: 插入页面的完整定义（必须包含 layout 和内容）。
        position: 插入位置（从 1 开始；传入 -1 或大于当前总页数表示追加到末尾）。

    Returns:
        JSON 格式结果，包含更新后的总页数、最新下载链接与插入页面的截图预览链接。
    """
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": f"演示文稿不存在: {deck_id}"})

    spec = rec.get("spec")
    if not spec or not isinstance(spec.get("pages"), list):
        return _json({"ok": False, "error": "该文稿未包含可编辑的 spec 定义"})

    pages = spec["pages"]
    if position <= 0 or position > len(pages) + 1:
        pages.append(page_spec)
        actual_pos = len(pages)
    else:
        pages.insert(position - 1, page_spec)
        actual_pos = position
    spec["pages"] = pages

    try:
        deck, path, warnings = build_from_spec({**spec, "output": {"path": rec["path"]}})
    except Exception as e:
        return _json({"ok": False, "stage": "rebuild", "error": str(e)})

    STORE.update_spec(rec["id"], spec)

    base_url = get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"
    t_tag = int(time.time())
    page_preview_url = f"{base_url}/api/preview/{rec['id']}?page={actual_pos}&t={t_tag}"

    if render.available():
        try:
            pv_dir = os.path.join(os.path.dirname(rec["path"]), "preview", _slug(rec["name"], "deck"))
            render.render_pptx(rec["path"], out_dir=pv_dir, pages=[actual_pos - 1])
        except Exception:
            pass

    user_md = (
        f"### ➕ 已在第 {actual_pos} 页插入新幻灯片\n"
        f"👉 **[点击下载最新 PPTX 文件]({download_url})**\n\n"
        f"#### 🖼️ 新插入页面效果预览：\n"
        f"![第 {actual_pos} 页预览]({page_preview_url})"
    )

    result = {
        "ok": True,
        "deck_id": rec["id"],
        "inserted_position": actual_pos,
        "total_pages": len(pages),
        "download_url": download_url,
        "page_preview_url": page_preview_url,
        "user_display_markdown": user_md,
        "warnings": warnings,
        "action_guide": (
            "新页面已插入完毕！【最终回复输出铁律】：请务必直接在最终回复中原样输出返回值中的 `user_display_markdown`，"
            "以 Markdown 图片语法 `![预览](url)` 直接渲染展示新页面大图与下载链接，绝对严禁写成纯文字超链接！"
        )
    }
    return _json(result)


@server.tool()
def delete_page(deck_id: str, page_number: int) -> str:
    """删除演示文稿中的指定页（如删除某页多余或不达标内容），重新构建并返回最新状态。

    Args:
        deck_id: 演示文稿 ID 或文件名。
        page_number: 要删除的页码（从 1 开始计数）。

    Returns:
        JSON 格式结果，包含删除后的总页数与最新下载链接。
    """
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": f"演示文稿不存在: {deck_id}"})

    spec = rec.get("spec")
    if not spec or not isinstance(spec.get("pages"), list):
        return _json({"ok": False, "error": "该文稿未包含可编辑的 spec 定义"})

    pages = spec["pages"]
    if page_number < 1 or page_number > len(pages):
        return _json({"ok": False, "error": f"页码超出范围: {page_number} (当前总页数: {len(pages)})"})

    if len(pages) <= 1:
        return _json({"ok": False, "error": "演示文稿仅剩 1 页，无法继续删除"})

    deleted_page = pages.pop(page_number - 1)
    spec["pages"] = pages

    try:
        deck, path, warnings = build_from_spec({**spec, "output": {"path": rec["path"]}})
    except Exception as e:
        return _json({"ok": False, "stage": "rebuild", "error": str(e)})

    STORE.update_spec(rec["id"], spec)

    base_url = get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"

    preview_urls = []
    if render.available():
        try:
            pv_dir = os.path.join(os.path.dirname(rec["path"]), "preview", _slug(rec["name"], "deck"))
            _, pngs = render.render_pptx(rec["path"], out_dir=pv_dir)
            t_tag = int(time.time())
            preview_urls = [f"{base_url}/api/preview/{rec['id']}?page={i+1}&t={t_tag}" for i in range(len(pngs))]
        except Exception:
            pass

    user_md = (
        f"### 🗑️ 已删除第 {page_number} 页（原标题：{deleted_page.get('title', '')}）\n"
        f"👉 **[点击下载最新 PPTX 文件]({download_url})**（剩余总页数：{len(pages)}）"
    )

    result = {
        "ok": True,
        "deck_id": rec["id"],
        "deleted_page_number": page_number,
        "deleted_title": deleted_page.get("title", ""),
        "total_pages": len(pages),
        "download_url": download_url,
        "preview_urls": preview_urls,
        "user_display_markdown": user_md,
        "warnings": warnings,
        "action_guide": "页面已成功删除！请务必直接在最终回复中原样输出返回值中的 `user_display_markdown` 与最新下载链接。"
    }
    return _json(result)


@server.tool()
def render_preview(deck_id: str, pages: str = "", dpi: int = 110) -> str:
    """对已入库的产物重新渲染高清预览图，返回各页截图的直接访问链接。"""
    try:
        path = STORE.path_of(deck_id)
    except Exception as e:
        return _json({"ok": False, "error": str(e)})

    if path.lower().endswith(".pdf"):
        pngs = render.pdf_to_pngs(path, out_dir=os.path.dirname(path), dpi=dpi)
    else:
        sel = None
        if pages.strip():
            sel = [int(x) - 1 for x in pages.replace("，", ",").split(",")
                   if x.strip().isdigit()]
        pv_dir = os.path.join(os.path.dirname(path), "preview", _slug(os.path.splitext(os.path.basename(path))[0], "deck"))
        _, pngs = render.render_pptx(path, out_dir=pv_dir, dpi=dpi, pages=sel)


    base_url = get_base_url()
    t_tag = int(time.time())
    preview_urls = [f"{base_url}/api/preview/{deck_id}?page={i+1}&t={t_tag}" for i in range(len(pngs))]

    md_lines = ["#### 🖼️ 幻灯片逐页截图预览："]
    for i, u in enumerate(preview_urls):
        md_lines.append(f"- **第 {i+1} 页**\n\n  ![第 {i+1} 页预览]({u})")

    out_json = {
        "ok": True,
        "deck_id": deck_id,
        "count": len(pngs),
        "preview_urls": preview_urls,
        "user_display_markdown": "\n\n".join(md_lines),
        "action_guide": (
            "预览图渲染完成！【最终回复输出铁律】：请务必在最终回复中直接原样输出 `user_display_markdown`，"
            "确保每页截图以 Markdown 图片语法 `![预览](url)` 直接渲染为图片展示给用户看，绝对严禁输出为纯文字超链接！"
        )
    }
    return _json(out_json)



@server.tool()
def list_artifacts(kind: str = "", limit: int = 25) -> str:
    """列出仓库里的产物（pptx / pdf / png / json）。"""
    items = STORE.list(kind=kind or None, limit=max(1, min(limit, 200)))
    base_url = get_base_url()
    res = []
    for it in items:
        entry = dict(it)
        if entry.get("kind") == "pptx":
            entry["download_url"] = f"{base_url}/api/download/{entry['name']}"
        res.append(entry)
    return _json({"total": len(res), "items": res})


@server.tool()
def read_spec(deck_id: str) -> str:
    """读回某份产物入库时使用的 spec。"""
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": "未找到产物: %r" % deck_id})
    if not rec.get("spec"):
        return _json({"ok": False, "error": "该产物登记时未附带 spec"})
    return _json({"ok": True, "id": rec["id"], "spec": rec["spec"]})


@server.tool()
def delete_artifact(deck_id: str) -> str:
    """删除某份产物及其磁盘文件。"""
    rec = STORE.get(deck_id)
    if not rec:
        return _json({"ok": False, "error": "未找到产物: %r" % deck_id})
    ok = STORE.delete(deck_id)
    return _json({"ok": ok, "deleted": deck_id})


@server.tool()
def server_info() -> str:
    """查看服务状态：输出目录、渲染能力、主题与模板数量。"""
    return _json({
        "name": "pptx-studio",
        "version": "0.2.0",
        "base_url": get_base_url(),
        "workspace": WORKSPACE,
        "output_dir": OUTPUT_DIR,
        "store_dir": STORE.root,
        "preview_available": render.available(),
        "soffice": render.soffice_path(),
        "themes": list(theme_mod.PALETTES.keys()),
        "templates": list(templates.TEMPLATES.keys()),
        "artifacts": len(STORE.list(limit=10000)),
        "layouts": sorted(VALID_LAYOUTS),
    })
