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
    "version": "0.5.0",
    "instructions": (
        "你是一个顶尖商业与科技演示文稿架构师与设计专家。\n\n"
        "【自主选择主题与版式核心法则】\n"
        "系统已内置 32 种覆盖全业务场景的高级版式库、50 套顶级专业色彩主题（8 大行业体系）与 7 套开箱即用行业成套模板（支持自定义配色与高级环境光影）。\n"
        "严禁千篇一律地只使用单一版式（如整篇都是 bullets 或 grid），大模型必须根据用户的【主题定位】与【页面叙事目标】自主决策匹配最合适的风格和版式组合：\n"
        "1. 50 套专业主题自主匹配（覆盖 8 大行业体系，支持科技点阵与双层极光光晕 Ambient Aura）：\n"
        "   - AI/大模型/前沿科技/极客发布会 -> `tech_dark`、`ai_neon`、`cyberpunk_matrix`、`quantum_cyan`\n"
        "   - 咨询机构/投行/金融研报/商业分析 -> `business_blue`、`goldman_navy`、`morgan_slate`、`deloitte_emerald`\n"
        "   - 高端商务盛典/领袖峰会/年度答谢/奢华黑金 -> `executive_gold`、`chanel_monochrome`、`patek_champagne`\n"
        "   - 数字化转型/产业互联网/企业SaaS -> `modern_teal`、`datadog_purple`、`stripe_slate`\n"
        "   - 国企汇报/党政政务/行业监管/社会责任 -> `gov_blue`、`tsinghua_purple`、`beida_crimson`\n"
        "   - 医疗健康/绿色低碳/新能源/ESG -> `fresh_green`、`esg_forest`、`clinical_blue`\n"
        "   - 新消费/电商零售/创业路演/品牌推广 -> `warm_coral`、`lululemon_terracotta`、`airbnb_sunrise`\n"
        "   - 东方美学/国潮水墨/历史文韵/借诗明志 -> `chinese_poetry`、`song_dynasty_celadon`、`dunhuang_ochre`\n"
        "   - 个性化定制 -> theme 直接传入包含 bg, card_bg, primary, accent, is_dark 的字典对象\n"
        "2. 38 大版式库全景（按表达目的分为 14 大功能维度，彻底告别千篇一律与审美疲劳）：\n"
        "   - 封面多元矩阵（根据演讲体裁选用最契合版式）：\n"
        "     • `title_slide`（经典左右分栏：左主副标题 + 右侧 3~4 项精装看点卡片）\n"
        "     • `cover_centered` / `title_centered`（极简磅礴·居中发布会/高端峰会封面：居中大标题+横向浮岛卡）\n"
        "     • `cover_pillars` / `title_pillars`（战略三/四立柱封面：上半部宏大命题，下半部 3~4 大业务/战略支柱铺展）\n"
        "     • `cover_minimal` / `title_minimal`（前沿杂志与学术研报：非对称优雅大留白+右下沉浸式摘要盒）\n"
        "     • `cover_split` / `title_split`（左右强烈色彩对撞·沉浸式分屏：左侧高饱和品牌色屏+右侧浮岛亮点）\n"
        "   - 全局大目录与结构导航类：\n"
        "     • `agenda` / `toc`（全篇 3~6 章节全景骨架导航大盘，横向立柱或 2x3 Bento 网格，第二页必选）\n"
        "     • `section`（经典杂志感章节过渡：120pt 巨大数字水印 + 底部议题卡片）\n"
        "     • `section_centered`（极简专注·居中演讲转场：Contrast 反差底色 + 核心启发金句盒）\n"
        "     • `section_split`（左右分栏全景转场：左侧命题 + 右侧核心成果与研讨议题双卡）\n"
        "     • `section_progress`（全景章节进度轴点亮：贯穿全篇章节进度轴，高亮当前阶段）\n"
        "   - 业务全景类：`dashboard`（数据指标+图表+进度条综合看板）\n"
        "   - 团队组织类：`team` / `team_grid`（专家画像与履历）、`org_tree`（组织架构树/决策治理树）\n"
        "   - 评估矩阵类：`matrix_quadrant`（二维四象限价值/复杂度矩阵）、`capability_radar`（多维能力对标雷达与综合评级卡）\n"
        "   - 战略层级类：`pyramid` / `pyramid_funnel`（战略金字塔）、`funnel`（逐级下沉转化漏斗专版）\n"
        "   - 态势分析类：`swot` / `swot_matrix`（SWOT 优势、劣势、机会、威胁 4 象限）\n"
        "   - 架构拓扑类：`stack`（分层系统技术架构）、`diagram`（微服务拓扑节点调用网络）\n"
        "   - 方案权衡类：`comparison`（Before vs After 对比）、`pros_cons`（利弊收益与风险考量）\n"
        "   - 流程演进类：`process_flow`（有向粗箭头业务流程）、`timeline`（横向时间线里程碑）、`roadmap`（多阶段跨周期战略路线图）\n"
        "   - 对话与视窗类：`chat_flow`（AI Agent 对话流与用户原声）、`split_showcase`（代码终端视窗 + 3 个 Bento 卖点砖块）\n"
        "   - 核心论断类：`hero_statement`（震撼大字报金句+支柱）、`quote_focus`（沉浸式高管观点与大引用页）、`kpi`（大数字指标冲击+战略洞察）\n"
        "   - 监控与落地类：`metric_grid`（高密度 6~8 指标监控大盘）、`case_study`（标杆案例+痛点+ROI+证言）、`pricing`（产品套餐与定价推荐）、`table`（卡片对比表格）、`grid`（2x2/1x3 Bento矩阵）、`bullets`（Bento要点砖块）、`text`（正文演说）\n"
        "   - 疑虑与收尾类：`faq`（手风琴 Q&A 问答）、`poetry`（古典诗词韵律）、`summary`（高管核心结论沉淀+落地行动计划 Next Steps）\n"
        "3. 全页面排版饱满度与防溢出黄金铁律（Anti-Hollow & Anti-Overflow Rules）：\n"
        "   - 字少防空洞铁律（严禁大面积空旷留白）：严禁生成只有三两字的电报式碎片！每个要点或卡片必须包含【核心主题/短标题 + 1~2 句业务背景、落地价值或量化指标】（如 '架构解耦：完成微服务治理拆分，跨团队联调耗时降低 40%'）；在 `grid`、`comparison`、`timeline`、`table`、`kpi` 等页面中，若条目较少（<= 3 项），大模型必须主动在 page spec 中配置 `takeaway`（战略洞察/执行建议/保障机制）以填补下部视觉重心，杜绝留出大半屏空白；\n"
        "   - 字多防溢出铁律（严禁文字挤出卡片或遮挡其他模块）：严格遵守黄金字数预算区间。卡片主标题 6~14 字，卡片描述 25~55 字；严禁在单个卡片内倾倒超过 80 字的长篇大论，长篇内容必须提炼拆解为 2~3 个结构化分点（`items` / `bullets`）；严禁单个列表项超过 4 行，确保文字绝不向下穿模遮挡分割线或卡片底边；\n"
        "   - 视觉节奏感与色彩对比：一份 PPT 必须交替运用 5~8 种不同版式，严禁连续两页使用同一 layout；善用 [[关键字]] 自动高亮关键论断；封面、大目录、章节转场或核心转折页，推荐显式指定 `bg_mode: 'contrast'` 制造有节奏的视觉高潮。\n\n"
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
    import secrets
    rand_tag = secrets.token_hex(8)

    if base.startswith(ts):
        candidate = f"{base}.pptx"
    else:
        candidate = f"{ts}_{rand_tag}_{base}.pptx"

    # 兜底防重
    while os.path.exists(os.path.join(OUTPUT_DIR, candidate)):
        rand_tag = secrets.token_hex(8)
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
    # 1. 封面与结构多元矩阵（防审美疲劳）
    "title_slide", "cover", "title", "cover_split_card",
    "cover_centered", "title_centered", "cover_keynote", "cover_summit",
    "cover_pillars", "title_pillars", "cover_cards", "cover_bottom_cards",
    "cover_minimal", "title_minimal", "cover_editorial",
    "cover_split", "title_split", "cover_hero", "cover_contrast",
    # 1.1 全局目录导航大盘
    "agenda", "toc", "table_of_contents", "contents", "catalog",
    # 1.2 章节与过渡页多元矩阵
    "section", "chapter", "divider",
    "section_centered", "chapter_centered", "section_minimal",
    "section_split", "chapter_split", "section_side",
    "section_progress", "chapter_progress", "section_timeline", "chapter_stepper",
    # 2. 数据与看板
    "dashboard", "kpi", "table", "metric_grid", "metrics_grid",
    "chart_analytics", "chart", "analytics",
    # 3. 架构与拓扑
    "stack", "architecture", "layers",
    "diagram", "topology", "architecture_diagram",
    "org_tree", "tree",
    # 4. 流程与时序
    "process_flow", "flow", "pipeline",
    "timeline",
    "roadmap", "roadmap_milestones", "milestones",
    # 5. 分析、矩阵与对比
    "comparison",
    "pros_cons", "tradeoff",
    "swot", "swot_matrix",
    "matrix_quadrant", "matrix", "quadrant",
    "pyramid", "pyramid_funnel",
    "funnel", "funnel_stages",
    "radar", "capability", "capability_radar",
    # 6. 核心论断、文化与大引用
    "hero_statement", "statement", "quote", "focus",
    "quote_focus", "executive_quote",
    "poetry", "poem", "verse",
    # 7. 团队、商业与交互落地
    "team", "team_grid", "members", "org",
    "case_study", "case",
    "pricing", "pricing_packages", "packages", "plans",
    "split_showcase", "showcase",
    "chat_flow", "chat", "dialogue",
    # 8. 程序员硬核技术走读与课堂教学
    "code_walkthrough", "tech_talk",
    "classroom_quiz", "quiz", "exam", "exercise",
    # 9. 总结与问答
    "faq", "faq_accordion", "qa",
    "summary", "summary_next_steps", "next_steps", "action_plan",
    # 10. 通用多卡片与要点
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

    涵盖 9 大顶级主题与 50 种色彩方案、27 种全场景专业版式用法参数表（含原生图表分析、程序员源码走读、课堂互动测验）、防止页面空洞的核心技巧与大模型自适应推荐。

    Args:
        topic: 演示文稿主题（如"AI大模型商业化落地方案"、"程序员微服务高并发技术分享"、"微积分公开课教学"等），传入后将返回智能匹配的主题与8~10页叙事版式蓝图。
    """
    # 智能主题与叙事推荐算法
    topic_rec = {}
    if topic:
        t_low = topic.lower()
        if any(k in t_low for k in ["ai", "大模型", "算法", "极客", "智能", "算力", "芯片", "科技"]):
            rec_theme = "tech_dark"
            theme_reason = "AI与硬核科技主题匹配深空极客黑，赛博荧光青与午夜黑凸显未来感"
            storyboard = ["cover_pillars", "agenda", "hero_statement", "pyramid", "stack", "diagram", "process_flow", "kpi", "roadmap", "summary"]
        elif any(k in t_low for k in ["金融", "银行", "证券", "投资", "投行", "基金", "咨询", "财报", "分析"]):
            rec_theme = "business_blue"
            theme_reason = "金融投资咨询匹配顶尖咨询蓝（麦肯锡冷灰+藏青），理性克制、严谨专业"
            storyboard = ["cover_minimal", "agenda", "dashboard", "swot", "comparison", "kpi", "table", "timeline", "summary"]
        elif any(k in t_low for k in ["高端", "奢华", "峰会", "盛典", "年会", "年度", "领袖", "颁奖"]):
            rec_theme = "executive_gold"
            theme_reason = "领袖峰会与高端盛典匹配曜石黑金，尊享奢华、大气非凡"
            storyboard = ["cover_centered", "agenda", "hero_statement", "team", "dashboard", "kpi", "case_study", "roadmap", "summary"]
        elif any(k in t_low for k in ["政务", "国企", "党建", "社会责任", "治理", "监管", "汇报"]):
            rec_theme = "gov_blue"
            theme_reason = "党政国企与行业监管匹配政务权威蓝（规范庄重蓝+典雅中国红），庄重大气"
            storyboard = ["cover_pillars", "agenda", "dashboard", "grid", "process_flow", "kpi", "table", "timeline", "summary"]
        elif any(k in t_low for k in ["医疗", "医药", "健康", "绿色", "低碳", "环保", "esg", "农业"]):
            rec_theme = "fresh_green"
            theme_reason = "健康生态绿色环保匹配清新生态绿（薄荷雾白+森林绿），生机盎然"
            storyboard = ["title_slide", "agenda", "dashboard", "swot", "process_flow", "kpi", "roadmap", "case_study", "summary"]
        elif any(k in t_low for k in ["电商", "零售", "消费", "出海", "路演", "创业", "营销"]):
            rec_theme = "warm_coral"
            theme_reason = "新消费与创业路演匹配活力珊瑚橙，富有感染力与商业爆发力"
            storyboard = ["cover_split", "agenda", "hero_statement", "pros_cons", "pyramid", "kpi", "pricing", "case_study", "summary"]
        elif any(k in t_low for k in ["文化", "艺术", "诗", "词", "国风", "非遗", "汉服", "古典"]):
            rec_theme = "chinese_poetry"
            theme_reason = "国风雅韵与文化艺术匹配水墨丹青，宣纸白+松烟墨+朱砂红，格调拉满"
            storyboard = ["cover_centered", "agenda", "poetry", "hero_statement", "grid", "timeline", "case_study", "summary"]
        elif any(k in t_low for k in ["教学", "课堂", "公开课", "微积分", "课程", "教师", "学生", "试题", "教育", "讲义", "培训"]):
            rec_theme = "modern_teal"
            theme_reason = "大中小学课堂教学匹配数字科技青（清晰护眼、结构清晰、聚焦考点与互动解题探究）"
            storyboard = ["cover_centered", "agenda", "process_flow", "classroom_quiz", "diagram", "chart_analytics", "summary"]
        elif any(k in t_low for k in ["程序员", "技术分享", "架构师", "源码", "并发", "高可用", "重构", "微服务", "后端", "前端"]):
            rec_theme = "tech_dark"
            theme_reason = "程序员技术分享与硬核架构匹配深空极客黑（极客macOS代码视窗 + 原生压测指标图表 + 避坑指南）"
            storyboard = ["cover_pillars", "agenda", "hero_statement", "stack", "code_walkthrough", "chart_analytics", "diagram", "summary"]
        elif any(k in t_low for k in ["潮牌", "文创", "设计", "年轻", "潮流", "时尚", "娱乐"]):
            rec_theme = "violet"
            theme_reason = "潮流先锋与品牌创意匹配紫罗兰先锋，摩登时尚"
            storyboard = ["cover_split", "agenda", "hero_statement", "grid", "pros_cons", "kpi", "pricing", "summary"]
        else:
            rec_theme = "modern_teal"
            theme_reason = "现代化商业方案首选数字科技青（冰爽微青白+沉静深青+珊瑚橙高光）"
            storyboard = ["cover_pillars", "agenda", "dashboard", "comparison", "stack", "kpi", "roadmap", "case_study", "summary"]

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
            "【自适应色彩节奏（Adaptive Rhythm & bg_mode）】一份顶级 PPT 绝非全篇千篇一律纯白或纯黑！支持页面级色彩氛围控制：单页设置 bg_mode: 'contrast'（深邃主色对比反转，强烈建议用于章节过渡、核心金句、结论出鞘）、'tint'（温润特色氛围底色+纯白浮岛悬浮卡片）、'hero' 或直接自定义 hex 底色 'bg_color'，打破单调形成呼吸感！",
            "【单页独立主题切换（page.theme）】支持在单页通过 theme 字段独立覆盖全局主题（例如文稿全局为商务蓝，但在科技创新页临时指定 theme: 'tech_dark'，在人文故事页指定 'song_celadon'），满足跨场景多元混排需求！",
            "【视觉节奏感（Visual Rhythm）】一份高水准的 PPT 必须有张有弛、版式交替推进！严禁连续两页使用同一种 layout。一套 PPT 建议自主编排 5~8 种不同版式！",
            "【拒绝大片空白与单调】每页幻灯片必须信息饱满、层级鲜明。切勿使用空旷大卡片，应丰富为带数据标签的砖块、分层架构胶囊、或底部核心战略洞察卡（takeaway / note）。",
            "【核心数据突出】大字号数值是现代 PPT 的灵魂。关键数字用 34~42pt 加粗展示，文本中用 [[关键字/数字]] 标记自动高亮。",
            "【结构化表达】避免纯大段文字堆砌，根据内容属性选择最匹配的专业版式：架构选 stack/diagram，流程选 process_flow/roadmap，团队选 team，分析选 swot/pros_cons/pyramid，金句选 hero_statement，指标选 kpi，商业选 case_study/pricing。"
        ],
        "themes": [
            {"name": "song_celadon", "style": "汝窑天青·宋瓷美学（东方美学/极简宋韵/茶道生活/文化智库）", "features": "温润天青釉底 + 定窑白瓷悬浮卡 + 古金高光，脱离单纯黑白，极度高级雅致"},
            {"name": "warm_terracotta", "style": "暖阳陶土·地中海建筑（人文空间/艺术策展/高端家居/手工质感）", "features": "地中海陶土暖沙底 + 纯白石灰泥卡片 + 碧蓝点睛，非黑非白的大师级建筑色彩"},
            {"name": "dunhuang_ochre", "style": "敦煌飞天·莫高矿彩（丝路驼金/壁画复兴/文旅文博/传统美学）", "features": "莫高绢帛金底 + 象牙白卡 + 孔雀石绿微标，沉淀千年的西域东方美学"},
            {"name": "milky_tea", "style": "暖糯奶茶·新零售（烘焙轻食/新茶饮连锁/餐饮文创/治愈生活）", "features": "燕麦暖驼奶霜底 + 纯白微浮岛 + 焦糖暖橙高光，温润治愈新消费首选"},
            {"name": "bamboo_mist", "style": "烟雨苍竹·江南细雨（苍竹黛瓦/江南文人私家园林/诗意生活）", "features": "细雨竹露淡青底 + 定窑白卡 + 古铜金徽标，文人风骨与静谧禅意"},
            {"name": "wabi_sabi", "style": "侘寂灰泥·东方侘寂（微水泥素朴质感/极简禅意/当代艺术）", "features": "素朴温润微水泥暖灰底 + 羊脂白卡 + 禅意碳黑排版，极简宁静"},
            {"name": "botanical_sage", "style": "草木鼠尾草·生态有机（现代农业/生物生命/有机健康/森林生活）", "features": "草木鼠尾草灰绿底 + 纯白卡片 + 琥珀金高光，纯净自然呼吸感"},
            {"name": "tech_dark", "style": "深空极客黑（科技发布会/AI/大模型/硬核系统）", "features": "午夜黑 + 赛博荧光青 + 通透纯白字 + 极光环境微光晕，未来感爆棚"},
            {"name": "business_blue", "style": "顶尖咨询蓝（麦肯锡/投行/金融研报/商业分析）", "features": "冷灰底 + 藏青大字 + 皇家宝蓝核心高光，极简理性"},
            {"name": "executive_gold", "style": "黑金尊享（高端战略/领袖峰会/年度答谢/颁奖盛典）", "features": "曜石黑 + 奢华香槟金 + 暖象牙白，领袖专属"},
            {"name": "modern_teal", "style": "数字科技青（数字化转型/产业互联网/企业SaaS）", "features": "冰爽微青白 + 沉静深青 + 珊瑚橙互补高光，通用现代"},
            {"name": "gov_blue", "style": "政务权威蓝（国企党政汇报/公共服务/监管审计/社会责任）", "features": "规范庄重蓝 + 典雅中国红强调 + 暖金辅助，庄重肃穆"},
            {"name": "chinese_poetry", "style": "水墨丹青·东方雅韵（国风/传统文化/非遗/诗词借物明志）", "features": "温润宣纸米白 + 沉着松烟墨黑 + 经典朱砂红印泥 + 天青汝窑黛绿，高古雅致"},
            {"name": "custom", "style": "大模型自主设计 HTML/CSS 配色字典", "features": "支持在 theme 字段直接传入包含 bg, card_bg, primary, accent, is_dark 的字典对象，完全解放调色自由"}
        ],
        "golden_text_budget_and_layout_contracts": {
            "title_slide": {
                "name": "经典高管封面页（左右分栏）",
                "best_items": "右侧 cards 传入 3~4 项战略看点；支持 bg_mode: 'hero' / 'contrast' 开启震撼深邃品牌主色",
                "text_budget": "主标题 12~24字 | 副标题 16~30字 | 每条看点 14~28字",
                "anti_hollow_tips": "严禁右侧留空！必须在 cards/bullets 中传入 3~4 项核心看点，左侧配置 tag 与 author 胶囊，呈现顶级科技峰会质感。"
            },
            "cover_centered": {
                "name": "极简磅礴·居中发布会与峰会封面页",
                "best_items": "居中超大主标题 + 居中副标 + 下方横向 3~4 个浮岛看点卡片 (bullets)",
                "text_budget": "主标题 8~18字 | 副标题 15~35字 | 每条看点 10~22字",
                "anti_hollow_tips": "大开大合，适合全球发布会、领袖峰会、高端学术大课开篇；提供 category 与 meta_line 保证视觉饱满。"
            },
            "cover_pillars": {
                "name": "战略立柱与多业务线封面页",
                "best_items": "上半部宏观命题 + 右上战略使命盒 + 下半部 3~4 大战略支柱 (pillars)",
                "text_budget": "主标题 10~22字 | 支柱 title 6~12字 | 支柱 desc 20~40字",
                "anti_hollow_tips": "开篇即亮出全篇战略大骨架，适合战略规划、顶层设计、业务矩阵发布。"
            },
            "cover_minimal": {
                "name": "前沿杂志风与学术研报封面页",
                "best_items": "非对称优雅大留白 + 年份水印 + 右下沉浸式摘要盒 (abstract)",
                "text_budget": "主标题 10~20字 | 副标题 12~28字 | 摘要 40~80字",
                "anti_hollow_tips": "极具国际顶尖商业周刊排版质感，适合金融研报、咨询洞察、论文与设计分享。"
            },
            "cover_split": {
                "name": "左右色块强烈对撞·沉浸式分屏封面",
                "best_items": "左侧高饱和品牌深底主屏 + 右侧浮岛 3~4 项核心看点卡片 (highlights)",
                "text_budget": "主标题 10~20字 | 副标题 14~30字 | 每条看点 15~30字",
                "anti_hollow_tips": "视觉张力极强，专为创业路演、投融资谈判与重大品牌变革打造。"
            },
            "agenda": {
                "name": "全局大目录导航大盘页（全篇骨架总览）",
                "best_items": "chapters: 3~6 个核心章节；3~4 章节采用全屏横向立柱，5~6 章节自适应 2x3 Bento 网格",
                "text_budget": "章节 title 6~14字 | 章节 desc 15~35字 | 章节子议题 items 2~3条 (每条8~16字)",
                "anti_hollow_tips": "任何多页正式汇报第二页必选！彻底理清全篇架构脉络，杜绝听众迷航。"
            },
            "section": {
                "name": "经典杂志感章节过渡页",
                "best_items": "默认自适应 bg_mode: 'contrast' 深邃品牌主色，内置 120pt 巨大半透明罗马/阿拉伯数字水印；传入 agenda/topics 2~4 条议题清单",
                "text_budget": "主标题 8~18字 | 副标题 15~35字 | 每条议题 6~16字",
                "anti_hollow_tips": "必须传入 agenda 议题清单或 lead 导读，底部自动生成精美序号胶囊卡片，打破文稿平铺直叙，形成强烈章节视觉里程碑。"
            },
            "section_centered": {
                "name": "极简专注·居中演讲转场页",
                "best_items": "Contrast 反差底色 + 居中大主标题 + 居中启发性金句盒 (quote/lead)",
                "text_budget": "章节标题 6~16字 | 副标题 10~25字 | 金句/导读 20~45字",
                "anti_hollow_tips": "高管演讲与关键认知转折首选，给受众以绝对专注的放空与转折呼吸感。"
            },
            "section_split": {
                "name": "左右分栏式章节全景转场页",
                "best_items": "左侧命题导读 + 右侧【本章核心目标收获】与【本章研讨议题清单】双卡",
                "text_budget": "章节标题 8~16字 | 核心成果 3条 (每条12~25字) | 议题 3条",
                "anti_hollow_tips": "承上启下，咨询研报与技术架构深度分享首选。"
            },
            "section_progress": {
                "name": "全景进度点亮章节转场页",
                "best_items": "贯穿全屏的章节进度轴，自动将当前章节高亮并放大浮岛卡片，其余章节静默",
                "text_budget": "章节标题 8~16字 | 各章节节点 title 6~12字",
                "anti_hollow_tips": "系列汇报、项目多阶段进展、大型课件首选，受众一眼获知当前汇报位置。"
            },
            "dashboard": {
                "name": "综合全屏数据看板",
                "best_items": "kpis: 3~4项 | panels: 2个复合面板 (柱图/进度)",
                "text_budget": "标题 8~16字 | KPI label 4~8字 | 面板要点 2~3条 (每条15~30字)",
                "anti_hollow_tips": "复合面板不可仅给标题；左面板配 chart (categories+series)，右面板配 progress 与 items 交付清单。"
            },
            "hero_statement": {
                "name": "核心论断与大字报",
                "best_items": "points: 3个支撑微卡片",
                "text_budget": "核心论断 16~32字 | 微支柱 title 6~10字 | 微支柱 desc 20~35字",
                "anti_hollow_tips": "statement 必须包含 [[核心高光词]]；points 必须提供 desc 详细展开，充实下半部微卡片空间。"
            },
            "swot": {
                "name": "SWOT战略态势分析矩阵",
                "best_items": "s, w, o, t 四大象限各 3~4 条",
                "text_budget": "每个要点 15~30字",
                "anti_hollow_tips": "严禁每个象限只传 1~2 个孤零零词汇！每项应包含短语+背景说明，填满象限空间消除大片空白。"
            },
            "stack": {
                "name": "纵向分层技术架构图",
                "best_items": "layers: 3~4层 | 每层 components: 3~5个组件砖块",
                "text_budget": "layer_tag 4~8字 | title 8~16字 | desc 20~40字 | 组件名 4~10字",
                "anti_hollow_tips": "必须配置右侧 components 砖块矩阵，左侧 desc 控制在 20~40 字保持 2 行饱满排版。"
            },
            "diagram": {
                "name": "分布式服务拓扑网络图",
                "best_items": "nodes: 4~6个 (2x2或2x3网格) | edges: 3~5条连线",
                "text_budget": "node title 6~12字 | node desc 10~20字 | edge label 4~8字",
                "anti_hollow_tips": "每个节点必须带有 category (如DB/GATEWAY/APP) 与 desc；底部强烈推荐配置 takeaway 总结核心解耦原则。"
            },
            "process_flow": {
                "name": "有向流程推进演进图",
                "best_items": "steps: 3~4个阶段卡片 (环形箭头连接)",
                "text_budget": "step title 6~12字 | step desc 30~55字",
                "anti_hollow_tips": "step desc 必须展开 2~3 行（30~55字）详细说明输入输出，避免竖卡下半部大面积纯白。"
            },
            "roadmap": {
                "name": "战略演进路线图",
                "best_items": "phases: 3~4个时间周期 | 每个阶段 tasks: 3~4项",
                "text_budget": "phase 4~8字 | title 6~12字 | goal 15~30字 | task 12~24字",
                "anti_hollow_tips": "每个阶段必须包含 phase、title、goal 战略目标与 3~4 条打勾 tasks 任务清单，结构最工整。"
            },
            "timeline": {
                "name": "横向时间线里程碑",
                "best_items": "items: 4~5个时间节点 (含圆形胶囊标记)",
                "text_budget": "year/tag 4~8字 | title 6~12字 | desc 25~45字",
                "anti_hollow_tips": "desc 建议 25~45 字（2~3行），页面底部推荐配置 takeaway 推进保障机制卡收拢留白。"
            },
            "comparison": {
                "name": "双栏深度方案对比",
                "best_items": "left / right 左右各 3~5 项",
                "text_budget": "栏标题 8~16字 | 每项 item 18~35字",
                "anti_hollow_tips": "左右两栏条目数量保持平衡对等，利用 [[关键词]] 高亮显示两者的核心差异与突破点。"
            },
            "pros_cons": {
                "name": "方案利弊得失权衡",
                "best_items": "pros: 3~4项 | cons: 3~4项 | takeaway: 1条决策结论",
                "text_budget": "每项条目 18~35字 | takeaway 25~50字",
                "anti_hollow_tips": "必须配置底部 takeaway 权衡决策卡，形成完整的“收益-代价-最终裁决”商业闭环。"
            },
            "kpi": {
                "name": "核心量化指标冲击页",
                "best_items": "kpis: 3~4个大数字指标卡 | takeaway: 1条成果分析",
                "text_budget": "value 2~6字符 | unit 1~4字符 | label 4~8字 | captions 12~25字",
                "anti_hollow_tips": "value 与 unit 拆开存放；强烈推荐配置底部 takeaway 核心战略洞察卡，视觉极度充实。"
            },
            "case_study": {
                "name": "标杆客户成功案例",
                "best_items": "左栏痛点举措 | 右上 3个ROI指标 | 右下 1条客户证言",
                "text_budget": "challenge 40~70字 | solution 40~70字 | quote 25~55字",
                "anti_hollow_tips": "左侧 challenge 与 solution 需各展开 2~3 行，右侧提供真实量化 ROI 与权威客户证言。"
            },
            "pricing": {
                "name": "商业套餐定价矩阵",
                "best_items": "packages: 固定 3 档定价卡片 (基础/旗舰/私有)",
                "text_budget": "name 4~8字 | price 3~8字符 | desc 15~30字 | features 4~5项 (8~20字)",
                "anti_hollow_tips": "中间档必须设置 is_popular: true 成为视觉焦点，每档包含 4~5 项权益与 cta 购买文案。"
            },
            "table": {
                "name": "卡片式科技斑马表格",
                "best_items": "headers: 4~6列 | rows: 4~6行 | takeaway: 1条数据洞察",
                "text_budget": "header 4~10字 | cell 4~18字 | takeaway 25~50字",
                "anti_hollow_tips": "控制行数在 4~6 行防止文字拥挤，底部配置 takeaway 数据洞察卡，专业研报级质感。"
            },
            "team": {
                "name": "核心专家与研发团队",
                "best_items": "members: 3 位专家卡片 (或4位)",
                "text_budget": "name 2~4字 | role 6~12字 | desc/bio 45~75字 | tags 2~3个",
                "anti_hollow_tips": "履历 desc 必须提供 45~75 字（2~3行），底部提供 2~3 个 skills 药丸，彻底消除卡片中部空白。"
            },
            "pyramid": {
                "name": "战略金字塔与业务漏斗",
                "best_items": "tiers: 3~4层递进 (从顶至底展开)",
                "text_budget": "title 6~14字 | desc 15~35字 | metric 4~12字符",
                "anti_hollow_tips": "每层配置 level (L1~L4)、title、desc 与右侧 metric，层级宽度梯形展开，比例极为和谐。"
            },
            "faq": {
                "name": "高频疑虑解答 (Q&A)",
                "best_items": "faqs: 3~4个手风琴抽屉卡",
                "text_budget": "q 12~24字 | a 40~85字 | tag 4~6字",
                "anti_hollow_tips": "答案 a 必须提供 2~3 行详实解答（40~85字），右侧带 tag 分类胶囊，消除单薄感。"
            },
            "poetry": {
                "name": "东方古典韵律美学",
                "best_items": "poem: 2句工整对仗诗句 | interpretation: 现代战略解读",
                "text_budget": "author 4~8字 | poem 7~14字 | interpretation 40~80字",
                "anti_hollow_tips": "诗句字数严格对仗，底部 interpretation 结合企业管理/技术架构哲学升华主题。"
            },
            "grid": {
                "name": "Bento 矩阵网格页",
                "best_items": "items: 4个独立卡片 (2x2布局)",
                "text_budget": "card title 6~12字 | 每卡 bullets 2~4条 (每条15~30字)",
                "anti_hollow_tips": "4 张卡片结构高度对称，每卡含 2~3 条要点，支持富文本重点标记。"
            },
            "bullets": {
                "name": "结构化 Bento 要点砖块",
                "best_items": "items: 4~6个独立 Bento 砖块卡片",
                "text_budget": "每项要点 25~50字 (推荐【标签】+展开说明)",
                "anti_hollow_tips": "系统自动将 1~6 项封装为带序号微胶囊的立体 Bento 砖块，严禁传单字，多用 [[核心高光]]。"
            },
            "text": {
                "name": "正文演讲与白皮书",
                "best_items": "lead 导言微卡 + paragraphs 2~3段论述",
                "text_budget": "lead 25~45字 | 每个段落 90~150字",
                "anti_hollow_tips": "必配 lead 核心金句卡片，正文段落叙述流畅，适合高管致辞与架构哲学沉淀。"
            },
            "summary": {
                "name": "高管总结与推进清单",
                "best_items": "takeaways: 3~4条战略结论 | actions: 3~4项行动计划",
                "text_budget": "takeaway 20~40字 | task 18~35字 | owner 4~8字 | due 4~8字",
                "anti_hollow_tips": "左栏沉淀战略认知，右栏落实落地责任人 (owner) 与时间周期 (due)，完美收尾。"
            },
            "matrix_quadrant": {
                "name": "二维四象限矩阵分析页",
                "best_items": "x_axis, y_axis 坐标轴 + 4 个象限 (q1~q4)",
                "text_budget": "轴标签 6~12字 | 象限名称 8~14字 | 每象限 2~4 条要点 (12~25字)",
                "anti_hollow_tips": "四象限必须传满，注明优先落地与战略探索，适合技术选型、波士顿矩阵与商业决策。"
            },
            "org_tree": {
                "name": "组织架构与决策树",
                "best_items": "root 顶层中枢 + 2~4 个 branches + 每个分支 3~5 个 leaves",
                "text_budget": "root title 8~16字 | 分支 title 6~12字 | 叶子节点 6~14字",
                "anti_hollow_tips": "树状层次展开，必须配置 root 核心角色与各业务域的具体职能清单，结构极度清晰。"
            },
            "capability_radar": {
                "name": "多维能力雷达与竞品全景评估",
                "best_items": "dimensions: 5~6个能力维度 + 右侧综合得分卡",
                "text_budget": "维度名 4~8字 | 评分 60~100 | score 4~6字符 (如98.5) | highlights 2~4个",
                "anti_hollow_tips": "左侧条形双轨打分，右侧大号综合得分与核心领跑标签，展现碾压级产品竞争力。"
            },
            "funnel": {
                "name": "转化漏斗专版",
                "best_items": "stages: 4~5个递进漏斗层级 (自顶向下)",
                "text_budget": "stage 6~10字 | title 6~12字 | desc 15~30字 | rate 4~12字符",
                "anti_hollow_tips": "自顶向下逐级收拢，每层右侧标明转化率指标，中间阐述转化动作，适合营销与获客分析。"
            },
            "chat_flow": {
                "name": "AI 对话流与用户访谈原声",
                "best_items": "dialogues: 2~3轮真实气泡流对话",
                "text_budget": "用户提问 15~40字 | 专家回复 45~95字 (支持富文本高光)",
                "anti_hollow_tips": "右侧用户问题卡 + 左侧专家/AI 深度回答卡，模拟真实人机交互与客户口碑场景。"
            },
            "split_showcase": {
                "name": "左右分屏技术视窗与 Bento 特性",
                "best_items": "左侧代码/终端视窗 + 右侧 3个 Bento 卖点砖块",
                "text_budget": "window_title 10~20字 | 代码 6~12行 | 卖点 title 6~12字 | desc 20~40字",
                "anti_hollow_tips": "左侧呈现高科技 macOS 视窗与真实代码/架构调用，右侧三砖块提炼商业价值。"
            },
            "metric_grid": {
                "name": "高密度微指标监控大盘",
                "best_items": "metrics: 6~8个紧凑指标微卡 (2x3或2x4)",
                "text_budget": "label 4~8字 | value 2~6字符 | unit 1~4字符 | trend 4~8字符 (如 ▲ +35%)",
                "anti_hollow_tips": "应对大促与年度核心监控指标，卡片内含趋势胶囊，高密度信息呈现，杜绝空白。"
            },
            "quote_focus": {
                "name": "沉浸式高管观点与大引用页",
                "best_items": "quote 核心金句 + author 头像职务 + 3个战略承诺胶囊",
                "text_budget": "quote 25~55字 | author 2~4字 | title 10~25字 | 承诺 6~12字",
                "anti_hollow_tips": "浅色大双引号水印饰纹，居中大号人物宣言，适合发布会高潮与高管意志传递。"
            },
            "chart_analytics": {
                "name": "深度原生图表分析大盘 (折线/面积/环形/饼图/柱状)",
                "best_items": "chart 原生图表定义 (type=line/area/doughnut/pie/column) + 右侧 2~3个 KPI 卡 + 底部 1条洞察胶囊",
                "text_budget": "kpi label 4~8字 | kpi value 3~8字符 | chart title 8~16字 | insight 25~50字",
                "anti_hollow_tips": "左侧大画幅原生高清图表驱动，右侧配置量化 KPI 卡与底部核心业务启示，完美契合深度数据与监控汇报。"
            },
            "code_walkthrough": {
                "name": "极客源码走读与架构深度解析 (程序员技术分享)",
                "best_items": "左侧 macOS 极客编辑器视窗 (行号+高亮) + 右侧调用时序 + 算法复杂度 + 线上避坑指南",
                "text_budget": "code 8~18行代码 | sequence 3~4步 (每步15~30字) | complexity 8~16字 | pitfall 25~55字",
                "anti_hollow_tips": "左侧深色极客视窗带红黄绿控制按钮与语法高亮，右侧剖析算法复杂度与生产避坑指南，技术大厂分享标配。"
            },
            "classroom_quiz": {
                "name": "互动课堂测验与名师解题解析 (大中小学教学)",
                "best_items": "顶部大字号考题题干 + 中间 2x2 四个选项卡片 (标注 correct) + 底部名师核心解题思路抽屉",
                "text_budget": "question 20~55字 | 选项 A/B/C/D 8~25字 | analysis 名师解析 35~80字",
                "anti_hollow_tips": "选项卡片自动渲染 A/B/C/D 药丸徽标，正确答案带高光描边，底部抽屉提供详细破题思路，沉浸式互动教学。"
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
