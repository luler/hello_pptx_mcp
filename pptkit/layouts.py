# -*- coding: utf-8 -*-
"""现代化版式构件库：页眉系统 / 导读摘要 / Bento KPI 卡 / 详情面板 / 2x2网格 / 对比卡 / 时间线 / 表格 / 页脚。

全面采用现代化网格系统（Grid System）与呼吸感留白，
自适应浅色商务、深色极客、自然生态等各类主题，杜绝大片留白与空洞。
"""
import re
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from .charts import bar_chart


MARGIN = 0.65
CONTENT_W = 12.033

_TAG_RE = re.compile(r"\[\[(.*?)\]\]")


def parse_rich(text, color=None, bold=True):
    """将 '普通[[强调]]文本' 解析为 Canvas.para 混排 runs。"""
    if isinstance(text, (list, tuple)):
        return text
    text = "" if text is None else str(text)
    runs, pos = [], 0
    for m in _TAG_RE.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], {}))
        opt = {"b": bold}
        if color is not None:
            opt["c"] = color
        runs.append((m.group(1), opt))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], {}))
    return runs or [("", {})]


def chrome(cv, sw, sh, thick=0.04):
    """现代化顶部微光呼吸条与空间环境光影（极简高级，彻底消除单调死板平铺）。"""
    pal = cv.pal
    is_dark = pal.get("is_dark", False)

    # 1. 空间环境光影层（Ambient Lighting Aura）：营造景深与空气感
    if is_dark:
        # 右上角柔和深邃极光光晕
        glow_fill = pal.get("primary_tint", pal["card_subtle"])
        cv.circle(sw - 3.8, -1.6, 5.2, fill=glow_fill)
        # 左下角次级微光晕
        cv.circle(-1.8, sh - 2.8, 4.6, fill=pal.get("bg_alt", pal["card_subtle"]))
        # 极简科技十字标（微光点缀）
        cv.text(MARGIN, 0.12, 0.4, 0.2, "＋", size=8.5, color=pal.get("card_border"))
        cv.text(sw - MARGIN - 0.25, 0.12, 0.4, 0.2, "＋", size=8.5, color=pal.get("card_border"))
    else:
        # 浅色商务背景：柔和顶角微晕
        cv.circle(sw - 3.2, -1.8, 4.8, fill=pal.get("primary_tint", pal["bg_alt"]))

    # 2. 顶部微光发光条与底边边线
    cv.rect(0, 0, sw, thick, fill=pal.get("primary", pal["primary"]))
    cv.rect(0, 0, 2.4, thick + 0.015, fill=pal.get("accent", pal["accent2"]))
    if is_dark:
        cv.rect(0, sh - 0.02, sw, 0.02, fill=pal.get("card_border"))


def header(cv, title, subtitle=None, category=None, y=0.38, size=23):
    """现代化幻灯片主页眉系统。"""
    pal = cv.pal
    title_col = pal.get("text_title", pal["primary"])
    sub_col = pal.get("text_subtitle", pal["text_body"])

    # 1. 主标题
    cv.text(MARGIN, y, 7.8, 0.52, str(title or ""), size=size, bold=True,
            color=title_col, line=1.05)

    # 2. 右侧副标题或分类提示
    if subtitle:
        cv.text(7.6, y + 0.14, 5.08, 0.36, str(subtitle), size=11.5,
                color=sub_col, align="right", line=1.1)

    # 3. 极细高级分割线 + 品牌强调微线
    line_y = y + 0.62
    cv.divider(MARGIN, line_y, CONTENT_W)
    cv.rect(MARGIN, line_y - 0.005, 0.85, 0.025, fill=pal.get("accent", pal["primary"]))


def format_item_text(it):
    """智能解析列表项：支持纯字符串或包含 label/value/title/desc 的结构化字典。"""
    if isinstance(it, dict):
        if "label" in it and "value" in it:
            val = it["value"]
            if isinstance(val, float) and 0 <= val <= 1.0:
                v_str = "%.0f%%" % (val * 100)
            else:
                v_str = str(val)
            return "%s: [[%s]]" % (it["label"], v_str)
        elif "title" in it:
            t = it["title"]
            d = it.get("desc") or it.get("text") or it.get("value") or ""
            return "%s: %s" % (t, d) if d else t
        elif "text" in it:
            return str(it["text"])
        elif "desc" in it:
            return str(it["desc"])
    return str(it)


def lead_bar(cv, runs, y=1.18, h=0.62):
    """导读/核心观点摘要卡片（整页战略结论提炼）。"""
    pal = cv.pal
    cv.card(MARGIN, y, CONTENT_W, h,
            fill=pal.get("card_subtle", pal["bar_bg"]),
            line=pal.get("card_border", pal["border"]),
            lw=0.75, radius=0.12)

    # 左侧垂直强调条
    cv.pill(MARGIN + 0.16, y + 0.14, 0.05, h - 0.28,
            fill=pal.get("accent", pal["primary"]))

    tf = cv.tbox(MARGIN + 0.32, y, CONTENT_W - 0.50, h, anchor=MSO_ANCHOR.MIDDLE)
    cv.para(tf, runs, first=True, size=11.5,
            color=pal.get("text_body", pal["text"]), line=1.2)


def kpi_row(cv, cards, y=1.92, h=1.50, gap=0.20, cols=None):
    """Bento-Grid 风格的现代化指标卡行，支持多彩动态 Accent，自适应高度与字体缩放。"""
    pal = cv.pal
    n = cols or len(cards)
    if n == 0:
        return y
    cw = (CONTENT_W - gap * (n - 1)) / float(n)
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

    for i, c in enumerate(cards[:n]):
        cx = MARGIN + i * (cw + gap)
        card_accent = c.get("color") or accents[i % len(accents)]

        # 1. 浮岛卡片容器
        cv.card(cx, y, cw, h, radius=0.08)

        # 2. 顶部精致微光发光条（赋予每个卡片独立科技/商业色彩）
        cv.pill(cx + 0.20, y + 0.05, min(1.20, cw - 0.40), 0.035, fill=card_accent)

        # 3. 分类徽章
        label_text = str(c.get("label", ""))
        cv.badge(cx + 0.20, y + 0.18, label_text, dot=True,
                 dot_color=card_accent, text_color=card_accent, size=9.5, h=0.28)

        val_str = str(c.get("value", ""))
        unit_str = str(c.get("unit", ""))

        caps = c.get("captions") or ([c["caption"]] if c.get("caption") else [])
        if not caps:
            if c.get("desc"):
                caps = [c["desc"]]
            elif c.get("sub"):
                caps = [c["sub"]]
            elif c.get("target"):
                caps = [c["target"]]

        if caps:
            num_y = y + (0.46 if h <= 1.8 else 0.58)
            num_sz = 30 if h <= 1.8 else 36
            num_runs = [(val_str, {"b": True, "c": card_accent, "sz": num_sz})]
            if unit_str:
                num_runs.append((" " + unit_str,
                                 {"b": True, "c": card_accent, "sz": 14 if h <= 1.8 else 16}))
            cv.text(cx + 0.20, num_y, cw - 0.40, 0.50 if h <= 1.8 else 0.62, num_runs, line=1.0)

            div_y = num_y + (0.48 if h <= 1.8 else 0.64)
            cv.divider(cx + 0.20, div_y, cw - 0.40)

            caps_y = div_y + 0.08
            caps_h = max(0.40, (y + h) - caps_y - 0.10)
            tf = cv.tbox(cx + 0.20, caps_y, cw - 0.40, caps_h)
            for j, line in enumerate(caps[:3]):
                cv.para(tf, parse_rich(line, color=card_accent),
                        size=9.5 if h <= 1.8 else 10.5,
                        color=pal.get("text_body", pal.get("text_subtitle")),
                        line=1.20, first=(j == 0))
        else:
            # 无副文本时，数值居中放大显示，底边放置强调胶囊
            num_y = y + 0.35 + max(0, (h - 1.20) * 0.40)
            num_sz = 36 if h <= 1.8 else 44
            num_runs = [(val_str, {"b": True, "c": card_accent, "sz": num_sz})]
            if unit_str:
                num_runs.append((" " + unit_str,
                                 {"b": True, "c": card_accent, "sz": 16}))
            cv.text(cx + 0.20, num_y, cw - 0.40, 0.70, num_runs, line=1.0)
            cv.pill(cx + 0.20, y + h - 0.14, min(1.35, cw - 0.40), 0.035, fill=card_accent)

    return y + h


def panel(cv, x, y, w, h, title, subtitle=None, accent_color=None):
    """Bento 详情面板容器（含小图标标号、主标题与副标题）。"""
    pal = cv.pal
    acc = accent_color or pal.get("accent", pal["primary"])
    cv.card(x, y, w, h, radius=0.08)

    # 顶部发光微条
    cv.pill(x + 0.24, y + 0.04, min(1.2, w - 0.48), 0.03, fill=acc)

    # 标题图标小徽标
    cv.icon_box(x + 0.24, y + 0.18, size=0.28, char="▪", color=acc)

    # 面板标题
    cv.text(x + 0.60, y + 0.15, w - 0.84, 0.32, str(title or ""),
            size=14.5, bold=True, color=pal.get("text_title", pal["primary"]),
            line=1.05)

    body_y = y + 0.54
    if subtitle:
        cv.text(x + 0.24, body_y, w - 0.48, 0.24, str(subtitle),
                size=9.8, color=pal.get("text_subtitle", pal["text_body"]), line=1.1)
        body_y += 0.26

    # 细微分割线
    cv.divider(x + 0.24, body_y, w - 0.48)
    body_y += 0.12

    return x + 0.24, body_y, w - 0.48


def item_block(cv, x, y, w, h, text, index=None, icon="•", size=10.5, accent_color=None):
    """单条重点的高级微卡片（用于将稀疏列表包裹为充实的 Bento 砖块，拒绝空洞）。"""
    pal = cv.pal
    acc = accent_color or pal.get("accent", pal["primary"])
    cv.card(x, y, w, h, fill=pal.get("card_subtle"), lw=0.5, radius=0.06)

    # 序号或图标徽章
    badge_str = ("%02d" % index) if index is not None else icon
    cv.badge(x + 0.16, y + (h - 0.28) / 2.0, badge_str,
             dot=False, bg=pal.get("badge_bg"), text_color=acc,
             size=9, h=0.28, pad_x=0.08)

    text_x = x + 0.75
    text_w = w - 0.90
    tf = cv.tbox(text_x, y, text_w, h, anchor=MSO_ANCHOR.MIDDLE)
    clean_text = format_item_text(text)
    runs = parse_rich(clean_text, color=acc)
    cv.para(tf, runs, size=size, bold=False,
            color=pal.get("text_body", pal["text"]), line=1.2, first=True)


def bullet_list(cv, x, y, w, h, items, size=11, gap=7, line=1.28, dot_color=None):
    """优雅的项目要点列表，自动支持结构化字典数据。"""
    tf = cv.tbox(x, y, w, h)
    for i, it in enumerate(items):
        clean_text = format_item_text(it)
        runs = parse_rich(clean_text, color=dot_color or cv.pal.get("accent"))
        cv.bullet(tf, runs, first=(i == 0), size=size,
                  before=(0 if i == 0 else gap), line=line, dot=dot_color)
    return tf


def grid_cards(cv, y, cards, cols=2, h=None):
    """现代化 Bento Grid 网格卡片（如 2x2 战略支柱、4 大核心能力矩阵），全面充盈宽屏空间。"""
    pal = cv.pal
    n = len(cards)
    if n == 0:
        return y

    cols = max(1, min(cols, 4))
    gap_x = 0.24
    gap_y = 0.22
    rows_cnt = (n + cols - 1) // cols
    cw = (CONTENT_W - gap_x * (cols - 1)) / float(cols)

    card_h = h or (2.4 if rows_cnt == 2 else (3.2 if rows_cnt == 1 else 1.8))
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

    for i, c in enumerate(cards):
        ri = i // cols
        ci = i % cols
        cx = MARGIN + ci * (cw + gap_x)
        cy = y + ri * (card_h + gap_y)
        card_accent = c.get("color") or accents[i % len(accents)]

        # 浮岛主卡
        cv.card(cx, cy, cw, card_h, radius=0.08)

        # 顶部精致发光微条
        cv.pill(cx + 0.22, cy + 0.04, min(1.30, cw - 0.44), 0.035, fill=card_accent)

        # 卡片头部：标号徽章 + 标题
        idx_str = "%02d" % (i + 1)
        cv.badge(cx + 0.22, cy + 0.18, idx_str, dot=False,
                 bg=pal.get("badge_bg"), text_color=card_accent,
                 size=9.5, h=0.28)

        title_str = str(c.get("title", ""))
        tag_str = c.get("tag") or c.get("metric") or c.get("label") or c.get("badge")
        tag_w = 1.35 if tag_str else 0.0
        cv.text(cx + 0.82, cy + 0.15, max(0.5, cw - 0.82 - tag_w - 0.20), 0.36, title_str,
                size=13.5, bold=True, color=pal.get("text_title", pal["primary"]),
                line=1.05)

        # 辅助标签或指标（右侧）
        if tag_str:
            cv.badge(cx + cw - tag_w - 0.10, cy + 0.18, str(tag_str), dot=True,
                     dot_color=card_accent, text_color=card_accent,
                     size=8.5, h=0.26)

        # 细微分割线
        cv.divider(cx + 0.22, cy + 0.58, cw - 0.44)

        # 内容区域：支持 items / bullets 列表 或 desc 段落
        items = c.get("items") or c.get("bullets")
        desc = c.get("desc") or c.get("description") or c.get("summary")
        content_y = cy + 0.68
        content_h = card_h - 0.80

        if items:
            bullet_list(cv, cx + 0.22, content_y, cw - 0.44, content_h,
                        items, size=10.5, gap=5, line=1.24, dot_color=card_accent)
        elif desc:
            tf = cv.tbox(cx + 0.22, content_y, cw - 0.44, content_h)
            runs = parse_rich(desc, color=card_accent)
            cv.para(tf, runs, size=11, color=pal.get("text_body", pal["text"]),
                    line=1.35, first=True)

    return y + rows_cnt * card_h + (rows_cnt - 1) * gap_y


def comparison_cards(cv, y, left_card, right_card, h=3.8):
    """现代化对比卡片（如 Before vs After / 痛点 vs 突破 / 现有架构 vs 演进方案）。"""
    pal = cv.pal
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    left_accent = accents[2 % len(accents)] if len(accents) > 2 else pal.get("accent2", pal.get("text_muted"))
    right_accent = accents[0] if accents else pal.get("accent", pal["primary"])

    gap = 0.32
    cw = (CONTENT_W - gap) / 2.0

    # 1. 左侧（现状/挑战卡片，柔和沉稳底色）
    lx = MARGIN
    cv.card(lx, y, cw, h, fill=pal.get("card_subtle"), radius=0.08)
    cv.pill(lx + 0.30, y + 0.05, min(1.4, cw - 0.60), 0.035, fill=left_accent)
    left_tag = left_card.get("tag") or left_card.get("badge") or "挑战 / 现状"
    cv.badge(lx + 0.30, y + 0.22, left_tag,
             dot=True, dot_color=left_accent, bg=pal.get("card_border"), text_color=pal.get("text_muted"),
             size=9.5, h=0.30)
    cv.text(lx + 0.30, y + 0.65, cw - 0.60, 0.40,
            str(left_card.get("title", "")), size=16, bold=True,
            color=pal.get("text_title", pal["primary"]), line=1.1)
    cv.divider(lx + 0.30, y + 1.12, cw - 0.60)
    left_items = left_card.get("items") or left_card.get("bullets") or []
    if left_items:
        bullet_list(cv, lx + 0.30, y + 1.25, cw - 0.60, h - 1.45,
                    left_items, size=11, gap=8, line=1.3, dot_color=left_accent)

    # 2. 右侧（突破/新方案卡片，品牌高光强调）
    rx = MARGIN + cw + gap
    cv.card(rx, y, cw, h, fill=pal.get("card_bg"),
            line=right_accent, lw=1.25, radius=0.08)
    cv.pill(rx + 0.30, y + 0.05, min(1.4, cw - 0.60), 0.035, fill=right_accent)
    right_tag = right_card.get("tag") or right_card.get("badge") or "核心突破 / 解决方案"
    cv.badge(rx + 0.30, y + 0.22, right_tag,
             dot=True, dot_color=right_accent, bg=pal.get("badge_bg"), text_color=right_accent,
             size=9.5, h=0.30)
    cv.text(rx + 0.30, y + 0.65, cw - 0.60, 0.40,
            str(right_card.get("title", "")), size=16, bold=True,
            color=right_accent, line=1.1)
    cv.divider(rx + 0.30, y + 1.12, cw - 0.60)
    right_items = right_card.get("items") or right_card.get("bullets") or []
    if right_items:
        bullet_list(cv, rx + 0.30, y + 1.25, cw - 0.60, h - 1.45,
                    right_items, size=11, gap=8, line=1.3, dot_color=right_accent)


def architecture_stack(cv, y, layers, h=4.3):
    """现代分层系统架构图（Layered Architecture Stack，如接入层 -> 业务计算层 -> 存储层）。"""
    pal = cv.pal
    n = len(layers)
    if n == 0:
        return y
    n = min(n, 4)
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    gap = 0.22
    layer_h = (h - gap * (n - 1)) / float(n)

    for i, lyr in enumerate(layers[:n]):
        ly = y + i * (layer_h + gap)
        layer_acc = lyr.get("color") or accents[i % len(accents)]

        # 主卡片底槽
        cv.card(MARGIN, ly, CONTENT_W, layer_h, radius=0.08)

        # 左侧层级标识胶囊条
        cv.pill(MARGIN + 0.16, ly + 0.12, 0.045, layer_h - 0.24, fill=layer_acc)

        # 层级徽标（例如 'L1 接入网关层'）
        idx_str = lyr.get("layer_tag") or ("LAYER %02d" % (i + 1))
        badge_end_x = cv.badge(MARGIN + 0.30, ly + 0.18, idx_str, dot=True,
                               dot_color=layer_acc, text_color=layer_acc, size=9.5, h=0.28)

        # 层级标题（动态计算位置，确保与徽标绝不重叠）
        right_w = 6.0
        rx = MARGIN + CONTENT_W - right_w - 0.25
        title_x = badge_end_x + 0.18
        title_w = max(1.0, rx - title_x - 0.20)
        title_str = str(lyr.get("title", ""))
        cv.text(title_x, ly + 0.15, title_w, 0.34, title_str,
                size=13.5, bold=True, color=pal.get("text_title", pal["primary"]))

        # 描述说明（标题下方）
        desc_str = lyr.get("desc")
        if desc_str:
            cv.text(MARGIN + 0.30, ly + 0.54, 5.2, layer_h - 0.62, str(desc_str),
                    size=10, color=pal.get("text_body", pal["text"]), line=1.2)

        # 右侧技术组件胶囊阵列（Components Pills）
        comps = lyr.get("components") or lyr.get("items") or lyr.get("chips") or []
        if comps:
            right_w = 6.0
            rx = MARGIN + CONTENT_W - right_w - 0.25
            n_comps = min(len(comps), 5)
            comp_gap = 0.15
            cw = (right_w - comp_gap * (n_comps - 1)) / float(n_comps)
            ch = min(0.68, layer_h - 0.36)
            cy_pos = ly + (layer_h - ch) / 2.0
            for ci, comp in enumerate(comps[:n_comps]):
                comp_x = rx + ci * (cw + comp_gap)
                cv.card(comp_x, cy_pos, cw, ch, fill=pal.get("card_subtle"),
                        line=layer_acc, lw=0.6, radius=0.06)
                cv.text(comp_x + 0.08, cy_pos, cw - 0.16, ch, str(comp),
                        size=10, bold=True, color=pal.get("text_body", pal["text"]),
                        align="center", anchor=MSO_ANCHOR.MIDDLE)

        # 层与层之间的微向下流转指示（在两层中间）
        if i < n - 1:
            arrow_y = ly + layer_h + 0.04
            cv.pill(MARGIN + CONTENT_W / 2.0 - 0.40, arrow_y, 0.80, 0.14,
                    fill=pal.get("badge_bg"), line=None)
            tf_arr = cv.tbox(MARGIN + CONTENT_W / 2.0 - 0.40, arrow_y, 0.80, 0.14, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(tf_arr, "▼", size=8.5, color=layer_acc, align="center", first=True)

    return y + h


def hero_statement(cv, y, statement, points=None, h=4.3):
    """核心论断/战略大字报页（打破满屏小卡片，以巨大冲击力震撼全场）。"""
    pal = cv.pal
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    acc = accents[0]

    # 外层大发光卡片
    cv.card(MARGIN, y, CONTENT_W, h, fill=pal.get("card_bg"),
            line=pal.get("card_border"), lw=1.0, radius=0.10)

    # 顶部品牌发光装饰
    cv.pill(MARGIN + 0.40, y + 0.05, 2.0, 0.04, fill=acc)

    # 核心大字论断（Hero Typography，22~28pt 震撼高对比）
    tf = cv.tbox(MARGIN + 0.50, y + 0.45, CONTENT_W - 1.0, 1.45, anchor=MSO_ANCHOR.MIDDLE)
    runs = parse_rich(statement, color=acc, bold=True)
    cv.para(tf, runs, size=24, bold=True,
            color=pal.get("text_title", pal["primary"]), align="center", line=1.35, first=True)

    # 中部微细分割线
    cv.divider(MARGIN + 0.80, y + 2.05, CONTENT_W - 1.60)

    # 下方 3~4 个关键支柱徽章卡片
    pts = points or []
    if pts:
        n_p = min(len(pts), 4)
        gap = 0.25
        card_w = (CONTENT_W - 1.0 - gap * (n_p - 1)) / float(n_p)
        card_h = 1.45
        card_y = y + 2.25

        for i, pt in enumerate(pts[:n_p]):
            cx = MARGIN + 0.50 + i * (card_w + gap)
            pt_acc = accents[i % len(accents)]
            cv.card(cx, card_y, card_w, card_h, fill=pal.get("card_subtle"),
                    line=pt_acc, lw=0.6, radius=0.08)

            idx_str = "%02d" % (i + 1)
            cv.badge(cx + 0.16, card_y + 0.16, idx_str, dot=False,
                     bg=pal.get("badge_bg"), text_color=pt_acc, size=9, h=0.26)

            pt_text = format_item_text(pt)
            tf_pt = cv.tbox(cx + 0.16, card_y + 0.48, card_w - 0.32, card_h - 0.56)
            pt_runs = parse_rich(pt_text, color=pt_acc)
            cv.para(tf_pt, pt_runs, size=10.5,
                    color=pal.get("text_body", pal["text"]), line=1.25, first=True)

    return y + h


def process_flow(cv, y, steps, h=3.6):
    """带清晰流向箭头（➔）的横向流程演进卡片。"""
    pal = cv.pal
    n = len(steps)
    if n == 0:
        return y
    n = min(n, 5)
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    arrow_w = 0.35
    total_arrows_w = arrow_w * (n - 1)
    card_w = (CONTENT_W - total_arrows_w - 0.20 * (n - 1)) / float(n)

    cur_x = MARGIN
    for i, s in enumerate(steps[:n]):
        step_acc = accents[i % len(accents)]

        # 步骤卡片
        cv.card(cur_x, y, card_w, h, radius=0.08)

        # 顶部微光发光条
        cv.pill(cur_x + 0.20, y + 0.05, min(1.0, card_w - 0.40), 0.035, fill=step_acc)

        # 步骤标号胶囊
        idx_str = "STEP %02d" % (i + 1)
        cv.badge(cur_x + 0.20, y + 0.18, idx_str, dot=True,
                 dot_color=step_acc, text_color=step_acc, size=9, h=0.28)

        # 步骤主标题
        title_str = str(s.get("title", ""))
        cv.text(cur_x + 0.20, y + 0.54, card_w - 0.40, 0.36, title_str,
                size=13, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        cv.divider(cur_x + 0.20, y + 0.94, card_w - 0.40)

        # 步骤详述或要点
        desc_str = s.get("desc") or s.get("summary")
        items = s.get("items") or s.get("bullets")
        content_y = y + 1.06
        content_h = h - 1.20

        if items:
            bullet_list(cv, cur_x + 0.20, content_y, card_w - 0.40, content_h,
                        items, size=10, gap=5, line=1.22, dot_color=step_acc)
        elif desc_str:
            tf = cv.tbox(cur_x + 0.20, content_y, card_w - 0.40, content_h)
            runs = parse_rich(desc_str, color=step_acc)
            cv.para(tf, runs, size=10.5, color=pal.get("text_body", pal["text"]),
                    line=1.30, first=True)

        cur_x += card_w

        # 步骤间的流向粗箭头
        if i < n - 1:
            arrow_x = cur_x + 0.05
            arrow_y = y + h * 0.40
            cv.pill(arrow_x, arrow_y, arrow_w + 0.10, 0.32,
                    fill=pal.get("badge_bg"), line=None)
            tf_a = cv.tbox(arrow_x, arrow_y, arrow_w + 0.10, 0.32, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(tf_a, "➔", size=13, bold=True, color=step_acc, align="center", first=True)
            cur_x += arrow_w + 0.20

    return y + h


def panel_with_chart(cv, x, y, w, h, title, subtitle=None, items=None,
                      categories=None, series=None, chart_h=1.15, accent_color=None):
    """复合图表面板：上部要点文字 + 下部现代柱状图。"""
    px, by, bw = panel(cv, x, y, w, h, title, subtitle, accent_color=accent_color)
    text_h = max(0.5, h - (by - y) - chart_h - 0.20)
    if items:
        bullet_list(cv, px, by, bw, text_h, items, size=10, gap=5, line=1.2, dot_color=accent_color)
    if categories and series:
        chart_y = y + h - chart_h - 0.15
        bar_chart(cv, px + 0.05, chart_y, bw - 0.10, chart_h,
                  categories, series)


def timeline(cv, y, items, x=None, w=None, h=None):
    """现代化横向里程碑时间线：节点胶囊 + 发光微卡片。"""
    pal = cv.pal
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    x = MARGIN if x is None else x
    w = CONTENT_W if w is None else w
    n = len(items)
    if n == 0:
        return

    step = w / float(n)
    line_y = y + 0.48

    # 贯通连接线
    cv.rect(x + step * 0.25, line_y, w - step * 0.5, 0.025,
            fill=pal.get("card_border", pal["border"]))

    card_h = h or 3.20

    for i, it in enumerate(items):
        cx = x + step * i
        center_x = cx + step * 0.5
        step_accent = accents[i % len(accents)]

        # 阶段序号徽章（如 01 / Q1）
        phase_str = "%02d" % (i + 1)
        cv.pill(center_x - 0.26, y, 0.52, 0.26,
                fill=pal.get("badge_bg", pal["primary_tint"]))
        tf_badge = cv.tbox(center_x - 0.26, y, 0.52, 0.26, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_badge, phase_str, size=9.5, bold=True,
                color=step_accent, align="center", first=True)

        # 轨道节点圆环
        cv.circle(center_x - 0.09, line_y - 0.075, 0.18,
                  fill=pal.get("card_bg"),
                  line=step_accent, lw=1.5)
        cv.circle(center_x - 0.045, line_y - 0.03, 0.09,
                  fill=step_accent)

        # 下方卡片
        card_w = step - 0.20
        card_y = line_y + 0.26
        cv.card(cx + 0.10, card_y, card_w, card_h, radius=0.08)

        # 卡片顶部发光微条
        cv.pill(cx + 0.24, card_y + 0.04, min(1.0, card_w - 0.48), 0.03, fill=step_accent)

        # 卡片内标题
        cv.text(cx + 0.24, card_y + 0.20, card_w - 0.48, 0.36,
                str(it.get("title", "")), size=13.5, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.1)

        cv.divider(cx + 0.24, card_y + 0.60, card_w - 0.48)

        desc = it.get("desc") or it.get("subtitle") or it.get("text")
        if desc:
            tf = cv.tbox(cx + 0.24, card_y + 0.72, card_w - 0.48, card_h - 0.85)
            runs = parse_rich(desc, color=step_accent)
            cv.para(tf, runs, size=10.5,
                    color=pal.get("text_body", pal["text"]), line=1.35, first=True)


def table(cv, x, y, w, rows, col_widths=None, header=True,
          row_h=0.48, size=11, header_size=11.5):
    """现代化卡片式表格：圆角表头、平滑行高、微光斑马纹。"""
    pal = cv.pal
    if not rows:
        return y
    ncol = len(rows[0])
    if col_widths:
        total = float(sum(col_widths))
        cws = [w * cw / total for cw in col_widths]
    else:
        cws = [w / float(ncol)] * ncol

    total_h = len(rows) * row_h
    # 外层卡片包围框
    cv.card(x, y, w, total_h, fill=pal.get("card_bg"), radius=0.08)

    for ri, row in enumerate(rows):
        is_head = header and ri == 0
        ry = y + ri * row_h
        cx = x

        # 行背景
        if is_head:
            row_fill = pal.get("table_header_bg", pal["primary"])
            cv.rrect(cx, ry, w, row_h, fill=row_fill, radius=0.08)
        else:
            row_fill = pal.get("table_alt_row") if (ri % 2 == 1) else pal.get("card_bg")
            cv.rect(cx, ry, w, row_h, fill=row_fill)
            cv.divider(cx, ry, w)

        for ci, cell in enumerate(row):
            txt_col = pal.get("table_header_text", pal["white"]) if is_head else (
                pal.get("text_title", pal["primary"]) if ci == 0 else pal.get("text_body", pal["text"])
            )
            cell_runs = parse_rich(str(cell), color=pal.get("accent")) if not is_head else str(cell)
            cv.text(cx + 0.14, ry, cws[ci] - 0.28, row_h, cell_runs,
                    size=header_size if is_head else size,
                    bold=is_head or (ci == 0),
                    color=txt_col,
                    align="center" if (is_head or ci > 0) else "left",
                    anchor=MSO_ANCHOR.MIDDLE)
            cx += cws[ci]

    return y + total_h


def takeaway_card(cv, text, y, title="核心战略洞察与建议", h=None, accent_color=None):
    """底部关键洞察提炼大卡片（用于填充 KPI、表格或时间线底部的空白，提升汇报价值感）。"""
    pal = cv.pal
    card_h = h or 0.95
    acc = accent_color or pal.get("accent", pal["primary"])
    cv.card(MARGIN, y, CONTENT_W, card_h,
            fill=pal.get("card_subtle"),
            line=pal.get("card_border"), lw=0.9, radius=0.08)

    # 左侧垂直高光条
    cv.pill(MARGIN + 0.14, y + 0.12, 0.045, card_h - 0.24, fill=acc)

    # 标志性小徽章
    cv.badge(MARGIN + 0.28, y + 0.16, title, dot=True, dot_color=acc, text_color=acc, size=9.5, h=0.28)
    tf = cv.tbox(MARGIN + 0.28, y + 0.48, CONTENT_W - 0.50, card_h - 0.56, anchor=MSO_ANCHOR.MIDDLE)
    runs = parse_rich(text, color=acc)
    cv.para(tf, runs, size=11, color=pal.get("text_body", pal["text"]), line=1.28, first=True)


def poetry_card(cv, y, poem_lines, title=None, author=None, interpretation=None, h=4.3):
    """东方水墨诗词与韵律美学卡片：中式典雅印章、名句对仗对阕与战略寓意解读。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])  # 朱砂红
    cyan_green = pal.get("accent2", pal["primary"])  # 汝窑天青/黛绿

    # 外层古纸浮岛大卡片
    cv.card(MARGIN, y, CONTENT_W, h, fill=pal.get("card_bg"),
            line=pal.get("card_border"), lw=1.0, radius=0.10)

    # 顶部朱砂印泥微饰条
    cv.pill(MARGIN + 0.40, y + 0.05, 2.0, 0.035, fill=acc)

    # 词牌名/出处印章徽标（例如 '【 词牌 · 破阵子 】' 或 '【 经典名篇 · 借诗明志 】'）
    tag_str = title or "经典名篇 · 借诗明志"
    badge_end = cv.badge(MARGIN + 0.45, y + 0.22, tag_str, dot=True,
                         dot_color=acc, bg=pal.get("badge_bg"), text_color=acc, size=10, h=0.32)

    if author:
        cv.text(badge_end + 0.20, y + 0.22, 4.0, 0.30, "〔 %s 〕" % author,
                size=11, color=pal.get("text_muted"))

    # 诗词正文：采用优雅大气对仗排版，字号 21pt 震撼居中
    poem_y = y + 0.72
    poem_h = 1.35
    tf_poem = cv.tbox(MARGIN + 0.60, poem_y, CONTENT_W - 1.20, poem_h, anchor=MSO_ANCHOR.MIDDLE)
    if isinstance(poem_lines, (list, tuple)):
        lines_text = "\n".join(str(l) for l in poem_lines)
    else:
        lines_text = str(poem_lines)

    runs = parse_rich(lines_text, color=acc, bold=True)
    cv.para(tf_poem, runs, size=21, bold=True,
            color=pal.get("text_title", pal["primary"]), align="center", line=1.45, first=True)

    # 优雅古朴分割线 + 居中朱红印章图样
    cv.divider(MARGIN + 0.80, y + 2.18, CONTENT_W - 1.60)
    cv.badge(MARGIN + CONTENT_W / 2.0 - 0.45, y + 2.02, "【印】",
             dot=False, bg=acc, text_color=pal.get("card_bg"), size=8.5, h=0.28, pad_x=0.08)

    # 下方：战略寓意或现代业务映射
    if interpretation:
        tf_int = cv.tbox(MARGIN + 0.60, y + 2.45, CONTENT_W - 1.20, h - 2.65)
        int_runs = parse_rich(interpretation, color=cyan_green)
        cv.para(tf_int, int_runs, size=12, color=pal.get("text_body"), line=1.40, first=True)

    return y + h


def topology_diagram(cv, y, nodes, edges=None, h=4.3):
    """现代分布式拓扑架构图（Topology Diagram）：节点方阵 + 全向调用流转指引。"""
    pal = cv.pal
    if not nodes:
        return y
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

    # 外层大架构底盘卡片（悬浮质感）
    cv.card(MARGIN, y, CONTENT_W, h, fill=pal.get("card_bg"),
            line=pal.get("card_border"), lw=0.75, radius=0.08)

    # 1. 智能行列分析与 0-based 归一化（自适应 0 起始或 1 起始）
    has_grid = any("col" in n and "row" in n for n in nodes)
    if has_grid:
        raw_cols = [int(n.get("col", 0)) for n in nodes]
        raw_rows = [int(n.get("row", 0)) for n in nodes]
        min_c = min(raw_cols)
        min_r = min(raw_rows)
        cols = max(raw_cols) - min_c + 1
        rows = max(raw_rows) - min_r + 1
    else:
        n_count = len(nodes)
        cols = 3 if n_count <= 6 else (4 if n_count <= 8 else 3)
        rows = (n_count + cols - 1) // cols
        min_c, min_r = 0, 0

    cols = max(1, min(cols, 5))
    rows = max(1, min(rows, 4))

    # 2. 动态计算节点尺寸与全局居中 Padding
    gap_x = 0.80 if cols <= 3 else 0.50
    gap_y = 0.60 if rows <= 2 else 0.40

    avail_w = CONTENT_W - 0.70
    avail_h = h - 0.60
    node_w = min(3.40, (avail_w - gap_x * (cols - 1)) / float(cols))
    node_h = min(1.60, (avail_h - gap_y * (rows - 1)) / float(rows))

    total_w = cols * node_w + (cols - 1) * gap_x
    total_h = rows * node_h + (rows - 1) * gap_y
    pad_x = (CONTENT_W - total_w) / 2.0
    pad_y = (h - total_h) / 2.0

    node_pos = {}
    for i, n in enumerate(nodes):
        if has_grid and "col" in n and "row" in n:
            ci = int(n["col"]) - min_c
            ri = int(n["row"]) - min_r
        else:
            ci = i % cols
            ri = i // cols
        ci = max(0, min(ci, cols - 1))
        ri = max(0, min(ri, rows - 1))

        nx = MARGIN + pad_x + ci * (node_w + gap_x)
        ny = y + pad_y + ri * (node_h + gap_y)

        info = (nx, ny, node_w, node_h, ci, ri)
        # 建立多别名键映射，确保 edges 通过 id, label, title 或索引均可命中
        for k in ("id", "label", "title", "name"):
            val = n.get(k)
            if val is not None:
                node_pos[str(val).strip()] = info
        node_pos[str(i)] = info
        node_pos[f"node_{i}"] = info
        node_pos[f"node_{i+1}"] = info

        node_acc = n.get("color") or accents[i % len(accents)]

        # 节点卡片：圆润边角与强调外框
        cv.card(nx, ny, node_w, node_h, fill=pal.get("card_subtle"),
                line=node_acc, lw=0.9, radius=0.06)

        # 顶部微光发光饰条
        cv.pill(nx + 0.14, ny + 0.04, min(1.2, node_w - 0.28), 0.035, fill=node_acc)

        # 节点类型标签胶囊（如 [DB]、[SERVICE]、[CACHE]、[APP]）
        tag_str = str(n.get("tag") or n.get("type") or "NODE").upper()
        cv.badge(nx + 0.14, ny + 0.14, tag_str, dot=True,
                 dot_color=node_acc, text_color=node_acc, size=8.5, h=0.24, pad_x=0.08)

        # 节点核心标题
        title_str = str(n.get("label") or n.get("title") or n.get("id") or f"节点 {i+1}")
        cv.text(nx + 0.14, ny + 0.44, node_w - 0.28, 0.32, title_str,
                size=11.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.05)

        # 节点技术说明 / 描述
        desc_str = n.get("desc") or n.get("sub")
        if desc_str and node_h >= 0.95:
            cv.text(nx + 0.14, ny + 0.82, node_w - 0.28, node_h - 0.90, str(desc_str),
                    size=9.2, color=pal.get("text_muted"), line=1.18)

    # 绘制全向有向拓扑边与调用流转标签
    if edges:
        for ed in edges:
            f_key = str(ed.get("from") or "").strip()
            t_key = str(ed.get("to") or "").strip()
            if f_key in node_pos and t_key in node_pos:
                fx, fy, fw, fh, fc, fr = node_pos[f_key]
                tx, ty, tw, th, tc, tr = node_pos[t_key]
                elbl = str(ed.get("label") or "").strip()

                elbl_chars = sum(0.15 if ord(c) > 127 else 0.085 for c in elbl)

                # 1. 同行从左至右
                if fr == tr and fc < tc:
                    span_w = tx - (fx + fw)
                    arr_w = max(0.90, min(span_w - 0.06, elbl_chars + 0.52))
                    arr_x = fx + fw + (span_w - arr_w) / 2.0
                    arr_y = fy + fh / 2.0 - 0.13
                    cv.pill(arr_x, arr_y, arr_w, 0.26,
                            fill=pal.get("badge_bg"), line=pal.get("card_border"))
                    tf_e = cv.tbox(arr_x, arr_y, arr_w, 0.26, anchor=MSO_ANCHOR.MIDDLE)
                    arrow_text = f"{elbl} ➔" if elbl else "➔"
                    cv.para(tf_e, arrow_text, size=8.5, bold=True,
                            color=pal.get("accent"), align="center", first=True)

                # 2. 同行从右至左
                elif fr == tr and fc > tc:
                    span_w = fx - (tx + tw)
                    arr_w = max(0.90, min(span_w - 0.06, elbl_chars + 0.52))
                    arr_x = tx + tw + (span_w - arr_w) / 2.0
                    arr_y = fy + fh / 2.0 - 0.13
                    cv.pill(arr_x, arr_y, arr_w, 0.26,
                            fill=pal.get("badge_bg"), line=pal.get("card_border"))
                    tf_e = cv.tbox(arr_x, arr_y, arr_w, 0.26, anchor=MSO_ANCHOR.MIDDLE)
                    arrow_text = f"◀ {elbl}" if elbl else "◀"
                    cv.para(tf_e, arrow_text, size=8.5, bold=True,
                            color=pal.get("accent"), align="center", first=True)

                # 3. 同列从上至下
                elif fc == tc and fr < tr:
                    span_h = ty - (fy + fh)
                    arr_h = max(0.24, min(span_h - 0.08, 0.32))
                    arr_w = max(1.00, min(node_w - 0.30, elbl_chars + 0.50))
                    arr_x = fx + (fw - arr_w) / 2.0
                    arr_y = fy + fh + (span_h - arr_h) / 2.0
                    cv.pill(arr_x, arr_y, arr_w, arr_h,
                            fill=pal.get("badge_bg"), line=pal.get("card_border"))
                    tf_e = cv.tbox(arr_x, arr_y, arr_w, arr_h, anchor=MSO_ANCHOR.MIDDLE)
                    arrow_text = f"{elbl} ▼" if elbl else "▼"
                    cv.para(tf_e, arrow_text, size=8.5, bold=True,
                            color=pal.get("accent"), align="center", first=True)

                # 4. 同列从下至上
                elif fc == tc and fr > tr:
                    span_h = fy - (ty + th)
                    arr_h = max(0.24, min(span_h - 0.08, 0.32))
                    arr_w = max(1.00, min(node_w - 0.30, elbl_chars + 0.50))
                    arr_x = fx + (fw - arr_w) / 2.0
                    arr_y = ty + th + (span_h - arr_h) / 2.0
                    cv.pill(arr_x, arr_y, arr_w, arr_h,
                            fill=pal.get("badge_bg"), line=pal.get("card_border"))
                    tf_e = cv.tbox(arr_x, arr_y, arr_w, arr_h, anchor=MSO_ANCHOR.MIDDLE)
                    arrow_text = f"{elbl} ▲" if elbl else "▲"
                    cv.para(tf_e, arrow_text, size=8.5, bold=True,
                            color=pal.get("accent"), align="center", first=True)

                # 5. 跨层跨列流转（如右上至左下，或上排节点下沉至下排其它列）
                else:
                    cx1 = fx + fw / 2.0
                    cx2 = tx + tw / 2.0
                    arr_w = max(1.20, min(2.40, elbl_chars + 0.58))
                    mid_x = (cx1 + cx2) / 2.0 - arr_w / 2.0
                    mid_y = (fy + fh + ty) / 2.0 - 0.13
                    cv.pill(mid_x, mid_y, arr_w, 0.26,
                            fill=pal.get("badge_bg"), line=pal.get("card_border"))
                    tf_e = cv.tbox(mid_x, mid_y, arr_w, 0.26, anchor=MSO_ANCHOR.MIDDLE)
                    sym = "↙" if tc < fc else "↘"
                    arrow_text = f"{elbl} {sym}" if elbl else sym
                    cv.para(tf_e, arrow_text, size=8.5, bold=True,
                            color=pal.get("accent"), align="center", first=True)

    return y + h


def footer_note(cv, sh, text, y=None):
    """页面底部轻量脚注信息。"""
    pal = cv.pal
    y_pos = (sh - 0.42) if y is None else y
    cv.text(MARGIN, y_pos, CONTENT_W * 0.75, 0.24, str(text),
            size=9, color=pal.get("text_muted", pal["text_light"]),
            align="left", line=1.0)


# ==============================================================================
# 全新丰富专业版式库（团队、金字塔/漏斗、SWOT、路线图、利弊、FAQ、案例、定价、总结）
# ==============================================================================

def team_grid(cv, y, members, h=4.8):
    """团队与组织架构介绍页：3~4 位核心成员精美卡片。"""
    pal = cv.pal
    n = max(1, min(len(members), 4))
    gap = 0.28
    card_w = (CONTENT_W - gap * (n - 1)) / n
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]

    for i, m in enumerate(members[:n]):
        cx = MARGIN + i * (card_w + gap)
        cv.card(cx, y, card_w, h)
        acc = accents[i % len(accents)]

        # 顶部微装饰色带
        cv.rect(cx, y, card_w, 0.045, fill=acc)

        # 头像占位圆形（优雅字母徽章）
        avatar_r = 0.72
        av_x = cx + (card_w - avatar_r) / 2.0
        av_y = y + 0.30
        cv.circle(av_x, av_y, avatar_r, fill=pal.get("badge_bg"), line=acc, lw=1.2)
        av_tf = cv.tbox(av_x, av_y, avatar_r, avatar_r, anchor=MSO_ANCHOR.MIDDLE)
        init_char = (m.get("avatar_text") or m.get("name") or "人")[:2]
        cv.para(av_tf, init_char, size=13, bold=True, color=acc, align="center", first=True)

        # 姓名
        name = str(m.get("name", "核心骨干"))
        cv.text(cx + 0.15, y + 1.15, card_w - 0.30, 0.36, name,
                size=13.5, bold=True, color=pal.get("text_title", pal["primary"]), align="center")

        # 角色/职位胶囊
        role = str(m.get("role") or m.get("title") or "核心骨干")
        role_y = y + 1.56
        rw = min(card_w - 0.30, len(role) * 0.11 + 0.40)
        cv.pill(cx + (card_w - rw) / 2.0, role_y, rw, 0.28, fill=pal.get("card_subtle"), line=pal.get("card_border"))
        rtf = cv.tbox(cx + (card_w - rw) / 2.0, role_y, rw, 0.28, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(rtf, role, size=9.2, bold=True, color=acc, align="center", first=True)

        # 分割线
        cv.divider(cx + 0.20, y + 1.98, card_w - 0.40)

        # 履历与背景描述
        desc = m.get("desc") or m.get("bio") or ""
        if desc:
            tf_d = cv.tbox(cx + 0.18, y + 2.12, card_w - 0.36, 1.60)
            cv.para(tf_d, parse_rich(desc, color=acc), size=9.2, color=pal.get("text_body"), line=1.22, first=True)

        # 专长标签列表（底部分布）
        tags = m.get("tags") or m.get("skills") or []
        if tags:
            tag_y = y + h - 0.54
            tx_cur = cx + 0.16
            for t in tags[:3]:
                t_str = str(t)
                tw = len(t_str) * 0.11 + 0.26
                if tx_cur + tw <= cx + card_w - 0.12:
                    cv.pill(tx_cur, tag_y, tw, 0.24, fill=pal.get("card_subtle"))
                    ttf = cv.tbox(tx_cur, tag_y, tw, 0.24, anchor=MSO_ANCHOR.MIDDLE)
                    cv.para(ttf, t_str, size=8.2, color=pal.get("text_muted"), align="center", first=True)
                    tx_cur += tw + 0.10
    return y + h


def pyramid_funnel(cv, y, tiers, h=4.8):
    """战略金字塔 / 业务漏斗递进页：多层居中递进渐变色块。"""
    pal = cv.pal
    n = max(1, min(len(tiers), 5))
    tier_gap = 0.15
    tier_h = (h - tier_gap * (n - 1)) / n
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]

    min_w = 7.0
    max_w = CONTENT_W

    for i, t in enumerate(tiers[:n]):
        ratio = (i / float(n - 1)) if n > 1 else 0.5
        w = min_w + (max_w - min_w) * ratio
        tx = MARGIN + (CONTENT_W - w) / 2.0
        ty = y + i * (tier_h + tier_gap)
        acc = accents[i % len(accents)]

        # 背景卡片
        cv.rrect(tx, ty, w, tier_h, fill=pal.get("card_bg"), line=pal.get("card_border"), radius=0.08)
        # 左侧强调色边条
        cv.rect(tx, ty, 0.09, tier_h, fill=acc)

        # 层级层数徽章
        step_str = f"L{i+1}" if "L" not in str(t.get("level", "")) else str(t.get("level"))
        cv.pill(tx + 0.20, ty + (tier_h - 0.32) / 2.0, 0.56, 0.32, fill=acc)
        stf = cv.tbox(tx + 0.20, ty + (tier_h - 0.32) / 2.0, 0.56, 0.32, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(stf, step_str, size=10, bold=True, color=pal.get("card_bg", pal.get("white")), align="center", first=True)

        # 标题与重点
        title = str(t.get("title") or t.get("name") or f"层级 {i+1}")
        rate = t.get("metric") or t.get("rate") or ""
        left_text_w = (w - 2.8) if rate else (w - 1.1)

        tf_main = cv.tbox(tx + 0.90, ty + 0.08, left_text_w, tier_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_main, title, size=12, bold=True, color=pal.get("text_title", pal["primary"]), first=True)
        desc = t.get("desc") or t.get("subtitle") or ""
        if desc:
            cv.para(tf_main, parse_rich(desc, color=acc), size=9.2, color=pal.get("text_body", pal["text"]), before=2)

        # 右侧核心指标/占比（若有）
        if rate:
            rx = tx + w - 1.80
            rtf = cv.tbox(rx, ty + (tier_h - 0.45) / 2.0, 1.60, 0.45, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(rtf, str(rate), size=16, bold=True, color=acc, align="right", first=True)

    return y + h


def swot_matrix(cv, y, swot_data, h=4.8):
    """SWOT 战略态势分析矩阵：4 大象限专业构图。"""
    pal = cv.pal
    quad_w = (CONTENT_W - 0.26) / 2.0
    quad_h = (h - 0.26) / 2.0
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]

    configs = [
        ("S", "优势 · Strengths", swot_data.get("s") or swot_data.get("strengths") or [],
         MARGIN, y, accents[0]),
        ("W", "劣势 · Weaknesses", swot_data.get("w") or swot_data.get("weaknesses") or [],
         MARGIN + quad_w + 0.26, y, pal.get("accent2", accents[1 % len(accents)])),
        ("O", "机会 · Opportunities", swot_data.get("o") or swot_data.get("opportunities") or [],
         MARGIN, y + quad_h + 0.26, accents[2 % len(accents)]),
        ("T", "威胁 · Threats", swot_data.get("t") or swot_data.get("threats") or [],
         MARGIN + quad_w + 0.26, y + quad_h + 0.26, accents[3 % len(accents)]),
    ]

    for letter, label, items, qx, qy, col in configs:
        cv.card(qx, qy, quad_w, quad_h)
        # 顶部左侧色标
        cv.rect(qx, qy, 1.2, 0.045, fill=col)

        # 水印大字母背景
        cv.text(qx + quad_w - 1.10, qy + 0.12, 0.95, 0.95, letter,
                size=44, bold=True, color=pal.get("card_subtle"), align="right")

        # 象限标签徽章
        cv.badge(qx + 0.18, qy + 0.16, label, dot=True, dot_color=col,
                 text_color=col, size=10, h=0.30, pad_x=0.12)

        # 要点列表
        items_list = items if isinstance(items, list) else [items]
        tf = cv.tbox(qx + 0.20, qy + 0.58, quad_w - 0.40, quad_h - 0.68)
        for j, it in enumerate(items_list[:4]):
            it_text = it if isinstance(it, str) else it.get("text", "")
            runs = [("▪ ", {"b": True, "c": col, "sz": 9.5})] + parse_rich(it_text, color=col)
            cv.para(tf, runs, size=9.5, color=pal.get("text_body"),
                    first=(j == 0), before=4 if j > 0 else 0)

    return y + h


def roadmap_milestones(cv, y, phases, h=4.8):
    """战略演进路线图：3~4 个阶段跨周期里程碑与任务清单。"""
    pal = cv.pal
    n = max(1, min(len(phases), 4))
    gap = 0.24
    pw = (CONTENT_W - gap * (n - 1)) / n
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]

    for i, ph in enumerate(phases[:n]):
        px = MARGIN + i * (pw + gap)
        acc = accents[i % len(accents)]
        cv.card(px, y, pw, h)

        # 阶段时间药丸
        phase_tag = str(ph.get("phase") or ph.get("time") or f"PHASE {i+1}")
        cv.pill(px + 0.16, y + 0.18, len(phase_tag) * 0.10 + 0.40, 0.28, fill=acc)
        ptf = cv.tbox(px + 0.16, y + 0.18, len(phase_tag) * 0.10 + 0.40, 0.28, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(ptf, phase_tag, size=9.2, bold=True, color=pal.get("card_bg", pal.get("white")), align="center", first=True)

        # 状态微胶囊
        status = ph.get("status") or ""
        if status:
            cv.badge(px + pw - 0.95, y + 0.18, status, dot=False, size=8.5, h=0.26, pad_x=0.08)

        # 阶段核心主题
        title = str(ph.get("title") or f"阶段目标 {i+1}")
        cv.text(px + 0.16, y + 0.58, pw - 0.32, 0.44, title,
                size=12, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        # 战略目标概述
        goal = ph.get("goal") or ph.get("desc") or ""
        if goal:
            cv.text(px + 0.16, y + 1.08, pw - 0.32, 0.50, goal,
                    size=9, color=pal.get("text_body", pal["text"]), line=1.2)

        # 分割线
        cv.divider(px + 0.16, y + 1.66, pw - 0.32)

        # 里程碑任务矩阵
        tasks = ph.get("tasks") or ph.get("milestones") or ph.get("items") or []
        tf_tasks = cv.tbox(px + 0.16, y + 1.76, pw - 0.32, h - 1.90)
        for j, t in enumerate(tasks[:5]):
            t_text = t if isinstance(t, str) else t.get("text", "")
            runs = [("✔ ", {"b": True, "c": acc, "sz": 9.2})] + parse_rich(t_text, color=acc)
            cv.para(tf_tasks, runs, size=9.2, color=pal.get("text_body"),
                    first=(j == 0), before=4 if j > 0 else 0)

        # 卡片之间箭头指示流转推进
        if i < n - 1:
            arr_x = px + pw
            arr_tf = cv.tbox(arr_x, y + h / 2.0 - 0.20, gap, 0.40, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(arr_tf, "➔", size=13, bold=True, color=pal.get("card_border"), align="center", first=True)

    return y + h


def pros_cons(cv, y, pros_data, cons_data, takeaway=None, h=4.8):
    """方案利弊得失权衡页：左侧优势收益 vs 右侧风险成本。"""
    pal = cv.pal
    has_takeaway = bool(takeaway)
    main_h = (h - 1.15) if has_takeaway else h
    half_w = (CONTENT_W - 0.28) / 2.0

    acc_pros = pal.get("accent2", pal["accent"])
    acc_cons = pal.get("accent", pal["primary"])

    # 1. 左侧优势卡片 (Pros)
    cv.card(MARGIN, y, half_w, main_h)
    cv.rect(MARGIN, y, half_w, 0.05, fill=acc_pros)
    cv.badge(MARGIN + 0.20, y + 0.20, "＋ 方案核心优势与超额收益 (PROS)", dot=False,
             dot_color=acc_pros, text_color=acc_pros, size=10, h=0.32, pad_x=0.14)

    pros_items = pros_data if isinstance(pros_data, list) else (pros_data.get("items") or [pros_data])
    tf_p = cv.tbox(MARGIN + 0.22, y + 0.65, half_w - 0.44, main_h - 0.85)
    for j, it in enumerate(pros_items[:5]):
        t = it if isinstance(it, str) else it.get("text", "")
        runs = [("✔  ", {"b": True, "c": acc_pros, "sz": 9.8})] + parse_rich(t, color=acc_pros)
        cv.para(tf_p, runs, size=9.8, color=pal.get("text_body"), first=(j == 0), before=6 if j > 0 else 0)

    # 2. 右侧弊端/挑战卡片 (Cons)
    rx = MARGIN + half_w + 0.28
    cv.card(rx, y, half_w, main_h)
    cv.rect(rx, y, half_w, 0.05, fill=acc_cons)
    cv.badge(rx + 0.20, y + 0.20, "▲ 潜在风险与约束考量 (CONS)", dot=False,
             dot_color=acc_cons, text_color=acc_cons, size=10, h=0.32, pad_x=0.14)

    cons_items = cons_data if isinstance(cons_data, list) else (cons_data.get("items") or [cons_data])
    tf_c = cv.tbox(rx + 0.22, y + 0.65, half_w - 0.44, main_h - 0.85)
    for j, it in enumerate(cons_items[:5]):
        t = it if isinstance(it, str) else it.get("text", "")
        runs = [("▲  ", {"b": True, "c": acc_cons, "sz": 9.8})] + parse_rich(t, color=acc_cons)
        cv.para(tf_c, runs, size=9.8, color=pal.get("text_body"), first=(j == 0), before=6 if j > 0 else 0)

    # 3. 底部权衡决策卡
    if has_takeaway:
        takeaway_card(cv, takeaway, y + main_h + 0.20, title="权衡决策与应对策略", h=0.95)

    return y + h


def faq_accordion(cv, y, faqs, h=4.8):
    """常见疑虑与问题解答页：3~4 个高质感手风琴式 Q&A 块。"""
    pal = cv.pal
    n = max(1, min(len(faqs), 4))
    gap = 0.20
    item_h = (h - gap * (n - 1)) / n
    acc = pal.get("accent")

    for i, f in enumerate(faqs[:n]):
        iy = y + i * (item_h + gap)
        cv.card(MARGIN, iy, CONTENT_W, item_h)

        q_text = str(f.get("q") or f.get("question") or f.get("title") or "核心疑问")
        tag = f.get("tag") or f.get("category")

        cv.pill(MARGIN + 0.16, iy + 0.14, 0.45, 0.28, fill=pal.get("badge_bg"))
        qtf = cv.tbox(MARGIN + 0.16, iy + 0.14, 0.45, 0.28, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(qtf, "Q", size=11, bold=True, color=acc, align="center", first=True)

        cv.text(MARGIN + 0.70, iy + 0.14, CONTENT_W - 2.2, 0.28, q_text,
                size=11.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        if tag:
            cv.badge(MARGIN + CONTENT_W - 1.40, iy + 0.14, str(tag), dot=False, size=8.5, h=0.26)

        cv.divider(MARGIN + 0.16, iy + 0.48, CONTENT_W - 0.32)

        a_text = str(f.get("a") or f.get("answer") or f.get("desc") or "")
        tf_a = cv.tbox(MARGIN + 0.20, iy + 0.54, CONTENT_W - 0.40, item_h - 0.60)
        cv.para(tf_a, parse_rich(a_text, color=acc), size=9.6, color=pal.get("text_body"), line=1.22, first=True)

    return y + h


def case_study(cv, y, case_data, h=4.8):
    """标杆客户案例故事页：背景痛点 + 解决方案 + 核心 ROI 成果。"""
    pal = cv.pal
    left_w = 4.8
    right_w = CONTENT_W - left_w - 0.30
    right_x = MARGIN + left_w + 0.30

    acc = pal.get("accent")
    acc2 = pal.get("accent2", pal["primary"])

    # 左侧：客户画像与痛点挑战
    cv.card(MARGIN, y, left_w, h)
    cv.rect(MARGIN, y, left_w, 0.045, fill=acc)

    client_name = str(case_data.get("client") or case_data.get("company") or "标杆客户案例")
    industry = str(case_data.get("industry") or "商业数字化转型")
    cv.badge(MARGIN + 0.20, y + 0.20, industry, dot=True, size=9.5, h=0.30)
    cv.text(MARGIN + 0.20, y + 0.60, left_w - 0.40, 0.40, client_name,
            size=14, bold=True, color=pal.get("text_title", pal["primary"]))

    cv.divider(MARGIN + 0.20, y + 1.10, left_w - 0.40)

    # 业务痛点
    challenge = case_data.get("challenge") or case_data.get("pain_points") or ""
    if challenge:
        cv.text(MARGIN + 0.20, y + 1.25, left_w - 0.40, 0.28, "【业务挑战与痛点】",
                size=10.5, bold=True, color=acc)
        tf_ch = cv.tbox(MARGIN + 0.20, y + 1.58, left_w - 0.40, 1.40)
        cv.para(tf_ch, parse_rich(challenge, color=acc), size=9.5, color=pal.get("text_body"), line=1.25, first=True)

    # 落地方案
    solution = case_data.get("solution") or case_data.get("approach") or ""
    if solution:
        sol_y = y + 3.05
        cv.text(MARGIN + 0.20, sol_y, left_w - 0.40, 0.28, "【关键实施举措】",
                size=10.5, bold=True, color=acc2)
        tf_sol = cv.tbox(MARGIN + 0.20, sol_y + 0.33, left_w - 0.40, h - (sol_y - y) - 0.45)
        cv.para(tf_sol, parse_rich(solution, color=acc2), size=9.5, color=pal.get("text_body"), line=1.25, first=True)

    # 右侧：量化 ROI 数据与客户证言
    metrics = case_data.get("results") or case_data.get("metrics") or []
    top_metrics_h = 2.40
    cv.card(right_x, y, right_w, top_metrics_h)
    cv.badge(right_x + 0.20, y + 0.20, "● 核心量化价值与 ROI 成效", dot=False, size=10, h=0.30)

    m_count = max(1, min(len(metrics), 3))
    m_gap = 0.20
    mw = (right_w - 0.40 - m_gap * (m_count - 1)) / m_count
    for i, m in enumerate(metrics[:m_count]):
        mx = right_x + 0.20 + i * (mw + m_gap)
        my = y + 0.65
        cv.rrect(mx, my, mw, 1.50, fill=pal.get("card_subtle"), line=pal.get("card_border"))
        val = str(m.get("metric") or m.get("val") or m.get("value") or "100%")
        lbl = str(m.get("label") or m.get("name") or "提升指标")
        sub = str(m.get("sub") or m.get("desc") or "")

        tf_m = cv.tbox(mx + 0.10, my + 0.12, mw - 0.20, 1.30, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_m, val, size=24, bold=True, color=acc, align="center", first=True)
        cv.para(tf_m, lbl, size=10, bold=True, color=pal.get("text_title", pal["primary"]), align="center", before=2)
        if sub:
            cv.para(tf_m, sub, size=8.5, color=pal.get("text_muted"), align="center", before=2)

    # 右侧下部：客户证言 / 评价
    bot_y = y + top_metrics_h + 0.25
    bot_h = h - top_metrics_h - 0.25
    cv.card(right_x, bot_y, right_w, bot_h)
    quote = case_data.get("quote") or case_data.get("testimonial") or "通过本方案的高效平稳实施，极大提升了核心系统的敏捷响应能力与综合服务质量。"
    author = case_data.get("quote_author") or case_data.get("author") or "客户业务负责人"

    cv.text(right_x + 0.25, bot_y + 0.20, 0.40, 0.40, "“", size=32, bold=True, color=acc)
    tf_q = cv.tbox(right_x + 0.60, bot_y + 0.22, right_w - 0.85, bot_h - 0.55)
    cv.para(tf_q, quote, size=10.5, color=pal.get("text_body"), line=1.3, first=True)
    cv.para(tf_q, f"— {author}", size=9.5, bold=True, color=pal.get("text_subtitle", pal["primary"]), align="right", before=6)

    return y + h


def pricing_packages(cv, y, packages, h=4.8):
    """产品套餐与商业定价对比页：3 档权益卡片与高光推荐卡。"""
    pal = cv.pal
    n = max(1, min(len(packages), 3))
    gap = 0.28
    pw = (CONTENT_W - gap * (n - 1)) / n
    acc = pal.get("accent")

    for i, p in enumerate(packages[:n]):
        px = MARGIN + i * (pw + gap)
        is_pop = p.get("is_popular", False) or p.get("highlight", False) or (i == 1 and n == 3)
        acc_card = acc if is_pop else pal.get("primary")

        border_col = acc if is_pop else pal.get("card_border")
        lw = 1.6 if is_pop else 0.75
        cv.card(px, y, pw, h, line=border_col, lw=lw)

        if is_pop:
            cv.pill(px + (pw - 1.20) / 2.0, y - 0.16, 1.20, 0.32, fill=acc)
            ptf = cv.tbox(px + (pw - 1.20) / 2.0, y - 0.16, 1.20, 0.32, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(ptf, "★ 强烈推荐", size=9.5, bold=True, color=pal.get("card_bg", pal.get("white")), align="center", first=True)

        p_name = str(p.get("name") or f"套餐档位 {i+1}")
        cv.text(px + 0.15, y + 0.30, pw - 0.30, 0.35, p_name,
                size=13.5, bold=True, color=pal.get("text_title", pal["primary"]), align="center")

        price = str(p.get("price") or "¥ 9,800")
        period = str(p.get("period") or p.get("unit") or "/ 年")
        tf_pr = cv.tbox(px + 0.10, y + 0.72, pw - 0.20, 0.55, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_pr, [
            (price, {"sz": 20, "b": True, "c": acc_card}),
            (f" {period}", {"sz": 10, "c": pal.get("text_muted")}),
        ], align="center", first=True)

        desc = p.get("desc") or ""
        if desc:
            cv.text(px + 0.15, y + 1.30, pw - 0.30, 0.36, desc,
                    size=9, color=pal.get("text_body", pal["text"]), align="center", line=1.1)

        cv.divider(px + 0.20, y + 1.72, pw - 0.40)

        features = p.get("features") or p.get("items") or []
        tf_f = cv.tbox(px + 0.20, y + 1.85, pw - 0.40, h - 2.50)
        for j, feat in enumerate(features[:6]):
            runs = [("✔ ", {"b": True, "c": acc_card, "sz": 9.5})] + parse_rich(str(feat), color=acc_card)
            cv.para(tf_f, runs, size=9.5, color=pal.get("text_body"), first=(j == 0), before=4 if j > 0 else 0)

        cta_text = str(p.get("cta") or "立即开通体验")
        cta_y = y + h - 0.52
        cta_w = pw - 0.60
        btn_bg = acc if is_pop else pal.get("card_subtle")
        btn_fg = pal.get("card_bg", pal.get("white")) if is_pop else pal.get("text_title", pal["primary"])
        cv.pill(px + 0.30, cta_y, cta_w, 0.36, fill=btn_bg, line=pal.get("card_border") if not is_pop else None)
        ctf = cv.tbox(px + 0.30, cta_y, cta_w, 0.36, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(ctf, cta_text, size=10, bold=True, color=btn_fg, align="center", first=True)

    return y + h


def summary_next_steps(cv, y, summary_data, h=4.8):
    """高管总结与下一步行动计划页：左侧核心结论 + 右侧落地行动清单。"""
    pal = cv.pal
    left_w = 4.8
    right_w = CONTENT_W - left_w - 0.28
    right_x = MARGIN + left_w + 0.28
    acc = pal.get("accent")

    # 左侧：核心汇报结论沉淀 (Key Takeaways)
    cv.card(MARGIN, y, left_w, h)
    cv.rect(MARGIN, y, left_w, 0.045, fill=acc)
    cv.badge(MARGIN + 0.20, y + 0.20, "★ 战略核心结论沉淀", dot=False, size=10, h=0.30)

    takeaways = summary_data.get("takeaways") or summary_data.get("conclusions") or []
    if isinstance(takeaways, str):
        takeaways = [takeaways]
    tf_tk = cv.tbox(MARGIN + 0.20, y + 0.65, left_w - 0.40, h - 0.85)
    for j, tk in enumerate(takeaways[:4]):
        runs = [(f"0{j+1} · ", {"b": True, "c": acc, "sz": 10.5})] + parse_rich(str(tk), color=acc)
        cv.para(tf_tk, runs, size=10, color=pal.get("text_body"), first=(j == 0), before=8 if j > 0 else 0)

    # 右侧：下一步行动方案卡片 (Action Plan)
    cv.card(right_x, y, right_w, h)
    cv.badge(right_x + 0.20, y + 0.20, "下一步落地行动计划 (Next Steps)", dot=True, size=10, h=0.30)

    actions = summary_data.get("actions") or summary_data.get("next_steps") or []
    act_count = max(1, min(len(actions), 4))
    act_gap = 0.16
    act_h = (h - 0.80 - act_gap * (act_count - 1)) / act_count

    for i, act in enumerate(actions[:act_count]):
        ay = y + 0.62 + i * (act_h + act_gap)
        cv.rrect(right_x + 0.18, ay, right_w - 0.36, act_h, fill=pal.get("card_subtle"), line=pal.get("card_border"))
        task_str = str(act.get("task") or act.get("title") or f"行动计划 {i+1}")
        owner_str = str(act.get("owner") or act.get("team") or "项目组")
        due_str = str(act.get("due") or act.get("deadline") or "近期交付")

        tf_act = cv.tbox(right_x + 0.30, ay + 0.10, right_w - 2.4, act_h - 0.20, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_act, parse_rich(task_str, color=acc), size=10.5, bold=True, color=pal.get("text_title", pal["primary"]), first=True)
        desc_str = act.get("desc") or ""
        if desc_str:
            cv.para(tf_act, parse_rich(desc_str, color=acc), size=8.5, color=pal.get("text_subtitle", pal["text_body"]), before=2)

        tf_meta = cv.tbox(right_x + right_w - 2.10, ay + 0.10, 1.80, act_h - 0.20, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_meta, owner_str, size=9.5, bold=True, color=acc, align="right", first=True)
        cv.para(tf_meta, due_str, size=8.5, color=pal.get("text_subtitle", pal["text_body"]), align="right", before=2)

    return y + h

