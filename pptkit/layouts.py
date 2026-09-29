# -*- coding: utf-8 -*-
"""现代化版式构件库：页眉系统 / 导读摘要 / Bento KPI 卡 / 详情面板 / 2x2网格 / 对比卡 / 时间线 / 表格 / 页脚。

全面采用现代化网格系统（Grid System）与呼吸感留白，
自适应浅色商务、深色极客、自然生态等各类主题，杜绝大片留白与空洞。
"""
import re
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt
from .charts import bar_chart, render_chart


MARGIN = 0.65
CONTENT_W = 12.033

_TAG_RE = re.compile(r"\[\[(.*?)\]\]|\*\*(.*?)\*\*")
_PREFIX_BRACKET_RE = re.compile(r"^(【[^】\n]+】[:：]?)\s*(.*)$")
_PREFIX_COLON_RE = re.compile(r"^([^:：\n]{2,8}[:：])\s*(.+)$")


def parse_rich(text, color=None, bold=True):
    """将 '普通[[强调]]文本'、'普通**粗体**文本' 或自然前缀 '【标题】：描述' 解析为 Canvas.para 混排 runs。"""
    if isinstance(text, (list, tuple)):
        return text
    text = "" if text is None else str(text)

    # 1. 若包含明确的 [[高亮]] 或 **粗体** 标记，按显式标记解析
    if _TAG_RE.search(text):
        runs, pos = [], 0
        for m in _TAG_RE.finditer(text):
            if m.start() > pos:
                runs.append((text[pos:m.start()], {}))
            opt = {"b": bold}
            if color is not None:
                opt["c"] = color
            highlight_text = m.group(1) if m.group(1) is not None else m.group(2)
            runs.append((highlight_text, opt))
            pos = m.end()
        if pos < len(text):
            runs.append((text[pos:], {}))
        return runs or [("", {})]

    # 2. 若无显式标记，智能识别高频结构化标题语法：如 '【核心考点】：描述' 或 '意象之美: 描述'
    mb = _PREFIX_BRACKET_RE.match(text)
    if mb:
        head, rest = mb.group(1), mb.group(2)
        return [(head, {"b": True, "c": color} if color else {"b": True}), (rest, {})]

    mc = _PREFIX_COLON_RE.match(text)
    if mc:
        head, rest = mc.group(1), mc.group(2)
        if not re.match(r"^\d+[:：]|^(http|https|ftp)[:：]", head, re.I):
            return [(head, {"b": True, "c": color} if color else {"b": True}), (rest, {})]

    return [(text, {})]


def chrome(cv, sw, sh, thick=0.04):
    """现代化顶部微光呼吸条与空间环境光影（极简高级，彻底消除单调死板平铺）。"""
    pal = cv.pal
    is_dark = pal.get("is_dark", False)

    # 1. 空间环境光影层（Ambient Lighting Aura）：营造景深与空气感
    if is_dark:
        # 右上角柔和深邃极光光晕（双层重叠，营造空间深度）
        glow_fill = pal.get("primary_tint", pal["card_subtle"])
        cv.circle(sw - 3.8, -1.6, 5.2, fill=glow_fill)
        cv.circle(sw - 2.2, -0.6, 3.4, fill=pal.get("accent_tint", glow_fill))
        # 左下角次级微光晕
        cv.circle(-1.8, sh - 2.8, 4.6, fill=pal.get("bg_alt", pal["card_subtle"]))
        # 极简科技十字标（微光点缀）
        cv.text(MARGIN, 0.12, 0.4, 0.2, "＋", size=8.5, color=pal.get("card_border"))
        cv.text(sw - MARGIN - 0.25, 0.12, 0.4, 0.2, "＋", size=8.5, color=pal.get("card_border"))
    else:
        # 浅色商务背景：柔和顶角微晕
        cv.circle(sw - 3.2, -1.8, 4.8, fill=pal.get("primary_tint", pal["bg_alt"]))
        cv.circle(-2.0, sh - 2.5, 4.0, fill=pal.get("bg_alt", pal["card_subtle"]))

    # 极简科技微点阵（Tech Dot Matrix Pattern，右上角呼吸感纹理）
    dot_color = pal.get("divider", pal.get("card_border"))
    for row_i in range(3):
        for col_i in range(5):
            cv.circle(sw - 2.2 + col_i * 0.26, 0.22 + row_i * 0.20, 0.022, fill=dot_color)

    # 2. 顶部微光发光条与底边边线
    cv.rect(0, 0, sw, thick, fill=pal.get("primary", pal["primary"]))
    cv.rect(0, 0, 2.4, thick + 0.015, fill=pal.get("accent", pal["accent2"]))
    if is_dark:
        cv.rect(0, sh - 0.02, sw, 0.02, fill=pal.get("card_border"))


def header(cv, title, subtitle=None, category=None, y=0.38, size=23):
    """现代化幻灯片主页眉系统（支持长标题自适应缩放与呼吸感留白，杜绝孤字折行）。"""
    pal = cv.pal
    title_col = pal.get("text_title", pal["primary"])
    sub_col = pal.get("text_subtitle", pal["text_body"])

    title_str = str(title or "")
    # 动态字号自适应：长标题自动平滑微调字号，确保单行优雅展示
    if len(title_str) >= 28:
        actual_size = min(size, 18.5)
    elif len(title_str) >= 22:
        actual_size = min(size, 20.0)
    elif len(title_str) >= 18:
        actual_size = min(size, 21.5)
    else:
        actual_size = size

    # 计算副标题与主标题宽度分配
    if subtitle:
        sub_text = str(subtitle)
        char_count = sum(2 if ord(c) > 127 else 1 for c in sub_text)
        sub_w = min(5.2, max(2.6, char_count * 0.085 + 0.4))
        title_w = CONTENT_W - sub_w - 0.30
        cv.text(MARGIN + title_w + 0.30, y + 0.14, sub_w, 0.36, sub_text, size=11.5,
                color=sub_col, align="right", line=1.1)
    else:
        title_w = CONTENT_W

    # 1. 主标题
    cv.text(MARGIN, y, title_w, 0.54, title_str, size=actual_size, bold=True,
            color=title_col, line=1.05)

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

        box_w = cw - 0.40
        total_text = (val_str + (" " + unit_str if unit_str else "")).strip()
        char_len = sum(1.8 if ord(ch) > 127 else 1.0 for ch in total_text)

        if caps:
            base_sz = 28 if h <= 1.8 else 34
            # 严格根据卡片宽度与文本长度自适应缩放字号，杜绝折行与穿模
            if char_len > 6:
                fit_sz = int((box_w * 72.0) / (char_len * 0.65))
                num_sz = max(18, min(base_sz, fit_sz))
            else:
                num_sz = base_sz

            unit_sz = max(11, int(num_sz * 0.48))
            num_h = 0.46 if num_sz <= 26 else (0.54 if h <= 1.8 else 0.62)
            num_y = y + (0.46 if h <= 1.8 else 0.58)

            num_runs = [(val_str, {"b": True, "c": card_accent, "sz": num_sz})]
            if unit_str:
                num_runs.append((" " + unit_str, {"b": True, "c": card_accent, "sz": unit_sz}))
            cv.text(cx + 0.20, num_y, box_w, num_h, num_runs, line=1.0)

            div_y = num_y + num_h + 0.04
            cv.divider(cx + 0.20, div_y, box_w)

            caps_y = div_y + 0.08
            caps_h = max(0.40, (y + h) - caps_y - 0.10)
            tf = cv.tbox(cx + 0.20, caps_y, box_w, caps_h)
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
    b_end = cv.badge(x + 0.16, y + (h - 0.28) / 2.0, badge_str,
                     dot=False, bg=pal.get("badge_bg"), text_color=acc,
                     size=9, h=0.28, pad_x=0.08)

    text_x = b_end + 0.14
    text_w = max(0.5, (x + w - 0.16) - text_x)
    clean_text = format_item_text(text)

    # 当微卡片垂直空间充足（h >= 0.95）且包含 "标题: 描述" 时，智能拆分为标题+详述，实现充盈饱满排布
    has_colon = (":" in clean_text or "：" in clean_text)
    if h >= 0.95 and has_colon:
        parts = [p.strip() for p in clean_text.replace("：", ":").split(":", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            header, body = parts
            cv.text(text_x, y + 0.14, text_w, 0.32, header, size=size + 1.0, bold=True,
                    color=pal.get("text_title", pal["primary"]), line=1.1)
            tf_b = cv.tbox(text_x, y + 0.46, text_w, h - 0.56)
            runs_b = parse_rich(body, color=acc)
            cv.para(tf_b, runs_b, size=size, bold=False,
                    color=pal.get("text_body", pal["text"]), line=1.28, first=True)
            return

    tf = cv.tbox(text_x, y, text_w, h, anchor=MSO_ANCHOR.MIDDLE)
    runs = parse_rich(clean_text, color=acc)
    cv.para(tf, runs, size=size, bold=False,
            color=pal.get("text_body", pal["text"]), line=1.22, first=True)


def bullet_list(cv, x, y, w, h, items, size=11, gap=7, line=1.28, dot_color=None):
    """优雅的项目要点列表，自动支持结构化字典数据与动态垂直均分间距。"""
    n = len(items)
    if n == 0:
        return None
    # 动态垂直均分间距：当要点较少（如 2~3 项）且可用高度充裕时，自动拉大间距与字号，充盈版面
    dyn_size = size
    dyn_gap = gap
    dyn_line = line
    if h >= 1.8 and n <= 4:
        avail_pt = h * 72.0
        per_item_pt = dyn_size * 2 * dyn_line
        free_pt = avail_pt - n * per_item_pt
        if free_pt > 0:
            dyn_gap = max(gap, min(36, int(free_pt / max(1, n - 1) * 0.48)))
            if n <= 3 and size <= 11.5:
                dyn_size = min(13.0, size + 1.5)
                dyn_line = min(1.40, line + 0.10)

    anchor_mode = MSO_ANCHOR.MIDDLE if (n <= 3 and h >= 2.4) else MSO_ANCHOR.TOP
    tf = cv.tbox(x, y, w, h, anchor=anchor_mode)
    for i, it in enumerate(items):
        clean_text = format_item_text(it)
        runs = parse_rich(clean_text, color=dot_color or cv.pal.get("accent"))
        cv.bullet(tf, runs, first=(i == 0), size=dyn_size,
                  before=(0 if i == 0 else dyn_gap), line=dyn_line, dot=dot_color)
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

    card_h = h or (2.4 if rows_cnt == 2 else (4.6 if rows_cnt == 1 else 1.8))
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
        cv.pill(cx + 0.22, cy + 0.05, min(1.30, cw - 0.44), 0.035, fill=card_accent)

        # 卡片头部：标号徽章 + 标题
        idx_str = "%02d" % (i + 1)
        b_end = cv.badge(cx + 0.22, cy + 0.18, idx_str, dot=False,
                         bg=pal.get("badge_bg"), text_color=card_accent,
                         size=9.5, h=0.28, pad_x=0.10)

        title_str = str(c.get("title", ""))
        tag_str = c.get("tag") or c.get("metric") or c.get("label") or c.get("badge")
        tag_w = 1.35 if tag_str else 0.0
        title_x = b_end + 0.14
        title_w = max(0.5, cw - (title_x - cx) - tag_w - 0.20)
        title_sz = 13.5 if (cols <= 2 or len(title_str) <= 10) else (12.0 if len(title_str) <= 16 else 10.5)
        cv.text(title_x, cy + 0.15, title_w, 0.36, title_str,
                size=title_sz, bold=True, color=pal.get("text_title", pal["primary"]),
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
            anchor_mode = MSO_ANCHOR.MIDDLE if (content_h >= 2.0 and len(str(desc)) <= 120) else MSO_ANCHOR.TOP
            tf = cv.tbox(cx + 0.22, content_y, cw - 0.44, content_h, anchor=anchor_mode)
            runs = parse_rich(desc, color=card_accent)
            desc_sz = 13.0 if (content_h >= 2.5 and len(str(desc)) <= 40) else 11.0
            cv.para(tf, runs, size=desc_sz, color=pal.get("text_body", pal["text"]),
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
    left_desc = left_card.get("desc") or left_card.get("summary")
    if left_items:
        bullet_list(cv, lx + 0.30, y + 1.25, cw - 0.60, h - 1.45,
                    left_items, size=11, gap=8, line=1.3, dot_color=left_accent)
    elif left_desc:
        anchor_mode = MSO_ANCHOR.MIDDLE if (h >= 2.5 and len(str(left_desc)) <= 120) else MSO_ANCHOR.TOP
        tf = cv.tbox(lx + 0.30, y + 1.25, cw - 0.60, h - 1.45, anchor=anchor_mode)
        runs = parse_rich(left_desc, color=left_accent)
        desc_sz = 13.0 if (h >= 3.0 and len(str(left_desc)) <= 45) else 11.5
        cv.para(tf, runs, size=desc_sz, color=pal.get("text_body", pal["text"]),
                line=1.35, first=True)

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
    right_desc = right_card.get("desc") or right_card.get("summary")
    if right_items:
        bullet_list(cv, rx + 0.30, y + 1.25, cw - 0.60, h - 1.45,
                    right_items, size=11, gap=8, line=1.3, dot_color=right_accent)
    elif right_desc:
        anchor_mode = MSO_ANCHOR.MIDDLE if (h >= 2.5 and len(str(right_desc)) <= 120) else MSO_ANCHOR.TOP
        tf = cv.tbox(rx + 0.30, y + 1.25, cw - 0.60, h - 1.45, anchor=anchor_mode)
        runs = parse_rich(right_desc, color=right_accent)
        desc_sz = 13.0 if (h >= 3.0 and len(str(right_desc)) <= 45) else 11.5
        cv.para(tf, runs, size=desc_sz, color=pal.get("text_body", pal["text"]),
                line=1.35, first=True)


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

    pts = points or []
    if not pts:
        # 当无分支要点时，核心大字论断全卡居中满屏显示，气势磅礴
        tf = cv.tbox(MARGIN + 0.80, y + 0.60, CONTENT_W - 1.60, h - 1.20, anchor=MSO_ANCHOR.MIDDLE)
        runs = parse_rich(statement, color=acc, bold=True)
        cv.para(tf, runs, size=28, bold=True,
                color=pal.get("text_title", pal["primary"]), align="center", line=1.45, first=True)
        return y + h

    # 核心大字论断（Hero Typography，20~24pt 震撼高对比）
    stmt_len = sum(1.0 if ord(c) > 127 else 0.52 for c in str(statement))
    stmt_sz = 24 if stmt_len <= 35 else (20 if stmt_len <= 65 else 16.5)
    tf = cv.tbox(MARGIN + 0.50, y + 0.35, CONTENT_W - 1.0, 1.55, anchor=MSO_ANCHOR.MIDDLE)
    runs = parse_rich(statement, color=acc, bold=True)
    cv.para(tf, runs, size=stmt_sz, bold=True,
            color=pal.get("text_title", pal["primary"]), align="center", line=1.30, first=True)

    # 中部微细分割线
    cv.divider(MARGIN + 0.80, y + 2.05, CONTENT_W - 1.60)

    # 下方 2~4 个关键支柱徽章卡片（自适应高度与间距）
    n_p = min(len(pts), 4)
    gap = 0.25
    card_w = (CONTENT_W - 1.0 - gap * (n_p - 1)) / float(n_p)
    card_h = h - 2.40
    card_y = y + 2.20

    for i, pt in enumerate(pts[:n_p]):
        cx = MARGIN + 0.50 + i * (card_w + gap)
        pt_acc = accents[i % len(accents)]
        cv.card(cx, card_y, card_w, card_h, fill=pal.get("card_subtle"),
                line=pt_acc, lw=0.6, radius=0.08)

        idx_str = "%02d" % (i + 1)
        cv.badge(cx + 0.16, card_y + 0.16, idx_str, dot=False,
                 bg=pal.get("badge_bg"), text_color=pt_acc, size=9, h=0.26)

        pt_text = format_item_text(pt)
        has_colon = (":" in pt_text or "：" in pt_text)
        if card_h >= 1.6 and has_colon:
            parts = [p.strip() for p in pt_text.replace("：", ":").split(":", 1)]
            if len(parts) == 2 and parts[0] and parts[1]:
                cv.text(cx + 0.16, card_y + 0.48, card_w - 0.32, 0.30, parts[0],
                        size=11.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)
                tf_body = cv.tbox(cx + 0.16, card_y + 0.82, card_w - 0.32, card_h - 0.94)
                cv.para(tf_body, parse_rich(parts[1], color=pt_acc), size=10,
                        color=pal.get("text_body", pal["text"]), line=1.24, first=True)
                continue

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
        idx_str = str(s.get("step")) if s.get("step") else ("STEP %02d" % (i + 1))
        cv.badge(cur_x + 0.20, y + 0.18, idx_str, dot=True,
                 dot_color=step_acc, text_color=step_acc, size=9, h=0.28)

        # 步骤主标题（兼容 title/name/label）
        title_str = str(s.get("title") or s.get("name") or s.get("label") or "")
        cv.text(cur_x + 0.20, y + 0.54, card_w - 0.40, 0.36, title_str,
                size=12.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        cv.divider(cur_x + 0.20, y + 0.94, card_w - 0.40)

        # 步骤详述或要点
        desc_str = s.get("desc") or s.get("summary")
        items = s.get("items") or s.get("bullets")
        owner_str = s.get("owner") or s.get("role") or s.get("tag")
        content_y = y + 1.06
        content_h = (h - 1.60) if owner_str else (h - 1.20)

        if items:
            bullet_list(cv, cur_x + 0.20, content_y, card_w - 0.40, content_h,
                        items, size=10, gap=5, line=1.22, dot_color=step_acc)
        elif desc_str:
            tf = cv.tbox(cur_x + 0.20, content_y, card_w - 0.40, content_h)
            runs = parse_rich(desc_str, color=step_acc)
            cv.para(tf, runs, size=10.5, color=pal.get("text_body", pal["text"]),
                    line=1.30, first=True)

        if owner_str:
            cv.badge(cur_x + 0.20, y + h - 0.44, str(owner_str),
                     bg=pal.get("badge_bg", pal["primary_tint"]),
                     text_color=step_acc, size=8.5, h=0.26)

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
                      categories=None, series=None, chart_h=1.35, accent_color=None,
                      chart_spec=None):
    """复合图表面板：上部要点文字 + 下部现代图表（支持柱状图、折线图、面积图、饼图）。"""
    px, by, bw = panel(cv, x, y, w, h, title, subtitle, accent_color=accent_color)
    text_h = max(0.5, h - (by - y) - chart_h - 0.20)
    if items:
        bullet_list(cv, px, by, bw, text_h, items, size=10, gap=5, line=1.2, dot_color=accent_color)
    chart_y = y + h - chart_h - 0.15
    if chart_spec and isinstance(chart_spec, dict):
        render_chart(cv, px + 0.05, chart_y, bw - 0.10, chart_h, chart_spec)
    elif categories and series:
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
        return y

    step = w / float(n)
    line_y = y + 0.48

    # 贯通连接线
    cv.rect(x + step * 0.25, line_y, w - step * 0.5, 0.025,
            fill=pal.get("card_border", pal["border"]))

    card_h = h or 3.40
    card_y = line_y + 0.26

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
        cv.card(cx + 0.10, card_y, card_w, card_h, radius=0.08)

        # 卡片顶部发光微条
        cv.pill(cx + 0.24, card_y + 0.04, min(1.0, card_w - 0.48), 0.03, fill=step_accent)

        # 卡片内标题
        title_str = str(it.get("title", ""))
        title_sz = 13.5 if len(title_str) <= 12 else 11.5
        cv.text(cx + 0.24, card_y + 0.20, card_w - 0.48, 0.36,
                title_str, size=title_sz, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.1)

        cv.divider(cx + 0.24, card_y + 0.60, card_w - 0.48)

        items_sub = it.get("items") or it.get("bullets")
        desc = it.get("desc") or it.get("subtitle") or it.get("text")
        content_y = card_y + 0.72
        content_h = card_h - 0.85

        if items_sub:
            bullet_list(cv, cx + 0.24, content_y, card_w - 0.48, content_h,
                        items_sub, size=10, gap=5, line=1.22, dot_color=step_accent)
        elif desc:
            anchor_mode = MSO_ANCHOR.MIDDLE if (content_h >= 1.8 and len(str(desc)) <= 80) else MSO_ANCHOR.TOP
            tf = cv.tbox(cx + 0.24, content_y, card_w - 0.48, content_h, anchor=anchor_mode)
            runs = parse_rich(desc, color=step_accent)
            desc_sz = 11.5 if len(str(desc)) <= 40 else 10.5
            cv.para(tf, runs, size=desc_sz,
                    color=pal.get("text_body", pal["text"]), line=1.35, first=True)

    return card_y + card_h


def table(cv, x, y, w, rows, col_widths=None, header=True,
          row_h=0.48, size=11, header_size=11.5):
    """现代化卡片式表格：圆角表头、平滑行高、微光斑马纹、智能自适应列宽。"""
    pal = cv.pal
    if not rows:
        return y
    ncol = len(rows[0])
    if col_widths:
        total = float(sum(col_widths))
        cws = [w * cw / total for cw in col_widths]
    else:
        # 智能根据各列最大字符长度自适应推算列宽比例，杜绝短列巨空、长列挤爆
        max_lens = [0] * ncol
        for r in rows:
            for ci, cell in enumerate(r[:ncol]):
                c_str = str(cell)
                c_len = sum(2 if ord(ch) > 127 else 1 for ch in c_str)
                if c_len > max_lens[ci]:
                    max_lens[ci] = c_len
        # 赋予平滑保底权重（避免极短列过窄，长列不过度吃满）
        col_weights = [max(8, min(45, l)) for l in max_lens]
        tot_wt = float(sum(col_weights)) or float(ncol)
        cws = [w * (wt / tot_wt) for wt in col_weights]
        min_w = 0.85
        if any(cw < min_w for cw in cws) and w >= min_w * ncol:
            cws = [max(min_w, cw) for cw in cws]
            scale = w / sum(cws)
            cws = [cw * scale for cw in cws]

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


def takeaway_card(cv, text, y, title="核心洞察与要点分析", h=None, accent_color=None):
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
        rate = str(t.get("metric") or t.get("rate") or "").strip()
        rate_len = sum(1.8 if ord(ch) > 127 else 1.0 for ch in rate) if rate else 0

        # 根据指标文本长度自适应分配合理的右侧宽度与字号
        if rate:
            metric_box_w = max(1.8, min(3.2, rate_len * 0.16 + 0.50))
            left_text_w = max(3.0, w - metric_box_w - 1.20)
            metric_sz = 16 if rate_len <= 5 else (13.5 if rate_len <= 10 else 12)
        else:
            metric_box_w = 0.0
            left_text_w = w - 1.15
            metric_sz = 14

        tf_main = cv.tbox(tx + 0.90, ty + 0.08, left_text_w, tier_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_main, title, size=12, bold=True, color=pal.get("text_title", pal["primary"]), first=True)
        desc = t.get("desc") or t.get("subtitle") or ""
        if desc:
            cv.para(tf_main, parse_rich(desc, color=acc), size=9.2, color=pal.get("text_body", pal["text"]), before=2)

        # 右侧核心指标/占比（自适应单行排版，杜绝折行与孤字）
        if rate:
            rx = tx + w - metric_box_w - 0.20
            rtf = cv.tbox(rx, ty + (tier_h - 0.45) / 2.0, metric_box_w, 0.45, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(rtf, rate, size=metric_sz, bold=True, color=acc, align="right", first=True)

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
        n_it = len(items_list)
        it_sz = 11.0 if n_it <= 2 else 9.5
        it_gap = 10 if n_it <= 2 else 4
        it_line = 1.32 if n_it <= 2 else 1.20
        tf = cv.tbox(qx + 0.20, qy + 0.58, quad_w - 0.40, quad_h - 0.68)
        for j, it in enumerate(items_list[:4]):
            it_text = it if isinstance(it, str) else it.get("text", "")
            runs = [("▪ ", {"b": True, "c": col, "sz": it_sz})] + parse_rich(it_text, color=col)
            cv.para(tf, runs, size=it_sz, color=pal.get("text_body"),
                    first=(j == 0), before=it_gap if j > 0 else 0, line=it_line)

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
    n_p = len(pros_items)
    p_sz = 11.5 if n_p <= 2 else (10.5 if n_p == 3 else 9.8)
    p_gap = 14 if n_p <= 2 else (8 if n_p == 3 else 5)
    p_line = 1.35 if n_p <= 2 else 1.22
    tf_p = cv.tbox(MARGIN + 0.22, y + 0.65, half_w - 0.44, main_h - 0.85)
    for j, it in enumerate(pros_items[:5]):
        t = it if isinstance(it, str) else it.get("text", "")
        runs = [("✔  ", {"b": True, "c": acc_pros, "sz": p_sz})] + parse_rich(t, color=acc_pros)
        cv.para(tf_p, runs, size=p_sz, color=pal.get("text_body"), first=(j == 0), before=p_gap if j > 0 else 0, line=p_line)

    # 2. 右侧弊端/挑战卡片 (Cons)
    rx = MARGIN + half_w + 0.28
    cv.card(rx, y, half_w, main_h)
    cv.rect(rx, y, half_w, 0.05, fill=acc_cons)
    cv.badge(rx + 0.20, y + 0.20, "▲ 潜在风险与约束考量 (CONS)", dot=False,
             dot_color=acc_cons, text_color=acc_cons, size=10, h=0.32, pad_x=0.14)

    cons_items = cons_data if isinstance(cons_data, list) else (cons_data.get("items") or [cons_data])
    n_c = len(cons_items)
    c_sz = 11.5 if n_c <= 2 else (10.5 if n_c == 3 else 9.8)
    c_gap = 14 if n_c <= 2 else (8 if n_c == 3 else 5)
    c_line = 1.35 if n_c <= 2 else 1.22
    tf_c = cv.tbox(rx + 0.22, y + 0.65, half_w - 0.44, main_h - 0.85)
    for j, it in enumerate(cons_items[:5]):
        t = it if isinstance(it, str) else it.get("text", "")
        runs = [("▲  ", {"b": True, "c": acc_cons, "sz": c_sz})] + parse_rich(t, color=acc_cons)
        cv.para(tf_c, runs, size=c_sz, color=pal.get("text_body"), first=(j == 0), before=c_gap if j > 0 else 0, line=c_line)

    # 3. 底部权衡决策卡
    if has_takeaway:
        takeaway_card(cv, takeaway, y + main_h + 0.20, title="权衡决策与应对策略", h=0.95)

    return y + h


def faq_accordion(cv, y, faqs, h=4.8):
    """常见疑虑与问题解答页：3~4 个高质感手风琴式 Q&A 块。"""
    pal = cv.pal
    n = max(1, min(len(faqs), 4))
    gap = 0.22 if n <= 2 else 0.18
    item_h = (h - gap * (n - 1)) / float(n)
    acc = pal.get("accent")

    a_size = 11.5 if n <= 2 else (10.2 if n == 3 else 9.5)
    a_line = 1.35 if n <= 2 else 1.22

    for i, f in enumerate(faqs[:n]):
        iy = y + i * (item_h + gap)
        cv.card(MARGIN, iy, CONTENT_W, item_h)

        q_text = str(f.get("q") or f.get("question") or f.get("title") or "核心疑问")
        tag = f.get("tag") or f.get("category")

        cv.pill(MARGIN + 0.16, iy + 0.14, 0.45, 0.28, fill=pal.get("badge_bg"))
        qtf = cv.tbox(MARGIN + 0.16, iy + 0.14, 0.45, 0.28, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(qtf, "Q", size=11, bold=True, color=acc, align="center", first=True)

        q_sz = 12.0 if len(q_text) <= 30 else 10.5
        cv.text(MARGIN + 0.70, iy + 0.14, CONTENT_W - 2.2, 0.28, q_text,
                size=q_sz, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        if tag:
            cv.badge(MARGIN + CONTENT_W - 1.40, iy + 0.14, str(tag), dot=False, size=8.5, h=0.26)

        cv.divider(MARGIN + 0.16, iy + 0.48, CONTENT_W - 0.32)

        a_text = str(f.get("a") or f.get("answer") or f.get("desc") or "")
        anchor_mode = MSO_ANCHOR.MIDDLE if (n <= 2 and len(a_text) <= 160) else MSO_ANCHOR.TOP
        tf_a = cv.tbox(MARGIN + 0.20, iy + 0.54, CONTENT_W - 0.40, item_h - 0.60, anchor=anchor_mode)
        cv.para(tf_a, parse_rich(a_text, color=acc), size=a_size, color=pal.get("text_body"), line=a_line, first=True)

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

    # 核心背景与挑战
    challenge = case_data.get("challenge") or case_data.get("pain_points") or ""
    ch_title = str(case_data.get("challenge_title") or "【背景与核心挑战】")
    if challenge:
        cv.text(MARGIN + 0.20, y + 1.25, left_w - 0.40, 0.28, ch_title,
                size=10.5, bold=True, color=acc)
        tf_ch = cv.tbox(MARGIN + 0.20, y + 1.58, left_w - 0.40, 1.40)
        cv.para(tf_ch, parse_rich(challenge, color=acc), size=9.5, color=pal.get("text_body"), line=1.25, first=True)

    # 解决方案与关键举措
    solution = case_data.get("solution") or case_data.get("approach") or ""
    sol_title = str(case_data.get("solution_title") or "【关键方案与举措】")
    if solution:
        sol_y = y + 3.05
        cv.text(MARGIN + 0.20, sol_y, left_w - 0.40, 0.28, sol_title,
                size=10.5, bold=True, color=acc2)
        tf_sol = cv.tbox(MARGIN + 0.20, sol_y + 0.33, left_w - 0.40, h - (sol_y - y) - 0.45)
        cv.para(tf_sol, parse_rich(solution, color=acc2), size=9.5, color=pal.get("text_body"), line=1.25, first=True)

    # 右侧：量化成果数据与客户证言
    metrics = case_data.get("results") or case_data.get("metrics") or []
    top_metrics_h = 2.40
    cv.card(right_x, y, right_w, top_metrics_h)
    res_title = str(case_data.get("results_title") or "● 核心成果与成效量化")
    cv.badge(right_x + 0.20, y + 0.20, res_title, dot=False, size=10, h=0.30)

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
    """总结与落地行动清单页：50:50 对称双栏（左侧核心要点/结论 + 右侧行动计划/作业清单）。"""
    pal = cv.pal
    gap = 0.28
    # 严格 50:50 黄金等宽对称布局，彻底解决左窄右肥的失衡与视觉错位
    col_w = (CONTENT_W - gap) / 2.0
    left_w = col_w
    right_w = col_w
    right_x = MARGIN + left_w + gap
    acc = pal.get("accent")
    accents = pal.get("accents") or [acc]

    # 动态上下文感知的徽章标签：拒绝生硬的“战略核心”、“Next Steps”企业黑话
    left_title = (summary_data.get("takeaways_title") or
                  summary_data.get("left_title") or
                  summary_data.get("conclusions_title") or
                  "核心要点总结")
    right_title = (summary_data.get("actions_title") or
                   summary_data.get("right_title") or
                   summary_data.get("tasks_title") or
                   "落地安排与任务清单")

    # 左侧外卡片
    cv.card(MARGIN, y, left_w, h)
    cv.rect(MARGIN, y, left_w, 0.045, fill=acc)
    cv.badge(MARGIN + 0.22, y + 0.20, str(left_title), dot=True, dot_color=acc,
             bg=pal.get("badge_bg"), text_color=pal.get("text_title", pal["primary"]),
             size=10, h=0.30)

    takeaways = (summary_data.get("takeaways") or summary_data.get("conclusions") or
                 summary_data.get("highlights") or summary_data.get("points") or
                 summary_data.get("items") or [])
    if isinstance(takeaways, str):
        takeaways = [takeaways]

    # 右侧外卡片
    cv.card(right_x, y, right_w, h)
    cv.rect(right_x, y, right_w, 0.045, fill=acc)
    cv.badge(right_x + 0.22, y + 0.20, str(right_title), dot=True, dot_color=acc,
             bg=pal.get("badge_bg"), text_color=pal.get("text_title", pal["primary"]),
             size=10, h=0.30)

    actions = summary_data.get("actions") or summary_data.get("next_steps") or []
    act_count = max(1, min(len(actions), 4))
    act_gap = 0.16
    act_h = (h - 0.80 - act_gap * (act_count - 1)) / act_count

    # 左侧自适应卡片布局
    principle = (summary_data.get("principle") or summary_data.get("quote") or
                 summary_data.get("vision") or summary_data.get("core_message"))
    has_bottom_card = bool(principle)

    tk_count = max(1, min(len(takeaways), 4))
    if has_bottom_card:
        bottom_h = 1.10
        tk_available_h = h - 0.80 - bottom_h - 0.16
        tk_gap = 0.12
        tk_h = (tk_available_h - tk_gap * (tk_count - 1)) / tk_count
    else:
        tk_gap = act_gap
        tk_h = (h - 0.80 - tk_gap * (tk_count - 1)) / tk_count

    for j, tk in enumerate(takeaways[:tk_count]):
        ty = y + 0.62 + j * (tk_h + tk_gap)
        tk_acc = accents[j % len(accents)]
        cv.rrect(MARGIN + 0.18, ty, left_w - 0.36, tk_h,
                 fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.5, radius=0.06)

        num_str = "%02d" % (j + 1)
        b_end = cv.badge(MARGIN + 0.28, ty + (0.14 if tk_h > 1.0 else 0.10), num_str,
                         dot=False, bg=pal.get("badge_bg"), text_color=tk_acc,
                         size=9, h=0.26, pad_x=0.08)

        tk_raw = str(tk).strip()
        parts = [p.strip() for p in tk_raw.replace("：", ":").split(":", 1)] if (":" in tk_raw or "：" in tk_raw) else [tk_raw]
        text_x = b_end + 0.12
        text_w = (MARGIN + left_w - 0.36) - text_x

        if len(parts) == 2 and tk_h > 0.85:
            header, body = parts
            cv.text(text_x, ty + 0.10, text_w, 0.28,
                    header, size=11, bold=True, color=pal.get("text_title", pal["primary"]))
            tf_body = cv.tbox(text_x, ty + 0.38, text_w, tk_h - 0.44, anchor=MSO_ANCHOR.TOP if tk_h > 1.1 else MSO_ANCHOR.MIDDLE)
            cv.para(tf_body, parse_rich(body, color=tk_acc), size=9.6, color=pal.get("text_body"), line=1.22, first=True)
        else:
            tf_tk = cv.tbox(text_x, ty + 0.08, text_w, tk_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(tf_tk, parse_rich(tk_raw, color=tk_acc), size=10, color=pal.get("text_body"), line=1.22, first=True)

    if has_bottom_card:
        by = y + h - bottom_h - 0.18
        cv.rrect(MARGIN + 0.18, by, left_w - 0.36, bottom_h,
                 fill=pal.get("card_bg"), line=acc, lw=0.8, radius=0.06)
        cv.badge(MARGIN + 0.30, by + 0.14, "★ 重点提示 / 核心指引", dot=False,
                 bg=pal.get("badge_bg"), text_color=acc, size=8.5, h=0.24)
        tf_pr = cv.tbox(MARGIN + 0.30, by + 0.42, left_w - 0.60, bottom_h - 0.50, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_pr, parse_rich(str(principle), color=acc), size=10.5, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.25, first=True)

    for i, act in enumerate(actions[:act_count]):
        ay = y + 0.62 + i * (act_h + act_gap)
        cv.rrect(right_x + 0.18, ay, right_w - 0.36, act_h, fill=pal.get("card_subtle"), line=pal.get("card_border"))
        task_str = str(act.get("task") or act.get("title") or f"任务 {i+1}")
        owner_str = str(act.get("owner") or act.get("team") or "")
        due_str = str(act.get("due") or act.get("deadline") or "")
        has_meta = bool(owner_str or due_str)

        meta_w = 1.60 if has_meta else 0.0
        task_w = right_w - 0.54 - meta_w

        tf_act = cv.tbox(right_x + 0.28, ay + 0.08, task_w, act_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_act, parse_rich(task_str, color=acc), size=10.2, bold=True, color=pal.get("text_title", pal["primary"]), first=True)
        desc_str = act.get("desc") or ""
        if desc_str:
            cv.para(tf_act, parse_rich(desc_str, color=acc), size=8.5, color=pal.get("text_subtitle", pal["text_body"]), before=2)

        if has_meta:
            tf_meta = cv.tbox(right_x + right_w - meta_w - 0.26, ay + 0.08, meta_w, act_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
            if owner_str:
                cv.para(tf_meta, owner_str, size=9.2, bold=True, color=acc, align="right", first=True)
            if due_str:
                cv.para(tf_meta, due_str, size=8.2, color=pal.get("text_subtitle", pal["text_body"]), align="right", before=2 if owner_str else 0, first=not owner_str)

    return y + h


def matrix_quadrant(cv, y, matrix_data, h=4.8):
    """二维四象限矩阵分析页：战略研报/投资收益/价值难度矩阵（含XY轴标签与象限项目）。"""
    pal = cv.pal
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]
    acc = accents[0]
    acc2 = accents[1 % len(accents)]

    # 象限坐标轴定义
    x_axis = matrix_data.get("x_axis") or {"left": "低投入 / 低复杂度", "right": "高投入 / 高复杂度"}
    y_axis = matrix_data.get("y_axis") or {"top": "高战略收益 / 高商业价值", "bottom": "低战略收益 / 低商业价值"}

    axis_w = CONTENT_W
    box_gap = 0.22
    qw = (axis_w - box_gap) / 2.0
    qh = (h - 0.40 - box_gap) / 2.0
    start_y = y + 0.32

    # 绘制背景坐标轴提示
    cv.text(MARGIN, y, qw, 0.24, f"◀ {x_axis.get('left', '')}", size=9.2, color=pal.get("text_muted"), line=1.0)
    cv.text(MARGIN + qw + box_gap, y, qw, 0.24, f"{x_axis.get('right', '')} ▶", size=9.2, color=pal.get("text_muted"), align="right", line=1.0)

    # 四大象限定义：q2(左上), q1(右上), q3(左下), q4(右下)
    quadrants = [
        {"key": "q2", "default_name": "快速见效区 (Quick Wins)", "col": 0, "row": 0, "acc": acc, "tag": "优先落地"},
        {"key": "q1", "default_name": "核心主攻区 (Strategic Bets)", "col": 1, "row": 0, "acc": pal.get("primary", acc), "tag": "重点投入"},
        {"key": "q3", "default_name": "维持/外包区 (Low Priority)", "col": 0, "row": 1, "acc": pal.get("text_muted", pal.get("card_border", acc)), "tag": "严格控本"},
        {"key": "q4", "default_name": "长期探索区 (Long-term Vision)", "col": 1, "row": 1, "acc": acc2, "tag": "技术预研"}
    ]

    for qd in quadrants:
        qx = MARGIN + qd["col"] * (qw + box_gap)
        qy = start_y + qd["row"] * (qh + box_gap)
        q_info = matrix_data.get(qd["key"]) or {}
        if isinstance(q_info, list):
            q_info = {"items": q_info}

        q_name = str(q_info.get("name") or q_info.get("title") or qd["default_name"])
        q_tag = str(q_info.get("tag") or qd["tag"])
        items = q_info.get("items") or []

        # 象限卡片容器
        cv.card(qx, qy, qw, qh, radius=0.08)
        # 顶边微光线
        cv.pill(qx + 0.20, qy + 0.04, min(1.20, qw - 0.40), 0.035, fill=qd["acc"])

        # 象限标题与标签
        b_end = cv.badge(qx + 0.18, qy + 0.16, q_tag, dot=True, dot_color=qd["acc"], text_color=qd["acc"], size=9, h=0.28)
        cv.text(b_end + 0.14, qy + 0.16, qw - (b_end - qx) - 0.25, 0.28, q_name, size=11, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        cv.divider(qx + 0.18, qy + 0.52, qw - 0.36)

        # 象限项目要点
        tf = cv.tbox(qx + 0.20, qy + 0.60, qw - 0.40, qh - 0.70)
        for j, it in enumerate(items[:4]):
            it_str = it if isinstance(it, str) else str(it.get("title", it))
            runs = [("▪ ", {"b": True, "c": qd["acc"], "sz": 9.5})] + parse_rich(it_str, color=qd["acc"])
            cv.para(tf, runs, size=9.5, color=pal.get("text_body"), first=(j == 0), before=4 if j > 0 else 0)

    return y + h


def org_tree(cv, y, tree_data, h=4.8):
    """组织架构与战略决策树：顶层中枢 + 分支业务域 + 底层能力支撑。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    # 1. 顶层根节点 (Root Node，居中大卡片)
    root = tree_data.get("root") or {"title": "架构战略中枢 / 集团决策委员会", "role": "顶级决策", "desc": "统一战略规划与技术决策治理"}
    rw, rh = 6.2, 0.95
    rx = MARGIN + (CONTENT_W - rw) / 2.0
    cv.card(rx, y, rw, rh, radius=0.08)
    cv.rect(rx, y, rw, 0.04, fill=acc)

    root_title = str(root.get("title") or "核心决策中心")
    root_role = str(root.get("role") or root.get("tag") or "HQ")
    b_end = cv.badge(rx + 0.20, y + 0.16, root_role, dot=True, dot_color=acc, text_color=acc, size=9.5, h=0.28)
    cv.text(b_end + 0.16, y + 0.14, rw - (b_end - rx) - 0.25, 0.30, root_title, size=12.5, bold=True, color=pal.get("text_title", pal["primary"]))

    root_desc = str(root.get("desc") or "")
    if root_desc:
        cv.text(rx + 0.20, y + 0.52, rw - 0.40, 0.32, root_desc, size=9.5, color=pal.get("text_subtitle", pal["text_body"]))

    # 根节点向下引线
    stem_top = y + rh
    stem_bottom = stem_top + 0.45
    cv.rect(rx + rw / 2.0 - 0.015, stem_top, 0.03, stem_bottom - stem_top, fill=pal.get("divider", pal["card_border"]))

    # 2. 第二层分支 (Branches)
    branches = tree_data.get("branches") or tree_data.get("items") or []
    n_br = max(1, min(len(branches), 4))
    b_gap = 0.24
    bw = (CONTENT_W - b_gap * (n_br - 1)) / float(n_br)
    by = stem_bottom + 0.05
    bh = h - (by - y)

    # 横向分流总线
    if n_br > 1:
        first_cx = MARGIN + bw / 2.0
        last_cx = MARGIN + (n_br - 1) * (bw + b_gap) + bw / 2.0
        cv.rect(first_cx, stem_bottom, last_cx - first_cx, 0.025, fill=pal.get("divider", pal["card_border"]))

    for i, br in enumerate(branches[:n_br]):
        bx = MARGIN + i * (bw + b_gap)
        branch_cx = bx + bw / 2.0
        # 从横向总线向下引到分支卡片
        cv.rect(branch_cx - 0.015, stem_bottom, 0.03, by - stem_bottom, fill=pal.get("divider", pal["card_border"]))

        cv.card(bx, by, bw, bh, radius=0.08)
        bar_col = acc if i % 2 == 0 else acc2
        cv.pill(bx + 0.18, by + 0.04, min(1.10, bw - 0.36), 0.035, fill=bar_col)

        br_title = str(br.get("title") or f"业务域 {i+1}")
        br_tag = str(br.get("role") or br.get("tag") or f"0{i+1}")
        b_end = cv.badge(bx + 0.16, by + 0.16, br_tag, dot=False, size=9, h=0.28)
        cv.text(b_end + 0.14, by + 0.16, bw - (b_end - bx) - 0.16, 0.30, br_title, size=11.5, bold=True, color=pal.get("text_title", pal["primary"]))

        cv.divider(bx + 0.16, by + 0.52, bw - 0.32)

        # 分支下属子团队/子模块清单 (Leaves)
        sub_items = br.get("leaves") or br.get("teams") or br.get("items") or []
        sub_h = bh - 0.70
        tf_sub = cv.tbox(bx + 0.18, by + 0.60, bw - 0.36, sub_h)
        for j, leaf in enumerate(sub_items[:5]):
            leaf_str = leaf if isinstance(leaf, str) else str(leaf.get("name", leaf.get("title", "")))
            runs = [("✔ ", {"b": True, "c": bar_col, "sz": 9.2})] + parse_rich(leaf_str, color=bar_col)
            cv.para(tf_sub, runs, size=9.5, color=pal.get("text_body"), first=(j == 0), before=4 if j > 0 else 0)

    return y + h


def capability_radar(cv, y, radar_data, h=4.8):
    """多维能力雷达与竞品全景评估页：左右分栏（左侧维度条形打分 + 右侧综合得分卡）。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    left_w = 7.6
    right_w = CONTENT_W - left_w - 0.28
    right_x = MARGIN + left_w + 0.28

    # 左侧：核心能力维度双轨评估矩阵
    cv.card(MARGIN, y, left_w, h, radius=0.08)
    cv.badge(MARGIN + 0.22, y + 0.20, "● 核心能力维度全景对标", dot=False, size=10, h=0.30)

    dims = radar_data.get("dimensions") or radar_data.get("items") or []
    n_dims = max(1, min(len(dims), 6))
    dim_space = (h - 0.85) / float(n_dims)

    from .charts import progress_bar
    for i, dm in enumerate(dims[:n_dims]):
        dy = y + 0.65 + i * dim_space
        d_name = str(dm.get("name") or dm.get("label") or f"维度 {i+1}")
        d_score = float(dm.get("score") or 85)
        d_bm = float(dm.get("benchmark") or dm.get("rival") or 60)
        d_ratio = min(1.0, max(0.05, d_score / 100.0))

        cv.text(MARGIN + 0.22, dy, 2.20, 0.26, d_name, size=10.5, bold=True, color=pal.get("text_title", pal["primary"]), anchor=MSO_ANCHOR.MIDDLE)
        # 本方案进度条
        progress_bar(cv, MARGIN + 2.45, dy + 0.05, left_w - 2.70, 0.16, d_ratio, fill=acc, bg=pal.get("card_subtle", pal.get("bg_alt")), label=f"{int(d_score)}分")
        # 备注简述
        d_desc = dm.get("desc") or ""
        if d_desc and dim_space >= 0.70:
            cv.text(MARGIN + 2.45, dy + 0.28, left_w - 2.70, 0.22, str(d_desc), size=8.5, color=pal.get("text_muted"))

    # 右侧：综合能力大得分与卓越评级卡
    cv.card(right_x, y, right_w, h, radius=0.08)
    cv.pill(right_x + 0.22, y + 0.04, min(1.2, right_w - 0.44), 0.035, fill=acc2)
    cv.badge(right_x + 0.22, y + 0.20, "综合评估指数", dot=True, dot_color=acc2, text_color=acc2, size=9.5, h=0.30)

    total_score = str(radar_data.get("score") or "98.5")
    rating = str(radar_data.get("rating") or "EXCELLENT · 卓越领跑")

    cv.text(right_x + 0.22, y + 0.65, right_w - 0.44, 0.75,
            [(total_score, {"b": True, "c": acc, "sz": 42}),
             (" 分", {"b": True, "c": acc, "sz": 16})], line=1.0)
    cv.badge(right_x + 0.22, y + 1.48, rating, dot=False, bg=pal.get("badge_bg"), text_color=acc, size=9.5, h=0.28)

    cv.divider(right_x + 0.22, y + 1.95, right_w - 0.44)

    # 优势标签药丸
    highlights = radar_data.get("highlights") or ["高并发极致承载", "金融级秒级容灾", "零侵入平滑接入"]
    tf_hl = cv.tbox(right_x + 0.22, y + 2.10, right_w - 0.44, h - 2.25)
    for j, hl in enumerate(highlights[:4]):
        runs = [("★ ", {"b": True, "c": acc2, "sz": 9.5})] + parse_rich(str(hl), color=acc2)
        cv.para(tf_hl, runs, size=9.8, color=pal.get("text_body"), first=(j == 0), before=6 if j > 0 else 0)

    return y + h


def funnel_stages(cv, y, funnel_data, h=4.8):
    """转化漏斗专版：逐级下沉转化率阶梯卡片（适合营销增长/获客转化/商业漏斗）。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    if isinstance(funnel_data, list):
        stages = funnel_data
    elif isinstance(funnel_data, dict):
        stages = funnel_data.get("stages") or funnel_data.get("items") or []
    else:
        stages = []
    n = max(1, min(len(stages), 5))

    gap = 0.16
    tier_h = (h - gap * (n - 1)) / float(n)

    min_w = 7.5
    max_w = CONTENT_W

    for i, st in enumerate(stages[:n]):
        ratio = (i / float(n - 1)) if n > 1 else 0.5
        w = max_w - (max_w - min_w) * ratio
        tx = MARGIN + (CONTENT_W - w) / 2.0
        ty = y + i * (tier_h + gap)

        cv.rrect(tx, ty, w, tier_h, fill=pal.get("card_bg"), line=pal.get("card_border"), radius=0.08)
        cv.rect(tx, ty, 0.08, tier_h, fill=acc)

        # 阶段序号
        st_num = f"STAGE 0{i+1}" if "0" not in str(st.get("stage", "")) else str(st.get("stage"))
        b_end = cv.badge(tx + 0.20, ty + (tier_h - 0.30) / 2.0, st_num, dot=False, size=9.2, h=0.30)

        # 阶段标题与举措
        st_title = str(st.get("title") or st.get("name") or f"阶段 {i+1}")
        rate_str = str(st.get("rate") or st.get("conversion") or "")
        left_w = w - (b_end - tx) - (2.4 if rate_str else 0.4)

        tf = cv.tbox(b_end + 0.15, ty + 0.06, left_w, tier_h - 0.12, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf, st_title, size=11.5, bold=True, color=pal.get("text_title", pal["primary"]), first=True)
        st_desc = st.get("desc") or st.get("action") or ""
        if st_desc:
            cv.para(tf, parse_rich(str(st_desc), color=acc), size=9.0, color=pal.get("text_body"), before=2)

        # 右侧转化指标
        if rate_str:
            rtf = cv.tbox(tx + w - 2.20, ty + (tier_h - 0.40) / 2.0, 2.0, 0.40, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(rtf, rate_str, size=14, bold=True, color=acc, align="right", first=True)

    return y + h


def chat_dialogue(cv, y, chat_data, h=4.8):
    """AI Agent 对话流与用户访谈原声：模拟真实气泡流（左侧专家回答 + 右侧用户提问）。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    if isinstance(chat_data, list):
        dialogues = chat_data
    elif isinstance(chat_data, dict):
        dialogues = chat_data.get("dialogues") or chat_data.get("items") or []
    else:
        dialogues = []
    n = max(1, min(len(dialogues), 3))
    row_h = (h - 0.20 * (n - 1)) / float(n)

    for i, d in enumerate(dialogues[:n]):
        dy = y + i * (row_h + 0.20)
        is_user = (d.get("role") == "user" or d.get("sender") == "user")

        if is_user:
            # 用户提问气泡（右对齐）
            bw = min(8.5, CONTENT_W * 0.75)
            bx = MARGIN + CONTENT_W - bw
            cv.card(bx, dy, bw, row_h, fill=pal.get("card_subtle"), line=pal.get("card_border"), radius=0.08)
            cv.badge(bx + bw - 1.25, dy + 0.14, "用户咨询", dot=True, dot_color=acc, size=8.5, h=0.26)
            tf = cv.tbox(bx + 0.25, dy + 0.44, bw - 0.50, row_h - 0.55)
            cv.para(tf, str(d.get("text") or d.get("q") or ""), size=10.5, color=pal.get("text_title", pal["primary"]), first=True)
        else:
            # AI / 专家回复气泡（左对齐，带双层微卡高光）
            bw = min(10.2, CONTENT_W * 0.90)
            bx = MARGIN
            cv.card(bx, dy, bw, row_h, radius=0.08)
            cv.rect(bx, dy, 0.08, row_h, fill=acc)
            bot_tag = str(d.get("name") or d.get("role") or "AI 架构专家 Agent")
            cv.badge(bx + 0.20, dy + 0.14, bot_tag, dot=False, size=9, h=0.28)
            tf = cv.tbox(bx + 0.20, dy + 0.48, bw - 0.40, row_h - 0.58)
            cv.para(tf, parse_rich(str(d.get("text") or d.get("a") or ""), color=acc),
                    size=10, color=pal.get("text_body"), line=1.35, first=True)

    return y + h


def split_showcase(cv, y, showcase_data, h=4.8):
    """左右分屏产品/技术展示页：左侧高科技编辑器视窗 + 右侧 3 个 Bento 卖点砖块。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    left_w = 6.2
    right_w = CONTENT_W - left_w - 0.28
    right_x = MARGIN + left_w + 0.28

    # 左侧：macOS 风格代码/终端产品视窗卡片
    cv.card(MARGIN, y, left_w, h, radius=0.08)
    # macOS 三色控制小圆点
    cv.circle(MARGIN + 0.25, y + 0.22, 0.08, fill="#FF5F56")
    cv.circle(MARGIN + 0.45, y + 0.22, 0.08, fill="#FFBD2E")
    cv.circle(MARGIN + 0.65, y + 0.22, 0.08, fill="#27C93F")

    window_title = str(showcase_data.get("window_title") or "Terminal · Architecture Core SDK")
    cv.text(MARGIN + 0.95, y + 0.14, left_w - 1.20, 0.26, window_title, size=9.2, color=pal.get("text_muted"))
    cv.divider(MARGIN, y + 0.45, left_w)

    code_lines = showcase_data.get("code") or showcase_data.get("console") or [
        "// 全局唯一 ID 批量自增发号服务",
        "const leaf = new SegmentBufferEngine({",
        "    step: 200000,         // 双 Buffer 步长",
        "    clockFallback: true,  // NTP 时钟防回拨",
        "    multiRegion: 'active' // 全球多活",
        "});",
        "const orderId = await leaf.nextId();",
        "console.log('QPS: 1,200,000 | Latency: 0.12ms');"
    ]
    if isinstance(code_lines, str):
        code_lines = code_lines.split("\n")

    tf_code = cv.tbox(MARGIN + 0.25, y + 0.60, left_w - 0.50, h - 0.80)
    for j, cl in enumerate(code_lines[:12]):
        cv.para(tf_code, cl, size=9.0, color=pal.get("text_subtitle", pal["text_body"]), line=1.35, first=(j == 0))

    # 右侧：3 个垂直排布的 Bento 卖点砖块
    features = showcase_data.get("features") or showcase_data.get("items") or []
    n_feat = max(1, min(len(features), 3))
    f_gap = 0.18
    f_h = (h - f_gap * (n_feat - 1)) / float(n_feat)

    for i, f in enumerate(features[:n_feat]):
        fy = y + i * (f_h + f_gap)
        cv.card(right_x, fy, right_w, f_h, radius=0.08)
        f_acc = acc if i % 2 == 0 else acc2
        cv.rect(right_x, fy, 0.08, f_h, fill=f_acc)

        f_title = str(f.get("title") or f"特性 0{i+1}")
        f_tag = str(f.get("tag") or f"0{i+1}")
        b_end = cv.badge(right_x + 0.20, fy + 0.16, f_tag, dot=False, size=9, h=0.28)
        cv.text(b_end + 0.14, fy + 0.16, right_w - (b_end - right_x) - 0.25, 0.30, f_title, size=11.5, bold=True, color=pal.get("text_title", pal["primary"]))

        f_desc = str(f.get("desc") or "")
        if f_desc:
            tf_f = cv.tbox(right_x + 0.20, fy + 0.52, right_w - 0.40, f_h - 0.60)
            cv.para(tf_f, parse_rich(f_desc, color=f_acc), size=9.5, color=pal.get("text_body"), line=1.28, first=True)


    return y + h


def metric_grid(cv, y, metrics_data, h=4.8):
    """高密度微指标监控大盘：2x3 或 2x4 紧凑指标阵列（杜绝空洞）。"""
    pal = cv.pal
    accents = pal.get("accents") or [pal["accent"], pal.get("accent2", pal["primary"])]
    if isinstance(metrics_data, list):
        cards = metrics_data
    elif isinstance(metrics_data, dict):
        cards = metrics_data.get("metrics") or metrics_data.get("kpis") or metrics_data.get("items") or []
    else:
        cards = []
    n = max(1, min(len(cards), 8))
    cols = 4 if n >= 7 else (3 if n >= 5 else min(n, 4))
    rows = (n + cols - 1) // cols

    gap_x = 0.20
    gap_y = 0.20
    cw = (CONTENT_W - gap_x * (cols - 1)) / float(cols)
    ch = (h - gap_y * (rows - 1)) / float(rows)

    for i, c in enumerate(cards[:cols * rows]):
        ri = i // cols
        ci = i % cols
        cx = MARGIN + ci * (cw + gap_x)
        cy = y + ri * (ch + gap_y)
        card_acc = accents[i % len(accents)]

        cv.card(cx, cy, cw, ch, radius=0.08)
        cv.pill(cx + 0.18, cy + 0.04, min(1.0, cw - 0.36), 0.03, fill=card_acc)

        label = str(c.get("label") or f"指标 0{i+1}")
        val_str = str(c.get("value") or "")
        unit_str = str(c.get("unit") or "")
        trend_str = str(c.get("trend") or "")
        desc = str(c.get("desc") or c.get("caption") or "")

        # 顶部标签徽章（独占顶部，保证 6~10 个汉字均不被右侧挤压）
        cv.badge(cx + 0.16, cy + 0.14, label, dot=True, dot_color=card_acc, text_color=card_acc, size=8.5, h=0.25)

        # 中部大字号指标数值
        num_sz = 24 if ch <= 2.1 else 30
        num_runs = [(val_str, {"b": True, "c": card_acc, "sz": num_sz})]
        if unit_str:
            num_runs.append((" " + unit_str, {"b": True, "c": card_acc, "sz": 12.5}))
        cv.text(cx + 0.16, cy + 0.42, cw - 0.32, 0.44, num_runs, line=1.0)

        # 底部区域：细微分割线 + 趋势胶囊 + 辅助说明
        if trend_str or desc:
            div_y = cy + ch - 0.46
            cv.divider(cx + 0.16, div_y, cw - 0.32)
            cur_bx = cx + 0.16
            item_y = div_y + 0.08
            if trend_str:
                cur_bx = cv.badge(cur_bx, item_y, trend_str, dot=False,
                                  bg=pal.get("badge_bg"), text_color=card_acc,
                                  size=8.2, h=0.24, pad_x=0.10) + 0.08
            if desc:
                avail_desc_w = max(0.4, cx + cw - 0.16 - cur_bx)
                cv.text(cur_bx, item_y, avail_desc_w, 0.24, desc,
                        size=8.5, color=pal.get("text_muted"), line=1.0,
                        anchor=MSO_ANCHOR.MIDDLE)

    return y + h


def quote_focus(cv, y, quote_data, h=4.8):
    """沉浸式高管观点与大引用页：超大双引号饰纹 + 人物独白 + 战略承诺胶囊。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])

    cv.card(MARGIN, y, CONTENT_W, h, radius=0.10)
    cv.rect(MARGIN, y, CONTENT_W, 0.045, fill=acc)

    # 浅色大双引号水印
    cv.text(MARGIN + 0.50, y + 0.15, 2.0, 1.2, "“", size=72, bold=True, color=pal.get("card_border"))

    # 核心引言金句 (根据文字长度自适应字号 16~24pt，杜绝撞线与溢出)
    quote_text = str(quote_data.get("quote") or quote_data.get("statement") or "技术不是炫技，而是业务价值的最坚固护城河。")
    q_len = sum(1.0 if ord(c) > 127 else 0.52 for c in quote_text)
    if q_len <= 35:
        q_sz = 24
        q_line = 1.40
    elif q_len <= 70:
        q_sz = 20
        q_line = 1.32
    elif q_len <= 110:
        q_sz = 17
        q_line = 1.25
    else:
        q_sz = 14.5
        q_line = 1.18

    tf_q = cv.tbox(MARGIN + 0.80, y + 0.55, CONTENT_W - 1.60, 1.65, anchor=MSO_ANCHOR.MIDDLE)
    cv.para(tf_q, parse_rich(quote_text, color=acc), size=q_sz, bold=True,
            color=pal.get("text_title", pal["primary"]), line=q_line, first=True)

    cv.divider(MARGIN + 0.80, y + 2.40, CONTENT_W - 1.60)

    # 发言人信息（头像 + 姓名 + 职务）
    author = str(quote_data.get("author") or "核心架构师")
    title = str(quote_data.get("title") or "基础技术委员会首席架构师")
    company = str(quote_data.get("company") or "")

    cv.circle(MARGIN + 0.80, y + 2.60, 0.65, fill=pal.get("badge_bg"), line=acc)
    cv.text(MARGIN + 0.80, y + 2.70, 0.65, 0.45, author[:2], size=13, bold=True, color=acc, align="center")

    cv.text(MARGIN + 1.65, y + 2.60, 6.0, 0.32, author, size=14, bold=True, color=pal.get("text_title", pal["primary"]))
    cv.text(MARGIN + 1.65, y + 2.93, 6.0, 0.28, f"{title}  {company}".strip(), size=10, color=pal.get("text_muted"))

    # 底部战略承诺胶囊 (自适应 3~4 个)
    commitments = quote_data.get("commitments") or quote_data.get("tags") or ["100% 故障自愈", "毫秒级发布容灾", "全球多活数据零丢失"]
    cx_cur = MARGIN + 0.80
    c_y = y + h - 0.70
    for com in commitments[:4]:
        c_str = str(com)
        cw = len(c_str) * 0.12 + 0.45
        if cx_cur + cw <= MARGIN + CONTENT_W - 0.80:
            cv.pill(cx_cur, c_y, cw, 0.32, fill=pal.get("card_subtle"), line=pal.get("card_border"))
            tf_c = cv.tbox(cx_cur, c_y, cw, 0.32, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(tf_c, c_str, size=9.2, bold=True, color=acc, align="center", first=True)
            cx_cur += cw + 0.20

    return y + h


def chart_analytics(cv, y, analytics_data, h=4.8):
    """数据洞察与原生图表分析页：左侧高保真大图表面板（折线/面积/环形/柱状） + 右侧核心洞察与指标。"""
    from .charts import render_chart
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    left_w = 7.5
    right_w = CONTENT_W - left_w - 0.28
    right_x = MARGIN + left_w + 0.28

    chart_info = analytics_data.get("chart") or analytics_data
    chart_title = str(chart_info.get("title") or "核心数据趋势分析")
    chart_tag = str(chart_info.get("tag") or chart_info.get("type", "DATA").upper())

    # 1. 左侧大图表容器卡片
    cv.card(MARGIN, y, left_w, h, radius=0.08)
    cv.pill(MARGIN + 0.20, y + 0.04, min(1.2, left_w - 0.40), 0.035, fill=acc)

    b_end = cv.badge(MARGIN + 0.20, y + 0.16, chart_tag, dot=True, dot_color=acc, text_color=acc, size=9.0, h=0.28)
    cv.text(b_end + 0.14, y + 0.16, left_w - (b_end - MARGIN) - 0.30, 0.30, chart_title, size=12.5, bold=True, color=pal.get("text_title", pal["primary"]))
    cv.divider(MARGIN + 0.20, y + 0.52, left_w - 0.40)

    # 绘制原生图表（折线/面积/饼图/柱状图）
    chart_y = y + 0.62
    chart_h = h - 0.85
    render_chart(cv, MARGIN + 0.20, chart_y, left_w - 0.40, chart_h, chart_info)

    # 2. 右侧：洞察与结论面板
    cv.card(right_x, y, right_w, h, radius=0.08)
    cv.pill(right_x + 0.20, y + 0.04, min(1.2, right_w - 0.40), 0.035, fill=acc2)

    kpi = analytics_data.get("kpi") or {}
    kpi_val = str(kpi.get("value") or "")
    kpi_unit = str(kpi.get("unit") or "")
    kpi_label = str(kpi.get("label") or "关键统计指标")
    kpi_trend = str(kpi.get("trend") or "")

    cv.badge(right_x + 0.20, y + 0.16, kpi_label, dot=False, size=9.0, h=0.28)
    if kpi_trend:
        cv.badge(right_x + right_w - 1.25, y + 0.16, kpi_trend, dot=False, bg=pal.get("badge_bg"), text_color=acc2, size=8.5, h=0.28)

    # 核心大数字
    if kpi_val:
        num_runs = [(kpi_val, {"b": True, "c": acc, "sz": 32})]
        if kpi_unit:
            num_runs.append((" " + kpi_unit, {"b": True, "c": acc, "sz": 14}))
        cv.text(right_x + 0.20, y + 0.52, right_w - 0.40, 0.58, num_runs, line=1.0)
        cv.divider(right_x + 0.20, y + 1.18, right_w - 0.40)
        insights_y = y + 1.28
    else:
        insights_y = y + 0.60

    # 洞察要点
    insights = analytics_data.get("insights") or analytics_data.get("items") or analytics_data.get("takeaways") or []
    remain_h = (y + h) - insights_y - 0.75
    tf = cv.tbox(right_x + 0.20, insights_y, right_w - 0.40, remain_h)
    for j, it in enumerate(insights[:5]):
        it_str = it if isinstance(it, str) else str(it.get("text", it.get("title", "")))
        runs = [("▪ ", {"b": True, "c": acc2, "sz": 9.5})] + parse_rich(it_str, color=acc2)
        cv.para(tf, runs, size=9.5, color=pal.get("text_body"), first=(j == 0), before=6 if j > 0 else 0)

    # 底部决策/启示胶囊卡片
    decision = analytics_data.get("decision") or analytics_data.get("note") or analytics_data.get("conclusion")
    if decision:
        dec_y = y + h - 0.60
        cv.pill(right_x + 0.20, dec_y, right_w - 0.40, 0.42, fill=pal.get("card_subtle"), line=pal.get("card_border"))
        tf_dec = cv.tbox(right_x + 0.25, dec_y, right_w - 0.50, 0.42, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_dec, [("💡 核心启示：", {"b": True, "c": acc}), (str(decision), {"c": pal.get("text_title", pal["primary"])})], size=9.0, first=True)

    return y + h


def code_walkthrough(cv, y, code_data, h=4.8):
    """程序员源码深度剖析与算法走读页：左侧高亮编辑器视窗 + 右侧算法复杂度、时序调用与踩坑 Tips。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    left_w = 6.6
    right_w = CONTENT_W - left_w - 0.28
    right_x = MARGIN + left_w + 0.28

    # 1. 左侧代码视窗
    cv.card(MARGIN, y, left_w, h, radius=0.08)
    cv.circle(MARGIN + 0.25, y + 0.22, 0.08, fill="#FF5F56")
    cv.circle(MARGIN + 0.45, y + 0.22, 0.08, fill="#FFBD2E")
    cv.circle(MARGIN + 0.65, y + 0.22, 0.08, fill="#27C93F")

    filename = str(code_data.get("file") or code_data.get("filename") or "SegmentBuffer.java · Core Engine")
    cv.text(MARGIN + 0.95, y + 0.14, left_w - 1.20, 0.26, filename, size=9.2, color=pal.get("text_muted"))
    cv.divider(MARGIN, y + 0.45, left_w)

    raw_code = code_data.get("code") or []
    if isinstance(raw_code, str):
        raw_code = raw_code.split("\n")

    tf_code = cv.tbox(MARGIN + 0.25, y + 0.55, left_w - 0.50, h - 0.70)
    for j, line in enumerate(raw_code[:14]):
        line_str = str(line)
        line_num = f"{j+1:02d}  "
        is_comment = line_str.strip().startswith("//") or line_str.strip().startswith("#")
        c_color = pal.get("text_muted") if is_comment else pal.get("text_title", pal["primary"])
        runs = [(line_num, {"b": False, "c": pal.get("card_border", pal.get("text_muted")), "sz": 8.5}),
                (line_str, {"b": not is_comment and ("class" in line_str or "def" in line_str or "return" in line_str), "c": c_color, "sz": 9.0})]
        cv.para(tf_code, runs, line=1.30, first=(j == 0))

    # 2. 右侧栏（拆为上、中、下三个精致功能卡）
    right_h_top = 1.35
    right_h_mid = 1.45
    right_h_bot = h - right_h_top - right_h_mid - 0.36

    # 2.1 上卡：调用时序与入参输出
    top_y = y
    cv.card(right_x, top_y, right_w, right_h_top, radius=0.08)
    cv.pill(right_x + 0.18, top_y + 0.04, min(1.0, right_w - 0.36), 0.03, fill=acc)
    b_end = cv.badge(right_x + 0.16, top_y + 0.14, "时序调用与链路", dot=True, dot_color=acc, size=8.8, h=0.26)
    call_flow = code_data.get("call_flow") or code_data.get("flow") or ["业务网关调用", "本地内存无锁 Atomic CAS", "成功写入环形 Buffer"]
    tf_flow = cv.tbox(right_x + 0.18, top_y + 0.44, right_w - 0.36, right_h_top - 0.50)
    for j, cf in enumerate(call_flow[:3]):
        cv.para(tf_flow, f"step {j+1}: {cf}", size=9.0, color=pal.get("text_body"), line=1.22, first=(j == 0))

    # 2.2 中卡：算法复杂度与资源消耗
    mid_y = top_y + right_h_top + 0.18
    cv.card(right_x, mid_y, right_w, right_h_mid, radius=0.08)
    cv.pill(right_x + 0.18, mid_y + 0.04, min(1.0, right_w - 0.36), 0.03, fill=acc2)
    cv.badge(right_x + 0.16, mid_y + 0.14, "复杂度与并发性能", dot=True, dot_color=acc2, size=8.8, h=0.26)
    metrics = code_data.get("performance") or code_data.get("complexity") or [
        {"k": "时间复杂度", "v": "O(1) 纳秒级"},
        {"k": "空间开销", "v": "常数级 2MB"},
        {"k": "锁竞争度", "v": "Zero-Lock 零阻塞"}
    ]
    tf_perf = cv.tbox(right_x + 0.18, mid_y + 0.44, right_w - 0.36, right_h_mid - 0.50)
    for j, pm in enumerate(metrics[:3]):
        if isinstance(pm, dict):
            runs_pm = [(pm.get("k", "") + ": ", {"b": True, "c": acc2}), (str(pm.get("v", "")), {"c": pal.get("text_title", pal["primary"])})]
        else:
            runs_pm = [("✔ ", {"b": True, "c": acc2}), (str(pm), {"c": pal.get("text_body")})]
        cv.para(tf_perf, runs_pm, size=9.0, line=1.22, first=(j == 0))

    # 2.3 下卡：⚠️ 线上避坑指南（Pitfalls）
    bot_y = mid_y + right_h_mid + 0.18
    cv.card(right_x, bot_y, right_w, right_h_bot, fill=pal.get("card_subtle"), line=pal.get("card_border"), radius=0.08)
    warn_col = pal.get("accent2", pal.get("accent"))
    cv.badge(right_x + 0.16, bot_y + 0.12, "⚠️ 线上避坑指南", dot=False, bg=pal.get("badge_bg"), text_color=warn_col, size=8.8, h=0.26)
    pitfalls = code_data.get("pitfalls") or code_data.get("tips") or code_data.get("warnings") or ["严禁在循环内同步调用远程 RPC", "时钟回拨超过 50ms 必须立即启动自旋熔断"]
    tf_pit = cv.tbox(right_x + 0.18, bot_y + 0.42, right_w - 0.36, right_h_bot - 0.48)
    for j, pf in enumerate(pitfalls[:3]):
        cv.para(tf_pit, f"• {pf}", size=8.8, color=pal.get("text_body"), line=1.20, first=(j == 0))

    return y + h


def classroom_quiz(cv, y, quiz_data, h=4.8):
    """大中小学课堂教学互动测验与例题解析页：题目卡片 + 4 个选项按钮卡片 + 名师权威思路点拨。"""
    pal = cv.pal
    acc = pal.get("accent", pal["primary"])
    acc2 = pal.get("accent2", pal["primary"])

    # 1. 顶部题目大卡片
    question = str(quiz_data.get("question") or quiz_data.get("title") or "请指出下列关于微积分极限定理表述正确的是？")
    q_type = str(quiz_data.get("type") or quiz_data.get("tag") or "随堂测验 · 经典考题")
    q_h = 1.25
    cv.card(MARGIN, y, CONTENT_W, q_h, radius=0.08)
    cv.pill(MARGIN + 0.20, y + 0.04, 1.2, 0.035, fill=acc)
    b_end = cv.badge(MARGIN + 0.20, y + 0.16, q_type, dot=True, dot_color=acc, text_color=acc, size=9.2, h=0.28)

    q_len = sum(1.0 if ord(c) > 127 else 0.52 for c in question)
    q_sz = 13.5 if q_len <= 45 else (11.5 if q_len <= 80 else 10.0)
    tf_q = cv.tbox(MARGIN + 0.20, y + 0.48, CONTENT_W - 0.40, q_h - 0.55)
    cv.para(tf_q, question, size=q_sz, bold=True, color=pal.get("text_title", pal["primary"]), line=1.20, first=True)

    # 2. 中部选项卡片（2x2 排布）
    options = quiz_data.get("options") or quiz_data.get("choices") or [
        {"tag": "A", "text": "连续函数在闭区间上必存在最大值与最小值"},
        {"tag": "B", "text": "可导函数在极值点处的导数值必然恒等于零"},
        {"tag": "C", "text": "若数列极限存在，则该数列必为单调有界数列"},
        {"tag": "D", "text": "无穷小量的倒数必然是无穷大量"}
    ]
    answer = str(quiz_data.get("answer") or quiz_data.get("correct") or "A")

    opt_y = y + q_h + 0.18
    opt_gap_x = 0.24
    opt_gap_y = 0.16
    opt_w = (CONTENT_W - opt_gap_x) / 2.0
    opt_h = 0.85

    for i, opt in enumerate(options[:4]):
        row_i = i // 2
        col_i = i % 2
        ox = MARGIN + col_i * (opt_w + opt_gap_x)
        oy = opt_y + row_i * (opt_h + opt_gap_y)

        opt_tag = str(opt.get("tag") or chr(65 + i)) if isinstance(opt, dict) else chr(65 + i)
        opt_text = str(opt.get("text", opt) if isinstance(opt, dict) else opt)
        is_answer = (opt_tag.strip().upper() == answer.strip().upper())

        card_line = acc if is_answer else pal.get("card_border")
        card_fill = pal.get("card_bg") if is_answer else pal.get("card_subtle")
        cv.card(ox, oy, opt_w, opt_h, fill=card_fill, line=card_line, lw=1.2 if is_answer else 0.6, radius=0.08)

        b_col = acc if is_answer else pal.get("text_muted")
        cv.circle(ox + 0.32, oy + (opt_h - 0.36) / 2.0, 0.36, fill=pal.get("badge_bg"), line=b_col)
        cv.text(ox + 0.32, oy + (opt_h - 0.36) / 2.0 + 0.05, 0.36, 0.26, opt_tag, size=11, bold=True, color=b_col, align="center")

        opt_text_len = sum(1.0 if ord(c) > 127 else 0.52 for c in opt_text)
        opt_sz = 10.0 if opt_text_len <= 35 else 8.8
        tf_opt = cv.tbox(ox + 0.80, oy + 0.08, opt_w - 0.95, opt_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
        cv.para(tf_opt, opt_text, size=opt_sz, bold=is_answer, color=pal.get("text_title", pal["primary"]) if is_answer else pal.get("text_body"), line=1.18, first=True)

    # 3. 底部名师点拨与思路解析卡片
    analysis_y = opt_y + 2 * opt_h + opt_gap_y + 0.18
    analysis_h = (y + h) - analysis_y
    cv.card(MARGIN, analysis_y, CONTENT_W, analysis_h, fill=pal.get("card_subtle"), line=pal.get("card_border"), radius=0.08)

    ans_pill = f"✔ 正确答案：{answer}" if answer else "💡 名师解题思路解析"
    b_ans = cv.badge(MARGIN + 0.20, analysis_y + 0.14, ans_pill, dot=False, bg=pal.get("badge_bg"), text_color=acc, size=9.5, h=0.28)

    analysis_text = str(quiz_data.get("analysis") or quiz_data.get("explanation") or "由魏尔斯特拉斯极值定理可知：在闭区间 [a, b] 上的连续函数必能取得其最大值和最小值。")
    an_len = sum(1.0 if ord(c) > 127 else 0.52 for c in analysis_text)
    an_sz = 9.8 if an_len <= 70 else (9.0 if an_len <= 140 else 8.2)
    tf_an = cv.tbox(b_ans + 0.20, analysis_y + 0.10, CONTENT_W - (b_ans - MARGIN) - 0.40, analysis_h - 0.20, anchor=MSO_ANCHOR.MIDDLE)
    cv.para(tf_an, [("核心解析：", {"b": True, "c": acc}), (analysis_text, {"c": pal.get("text_body")})], size=an_sz, line=1.22, first=True)

    return y + h


def agenda_grid(cv, y, chapters, h=4.8):
    """全局大目录导航大盘页（支持 3~6 章节）：横向立柱或 2x3 Bento 网格，自适应层次与子议题。"""
    pal = cv.pal
    n = len(chapters)
    if n == 0:
        return y

    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

    if n <= 4:
        # 3 或 4 个章节：全屏横向立柱排布
        cols = n
        gap = 0.24
        cw = (CONTENT_W - gap * (cols - 1)) / float(cols)
        card_h = h

        for i, ch in enumerate(chapters):
            cx = MARGIN + i * (cw + gap)
            acc_i = accents[i % len(accents)]
            cv.card(cx, y, cw, card_h, radius=0.08)

            # 顶部发光微条
            cv.pill(cx + 0.20, y + 0.05, min(1.30, cw - 0.40), 0.035, fill=acc_i)

            ch_dict = ch if isinstance(ch, dict) else {"title": str(ch)}
            num_str = str(ch_dict.get("number") or ch_dict.get("chapter") or ("%02d" % (i + 1)))
            if num_str.isdigit():
                num_str = "%02d" % int(num_str)

            b_end = cv.badge(cx + 0.20, y + 0.20, num_str, dot=False,
                             bg=pal.get("badge_bg"), text_color=acc_i, size=9.5, h=0.28, pad_x=0.10)

            tag_str = ch_dict.get("tag") or ch_dict.get("subtitle") or ch_dict.get("en")
            if tag_str:
                cv.text(b_end + 0.12, y + 0.18, max(0.5, cw - (b_end - cx) - 0.32), 0.28,
                        str(tag_str).upper(), size=8.5, color=pal.get("text_muted"), anchor=MSO_ANCHOR.MIDDLE)

            title_str = str(ch_dict.get("title") or f"章节 {i+1}")
            cv.text(cx + 0.20, y + 0.60, cw - 0.40, 0.45, title_str,
                    size=13.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.12)

            cv.divider(cx + 0.20, y + 1.15, cw - 0.40)

            items = ch_dict.get("items") or ch_dict.get("topics") or ch_dict.get("bullets")
            desc = ch_dict.get("desc") or ch_dict.get("summary")
            content_y = y + 1.30
            content_h = card_h - 1.50

            if items:
                bullet_list(cv, cx + 0.20, content_y, cw - 0.40, content_h,
                            items, size=10, gap=6, line=1.24, dot_color=acc_i)
            elif desc:
                tf = cv.tbox(cx + 0.20, content_y, cw - 0.40, content_h)
                cv.para(tf, parse_rich(str(desc), color=acc_i), size=10.5,
                        color=pal.get("text_body"), line=1.32, first=True)
    else:
        # 5~6 个章节：2 行 x 3 列 Bento 网格排布
        cols = 3
        rows = (n + cols - 1) // cols
        gap_x = 0.24
        gap_y = 0.20
        cw = (CONTENT_W - gap_x * (cols - 1)) / float(cols)
        card_h = (h - gap_y * (rows - 1)) / float(rows)

        for i, ch in enumerate(chapters[:6]):
            ri = i // cols
            ci = i % cols
            cx = MARGIN + ci * (cw + gap_x)
            cy = y + ri * (card_h + gap_y)
            acc_i = accents[i % len(accents)]

            cv.card(cx, cy, cw, card_h, radius=0.08)
            cv.pill(cx + 0.18, cy + 0.04, min(1.20, cw - 0.36), 0.03, fill=acc_i)

            ch_dict = ch if isinstance(ch, dict) else {"title": str(ch)}
            num_str = str(ch_dict.get("number") or ch_dict.get("chapter") or ("%02d" % (i + 1)))
            if num_str.isdigit():
                num_str = "%02d" % int(num_str)

            b_end = cv.badge(cx + 0.18, cy + 0.15, num_str, dot=False,
                             bg=pal.get("badge_bg"), text_color=acc_i, size=9.0, h=0.26, pad_x=0.08)

            title_str = str(ch_dict.get("title") or f"章节 {i+1}")
            cv.text(b_end + 0.14, cy + 0.13, cw - (b_end - cx) - 0.30, 0.32, title_str,
                    size=12, bold=True, color=pal.get("text_title", pal["primary"]), line=1.10)

            cv.divider(cx + 0.18, cy + 0.52, cw - 0.36)

            items = ch_dict.get("items") or ch_dict.get("topics") or ch_dict.get("bullets")
            desc = ch_dict.get("desc") or ch_dict.get("summary")
            content_y = cy + 0.64
            content_h = card_h - 0.76

            if items:
                bullet_list(cv, cx + 0.18, content_y, cw - 0.36, content_h,
                            items[:3], size=9.8, gap=5, line=1.22, dot_color=acc_i)
            elif desc:
                tf = cv.tbox(cx + 0.18, content_y, cw - 0.36, content_h)
                cv.para(tf, parse_rich(str(desc), color=acc_i), size=10.5,
                        color=pal.get("text_body"), line=1.32, first=True)

    return y + h


def section_stepper(cv, y, chapters, current_idx=0, h=1.65):
    """章节进度轴导航条：清晰高亮当前章节，静默已过/未来章节。"""
    pal = cv.pal
    n = len(chapters)
    if n == 0:
        return y
    accents = pal.get("accents") or [pal.get("accent", pal["primary"])]
    curr_acc = accents[current_idx % len(accents)]

    gap = 0.18
    n_display = min(n, 6)
    cw = (CONTENT_W - gap * (n_display - 1)) / float(n_display)

    for i, ch in enumerate(chapters[:n_display]):
        cx = MARGIN + i * (cw + gap)
        ch_dict = ch if isinstance(ch, dict) else {"title": str(ch)}
        is_curr = (i == current_idx)
        is_past = (i < current_idx)

        c_fill = pal.get("card_bg") if is_curr else pal.get("card_subtle")
        c_line = curr_acc if is_curr else pal.get("card_border")
        cv.card(cx, y, cw, h, fill=c_fill, line=c_line, lw=1.2 if is_curr else 0.5, radius=0.08)

        num_str = "%02d" % (i + 1)
        if is_curr:
            cv.pill(cx + 0.18, y + 0.04, min(1.2, cw - 0.36), 0.035, fill=curr_acc)
            b_text = f"★ {num_str}"
            b_col = curr_acc
        elif is_past:
            b_text = f"✔ {num_str}"
            b_col = pal.get("text_muted")
        else:
            b_text = num_str
            b_col = pal.get("text_muted")

        cv.badge(cx + 0.16, y + 0.16, b_text, dot=False,
                 bg=pal.get("badge_bg"), text_color=b_col, size=8.5, h=0.25, pad_x=0.08)

        t_title = str(ch_dict.get("title") or f"章节 {i+1}")
        t_col = pal.get("text_title", pal["primary"]) if is_curr else pal.get("text_muted")
        cv.text(cx + 0.16, y + 0.52, cw - 0.32, 0.40, t_title,
                size=11 if is_curr else 10, bold=is_curr, color=t_col, line=1.12)

        if is_curr:
            sub = ch_dict.get("subtitle") or ch_dict.get("tag") or "当前进行章节"
            cv.text(cx + 0.16, y + 0.98, cw - 0.32, 0.35, str(sub),
                    size=9, color=curr_acc, line=1.1)

    return y + h




