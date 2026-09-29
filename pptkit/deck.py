# -*- coding: utf-8 -*-
"""现代化 Deck 构建器：将声明式 JSON spec 渲染为顶级商业与科技汇报演示文稿。

支持现代演示文稿视觉层级、Bento-Grid 栅格系统、自适应留白填充与单页智能体检。
"""
from __future__ import annotations

import os
import re

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR

from . import layouts as L
from . import theme as T
from .primitives import Canvas

_TAG_RE = re.compile(r"\[\[(.*?)\]\]|\*\*(.*?)\*\*")


def rich(text, color=None, size=None, bold=None):
    """将 '普通[[强调]]文本' 或 '普通**粗体**文本' 解析为 Canvas.para 混排 runs。"""
    if isinstance(text, (list, tuple)):
        return text
    text = "" if text is None else str(text)
    runs, pos = [], 0
    for m in _TAG_RE.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], {}))
        opt = {"b": True if bold is None else bold}
        if color is not None:
            opt["c"] = color
        if size is not None:
            opt["sz"] = size
        highlight_text = m.group(1) if m.group(1) is not None else m.group(2)
        runs.append((highlight_text, opt))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], {}))
    return runs or [("", {})]


class Deck:
    def __init__(self, theme="gov_blue", size="16:9", font_cjk=None,
                 font_latin=None, meta=None):
        if isinstance(theme, dict):
            self.theme_name = theme.get("label", "custom_theme")
        else:
            self.theme_name = theme
        self._default_pal = T.get_palette(theme)
        self.pal = self._default_pal
        self.font_cjk = font_cjk or T.FONT_CJK_DEFAULT
        self.font_latin = font_latin or T.FONT_LATIN_DEFAULT
        self.sw, self.sh = T.SLIDE_SIZES.get(size, T.SLIDE_SIZES["16:9"])
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.sw)
        self.prs.slide_height = Inches(self.sh)
        self.warnings = []
        self.total_pages = 1
        self.meta = meta or {}
        self._current_page = None
        self._apply_meta(meta)

    def _apply_meta(self, meta):
        meta = meta or {}
        cp = self.prs.core_properties
        for key, attr in (("title", "title"), ("author", "author"),
                          ("subject", "subject"), ("comments", "comments"),
                          ("category", "category"), ("keywords", "keywords")):
            if meta.get(key):
                setattr(cp, attr, str(meta[key]))
        if not cp.title and meta.get("title"):
            cp.title = str(meta["title"])

    def _new_canvas(self, bg=True, bg_color=None, default_bg_mode="default"):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = getattr(self, "_current_page", None) or {}

        # 1. 优先支持单页指定独立主题 (page.theme)
        page_theme = page.get("theme")
        base_pal = T.get_palette(page_theme) if page_theme else self._default_pal

        # 2. 支持页面自适应色彩节奏模式 (page.bg_mode / page.style)
        bg_mode = page.get("bg_mode") or page.get("style") or default_bg_mode
        custom_bg = page.get("bg_color") or bg_color

        slide_pal = T.adapt_palette(base_pal, bg_mode=bg_mode, custom_bg=custom_bg)
        self.pal = slide_pal  # 保证全局 self.pal 与当前页调色板深度同步

        fill_color = custom_bg or slide_pal.get("bg")
        if bg and fill_color:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = T.rgb(fill_color)
        return Canvas(slide, slide_pal, self.font_cjk, self.font_latin)

    def _warn(self, page_idx, msg):
        self.warnings.append({"page": page_idx, "message": msg})

    def _add_page_tracker(self, cv, idx):
        """右下角轻量页码胶囊。"""
        if self.total_pages > 1 and idx > 0:
            text = "%02d / %02d" % (idx + 1, self.total_pages)
            px = self.sw - L.MARGIN - 1.05
            py = self.sh - 0.44
            cv.badge(px, py, text, dot=False,
                     bg=cv.pal.get("badge_bg"),
                     text_color=cv.pal.get("badge_text", cv.pal.get("text_subtitle")),
                     size=8.5, h=0.25, pad_x=0.10)

    # ---------------- 核心版式渲染器 ----------------

    def page_title_slide(self, idx, page):
        """现代化封面页：支持经典左右分栏卡片、居中全景、战略三立柱、极简杂志风、色块分屏等多变体自适应。"""
        variant = page.get("variant") or page.get("style")
        if variant in ("centered", "keynote", "summit"):
            return self.page_cover_centered(idx, page)
        elif variant in ("pillars", "cards", "bottom_cards", "strategy"):
            return self.page_cover_pillars(idx, page)
        elif variant in ("minimal", "editorial", "magazine"):
            return self.page_cover_minimal(idx, page)
        elif variant in ("split", "hero", "contrast"):
            return self.page_cover_split(idx, page)

        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh

        # 顶部品牌发光微线
        cv.rect(0, 0, sw, 0.045, fill=pal.get("primary"))
        cv.rect(0, 0, 2.4, 0.055, fill=pal.get("accent"))

        right_w = 4.6
        right_x = sw - L.MARGIN - right_w
        left_w = right_x - L.MARGIN - 0.50
        right_y = 1.25
        right_h = sh - 2.50

        # 类别标签
        cat = page.get("category") or self.meta.get("subject") or "专题汇报与展示"
        cv.badge(L.MARGIN, 1.45, str(cat), dot=True, size=10, h=0.32)

        # 主标题
        title_text = str(page.get("title", ""))
        title_size = 33 if len(title_text) > 13 else 37
        cv.text(L.MARGIN, 2.05, left_w, 1.35, title_text,
                size=title_size, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.12)

        # 副标题
        subtitle_text = page.get("subtitle")
        if subtitle_text:
            cv.text(L.MARGIN, 3.48, left_w, 0.60, str(subtitle_text),
                    size=16, color=pal.get("text_subtitle", pal["text_light"]), line=1.2)

        # 核心强调色微饰条
        cv.pill(L.MARGIN, 4.18, 1.35, 0.045, fill=pal.get("accent"))

        # 元信息胶囊条
        meta_line = page.get("meta_line") or ""
        if not meta_line and self.meta.get("author"):
            meta_line = "主讲 / 汇报：%s" % self.meta.get("author")
        if meta_line:
            cv.badge(L.MARGIN, 4.45, meta_line, dot=False,
                     bg=pal.get("card_subtle"),
                     text_color=pal.get("text_body"),
                     size=9.5, h=0.32, pad_x=0.16)

        # 右侧卡片区域：若有 highlights 或 bullets 则展示核心要点，若无则展示全景信息卡
        bullets = page.get("highlights") or page.get("bullets") or []
        cv.card(right_x, right_y, right_w, right_h, radius=0.10)

        if bullets:
            cv.badge(right_x + 0.35, right_y + 0.30, "核心要点 · HIGHLIGHTS",
                     dot=True, size=9.5, h=0.28)
            accents = pal.get("accents") or [pal.get("accent")]
            n_items = min(len(bullets), 4)
            card_item_h = (right_h - 1.15) / float(n_items)
            for i, b in enumerate(bullets[:n_items]):
                by = right_y + 0.88 + i * card_item_h
                sub_card_h = card_item_h - 0.16
                item_acc = accents[i % len(accents)]
                cv.card(right_x + 0.30, by, right_w - 0.60, sub_card_h,
                        fill=pal.get("card_subtle"),
                        line=pal.get("card_border"), lw=0.5, radius=0.06)

                num_str = "%02d" % (i + 1)
                b_end = cv.badge(right_x + 0.45, by + (sub_card_h - 0.28) / 2.0, num_str,
                                 dot=False, bg=pal.get("badge_bg"), text_color=item_acc,
                                 size=9, h=0.28, pad_x=0.10)

                text_x = b_end + 0.14
                text_w = max(0.5, (right_x + right_w - 0.45) - text_x)
                cv.text(text_x, by + 0.08, text_w, sub_card_h - 0.16,
                        rich(b, color=item_acc),
                        size=11, bold=True, color=pal.get("text_title"),
                        anchor=MSO_ANCHOR.MIDDLE)
        else:
            cv.badge(right_x + 0.35, right_y + 0.30, "方案概览 · OVERVIEW",
                     dot=True, size=9.5, h=0.28)
            cv.divider(right_x + 0.35, right_y + 0.72, right_w - 0.70)

            info_items = [
                ("战略定位", title_text[:14]),
                ("主讲团队", page.get("meta_line") or self.meta.get("author") or "核心专项工作组"),
                ("实施周期", "2025 - 2026 年度"),
                ("文档密级", "内部审阅 / 重点工作"),
            ]
            item_h = (right_h - 1.15) / 4.0
            accents = pal.get("accents") or [pal.get("accent")]
            for i, (k, v) in enumerate(info_items):
                by = right_y + 0.88 + i * item_h
                sub_card_h = item_h - 0.16
                item_acc = accents[i % len(accents)]
                cv.card(right_x + 0.30, by, right_w - 0.60, sub_card_h,
                        fill=pal.get("card_subtle"), lw=0.5, radius=0.06)

                cv.text(right_x + 0.45, by + 0.08, 1.20, sub_card_h - 0.16, k,
                        size=10, bold=True, color=item_acc, anchor=MSO_ANCHOR.MIDDLE)
                cv.text(right_x + 1.75, by + 0.08, right_w - 2.15, sub_card_h - 0.16, str(v),
                        size=10.5, color=pal.get("text_title"), anchor=MSO_ANCHOR.MIDDLE)

        return cv

    def page_cover_centered(self, idx, page):
        """极简磅礴·居中发布会与峰会封面页：大开大合，气势恢宏。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        accents = pal.get("accents") or [acc]

        L.chrome(cv, sw, sh)

        # 背景景深圆环装饰
        cv.circle(sw / 2.0 - 4.5, 0.2, 9.0, line=pal.get("card_border"), lw=0.5)
        cv.circle(sw / 2.0 - 2.5, 1.2, 5.0, line=pal.get("card_border"), lw=0.4)

        # 居中分类标签
        cat = page.get("category") or self.meta.get("subject") or "KEYNOTE 2026"
        cat_str = str(cat)
        approx_w = max(1.4, sum(2 if ord(c) > 127 else 1 for c in cat_str) * 0.09 + 0.6)
        cv.badge((sw - approx_w) / 2.0, 1.45, cat_str, dot=True, dot_color=acc, size=10, h=0.32)

        # 居中超大主标题
        title_text = str(page.get("title", ""))
        title_size = 40 if len(title_text) <= 12 else (34 if len(title_text) <= 22 else 28)
        cv.text(L.MARGIN, 2.15, sw - 2 * L.MARGIN, 1.50, title_text,
                size=title_size, bold=True, color=pal.get("text_title", pal["primary"]),
                align="center", line=1.12)

        # 居中副标题
        subtitle_text = page.get("subtitle")
        sub_y = 3.85
        if subtitle_text:
            cv.text(L.MARGIN + 1.0, sub_y, sw - 2 * L.MARGIN - 2.0, 0.60, str(subtitle_text),
                    size=16, color=pal.get("text_subtitle", pal["text_light"]), align="center", line=1.2)
            sub_y += 0.70

        # 居中核心强调微线
        cv.pill((sw - 1.8) / 2.0, sub_y, 1.8, 0.045, fill=acc)

        # 要点卡片或元信息区域
        bullets = page.get("highlights") or page.get("bullets") or []
        if bullets:
            n_b = min(len(bullets), 4)
            card_gap = 0.22
            total_w = min(11.4, sw - 2 * L.MARGIN)
            card_w = (total_w - card_gap * (n_b - 1)) / float(n_b)
            start_x = (sw - total_w) / 2.0
            by = sub_y + 0.35
            bh = sh - by - 0.95

            for i, b in enumerate(bullets[:n_b]):
                cx = start_x + i * (card_w + card_gap)
                item_acc = accents[i % len(accents)]
                cv.card(cx, by, card_w, bh, fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.6, radius=0.08)
                cv.pill(cx + 0.16, by + 0.04, min(1.0, card_w - 0.32), 0.03, fill=item_acc)

                num_str = "%02d" % (i + 1)
                b_end = cv.badge(cx + 0.16, by + 0.16, num_str, dot=False,
                                 bg=pal.get("badge_bg"), text_color=item_acc, size=8.5, h=0.25, pad_x=0.08)

                tf = cv.tbox(cx + 0.16, by + 0.48, card_w - 0.32, bh - 0.56, anchor=MSO_ANCHOR.TOP)
                cv.para(tf, rich(str(b), color=item_acc), size=10, color=pal.get("text_body"), line=1.22, first=True)

            meta_y = by + bh + 0.20
        else:
            meta_y = sub_y + 0.55

        # 底部元信息条
        meta_line = page.get("meta_line") or ""
        if not meta_line and self.meta.get("author"):
            meta_line = f"主讲人：{self.meta.get('author')}    发布时间：2026 年度"
        if meta_line:
            mw = max(3.4, sum(2 if ord(c) > 127 else 1 for c in meta_line) * 0.085 + 0.8)
            cv.badge((sw - mw) / 2.0, meta_y, meta_line, dot=False,
                     bg=pal.get("card_subtle"), text_color=pal.get("text_body"), size=9.5, h=0.32, pad_x=0.18)

        return cv

    def page_cover_pillars(self, idx, page):
        """战略支柱与多矩阵封面页：上半部宏大命题，下半部 3~4 大支柱立卡。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        accents = pal.get("accents") or [acc]

        # 顶部品牌线
        cv.rect(0, 0, sw, 0.045, fill=pal.get("primary"))
        cv.rect(0, 0, 3.2, 0.055, fill=acc)

        # 上半部：左侧标题与副标 + 右侧使命/愿景卡
        left_w = 7.6
        cat = page.get("category") or self.meta.get("subject") or "专题规划与设计"
        cv.badge(L.MARGIN, 1.25, str(cat), dot=True, size=10, h=0.32)

        title_text = str(page.get("title", ""))
        cv.text(L.MARGIN, 1.70, left_w, 1.15, title_text,
                size=32 if len(title_text) > 12 else 36, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.12)

        subtitle_text = page.get("subtitle")
        if subtitle_text:
            cv.text(L.MARGIN, 2.92, left_w, 0.45, str(subtitle_text),
                    size=14.5, color=pal.get("text_subtitle", pal["text_light"]), line=1.2)

        meta_line = page.get("meta_line") or (f"主讲团队：{self.meta.get('author')}" if self.meta.get('author') else "项目汇报团队")
        cv.badge(L.MARGIN, 3.45, meta_line, dot=False,
                 bg=pal.get("card_subtle"), text_color=pal.get("text_body"), size=9, h=0.28, pad_x=0.12)

        # 右上方：核心使命/导读盒
        right_w = sw - left_w - 2 * L.MARGIN - 0.35
        right_x = sw - L.MARGIN - right_w
        cv.card(right_x, 1.25, right_w, 2.50, fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.6, radius=0.08)
        mission_badge = page.get("mission_title") or "核心定位与目标"
        cv.badge(right_x + 0.25, 1.45, mission_badge, dot=False,
                 bg=pal.get("badge_bg"), text_color=acc, size=8.8, h=0.26)
        mission_text = (page.get("abstract") or page.get("vision") or page.get("lead") or
                        "以全面创新为驱动，夯实底层基座，实现全周期跨越式增长与价值赋能。")
        cv.text(right_x + 0.25, 1.85, right_w - 0.50, 1.75,
                rich(mission_text, color=acc), size=10.5, color=pal.get("text_body"), line=1.28)

        # 下半部：3~4 个战略立柱卡片
        pillars = (page.get("pillars") or page.get("cards") or
                   page.get("highlights") or page.get("bullets") or [
                       {"title": "技术底座筑基", "desc": "构建自主高可用基础架构，保障底层敏捷演进与性能突围"},
                       {"title": "业务场景赋能", "desc": "深度链接终端核心业务痛点，沉淀高质量全链路解决方案"},
                       {"title": "组织效能跃迁", "desc": "打通跨职能协作协同链路，实现战略目标精准敏捷拆解"}
                   ])

        n_p = min(len(pillars), 4)
        p_gap = 0.24
        p_w = (L.CONTENT_W - p_gap * (n_p - 1)) / float(n_p)
        p_y = 3.95
        p_h = sh - p_y - 0.65

        for i, pil in enumerate(pillars[:n_p]):
            px = L.MARGIN + i * (p_w + p_gap)
            pil_acc = accents[i % len(accents)]
            cv.card(px, p_y, p_w, p_h, fill=pal.get("card_bg"), line=pal.get("card_border"), lw=0.6, radius=0.08)
            cv.pill(px + 0.20, p_y + 0.05, min(1.20, p_w - 0.40), 0.035, fill=pil_acc)

            pil_dict = pil if isinstance(pil, dict) else {"title": str(pil)}
            p_idx_str = "PILLAR %02d" % (i + 1)
            b_end = cv.badge(px + 0.20, p_y + 0.18, p_idx_str, dot=False,
                             bg=pal.get("badge_bg"), text_color=pil_acc, size=8.5, h=0.25, pad_x=0.08)

            p_title = str(pil_dict.get("title") or f"支柱 {i+1}")
            cv.text(px + 0.20, p_y + 0.52, p_w - 0.40, 0.40, p_title,
                    size=12.5, bold=True, color=pal.get("text_title", pal["primary"]), line=1.10)

            cv.divider(px + 0.20, p_y + 0.98, p_w - 0.40)

            p_items = pil_dict.get("items") or pil_dict.get("bullets")
            p_desc = pil_dict.get("desc") or pil_dict.get("summary")
            cy = p_y + 1.10
            ch = p_h - 1.25

            if p_items:
                L.bullet_list(cv, px + 0.20, cy, p_w - 0.40, ch,
                              p_items, size=9.8, gap=5, line=1.22, dot_color=pil_acc)
            elif p_desc:
                tf = cv.tbox(px + 0.20, cy, p_w - 0.40, ch)
                cv.para(tf, rich(str(p_desc), color=pil_acc), size=10,
                        color=pal.get("text_body"), line=1.26, first=True)

        return cv

    def page_cover_minimal(self, idx, page):
        """前沿杂志风与学术研报封面页：极简非对称留白，典雅克制。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])

        # 巨大年份/版本背景浮印
        year_str = str(page.get("year") or "2026")
        cv.text(sw - L.MARGIN - 3.4, 1.05, 3.4, 1.8, year_str,
                size=68, bold=True, color=pal.get("card_subtle", pal.get("bg_alt")), align="right", line=0.9)

        # 分类微标
        cat = page.get("category") or self.meta.get("subject") or "SPECIAL RESEARCH REPORT"
        cv.badge(L.MARGIN, 1.40, str(cat), dot=True, size=10, h=0.32)

        # 超大优雅主标题
        title_text = str(page.get("title", ""))
        cv.text(L.MARGIN, 2.15, sw - 2 * L.MARGIN - 2.0, 1.55, title_text,
                size=36 if len(title_text) > 13 else 40, bold=True,
                color=pal.get("text_title", pal["primary"]), line=1.12)

        # 纤细横向分割线贯穿
        cv.divider(L.MARGIN, 3.90, sw - 2 * L.MARGIN)

        # 左下区：副标题与作者元信息
        subtitle_text = page.get("subtitle") or "探索未知边界 · 洞见前瞻趋势"
        cv.text(L.MARGIN, 4.25, 6.2, 0.85, str(subtitle_text),
                size=16, color=pal.get("text_subtitle", pal["text_light"]), line=1.25)

        meta_line = page.get("meta_line") or (f"研究团队：{self.meta.get('author') or '核心前瞻实验室'}\n密级：内部研阅 · 核心资产")
        cv.text(L.MARGIN, 5.35, 6.2, 0.80, str(meta_line),
                size=10.5, color=pal.get("text_muted"), line=1.35)

        # 右下区：沉浸式精装摘要浮岛盒
        abstract_w = 4.8
        abstract_x = sw - L.MARGIN - abstract_w
        abstract_y = 4.20
        abstract_h = 2.45
        cv.card(abstract_x, abstract_y, abstract_w, abstract_h,
                fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.6, radius=0.08)

        cv.badge(abstract_x + 0.28, abstract_y + 0.20, "★ EXECUTIVE SUMMARY", dot=False,
                 bg=pal.get("badge_bg"), text_color=acc, size=9, h=0.26)

        abstract_text = (page.get("abstract") or page.get("lead") or page.get("summary") or
                         "本报告立足全球最新实践，全方位系统梳理核心架构演进路线与落地方案，为下一步重大战略决策提供权威理论依据与实操范式。")
        cv.text(abstract_x + 0.28, abstract_y + 0.58, abstract_w - 0.56, abstract_h - 0.70,
                rich(abstract_text, color=acc), size=10.5, color=pal.get("text_body"), line=1.30)

        return cv

    def page_cover_split(self, idx, page):
        """左右色彩碰撞·沉浸式分屏封面页：左侧高饱和深底，右侧浮岛亮点。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        accents = pal.get("accents") or [acc]

        # 左屏高饱和主屏（宽 52%）
        split_w = sw * 0.52
        left_bg = pal.get("primary_mid", pal.get("primary"))
        cv.rect(0, 0, split_w, sh, fill=left_bg)
        cv.rect(split_w - 0.05, 0, 0.05, sh, fill=acc)

        # 左屏文字
        cat = page.get("category") or self.meta.get("subject") or "GLOBAL ROADSHOW"
        cv.badge(L.MARGIN, 1.55, str(cat), dot=True, dot_color=acc,
                 bg=pal.get("primary", left_bg), text_color=T.rgb("FFFFFF"), size=10, h=0.32)

        title_text = str(page.get("title", ""))
        cv.text(L.MARGIN, 2.25, split_w - 2 * L.MARGIN, 1.70, title_text,
                size=34 if len(title_text) > 13 else 38, bold=True,
                color=T.rgb("FFFFFF"), line=1.12)

        subtitle_text = page.get("subtitle")
        if subtitle_text:
            cv.text(L.MARGIN, 4.15, split_w - 2 * L.MARGIN, 0.75, str(subtitle_text),
                    size=15, color=pal.get("accent_tint", T.rgb("E0E0E0")), line=1.25)

        meta_line = page.get("meta_line") or (f"主讲团队：{self.meta.get('author') or '核心创始团队'}\n日期：2026 年度")
        cv.text(L.MARGIN, 5.60, split_w - 2 * L.MARGIN, 0.75, str(meta_line),
                size=11, color=T.rgb("D0D0D0"), line=1.3)

        # 右屏浮岛亮点区域（宽 48%）
        rx = split_w + 0.40
        rw = sw - rx - L.MARGIN
        cv.badge(rx, 1.55, "● 核心看点 / KEY HIGHLIGHTS", dot=True, dot_color=acc, size=9.5, h=0.30)

        bullets = page.get("highlights") or page.get("bullets") or [
            "突破性技术演进：底层吞吐跃升 300%",
            "行业级生态赋能：全流程标准化闭环落地",
            "全球化战略纵深：跨区域协同高效赋能"
        ]

        n_b = min(len(bullets), 4)
        b_gap = 0.18
        b_h = (sh - 2.80 - b_gap * (n_b - 1)) / float(n_b)

        for i, b in enumerate(bullets[:n_b]):
            by = 2.20 + i * (b_h + b_gap)
            item_acc = accents[i % len(accents)]
            cv.card(rx, by, rw, b_h, fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.6, radius=0.08)
            cv.pill(rx + 0.20, by + 0.05, min(1.2, rw - 0.40), 0.035, fill=item_acc)

            num_str = "%02d" % (i + 1)
            b_end = cv.badge(rx + 0.20, by + (b_h - 0.28) / 2.0, num_str, dot=False,
                             bg=pal.get("badge_bg"), text_color=item_acc, size=9, h=0.28, pad_x=0.08)

            tf = cv.tbox(b_end + 0.14, by + 0.08, rw - (b_end - rx) - 0.30, b_h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
            cv.para(tf, rich(str(b), color=item_acc), size=10.5, bold=True,
                    color=pal.get("text_title", pal["primary"]), line=1.22, first=True)

        return cv

    def page_section(self, idx, page):
        """章节过渡页：大牌感章节标号 + 议题导航概览 + 视觉节奏深邃转场。"""
        variant = page.get("variant") or page.get("style")
        if variant in ("centered", "minimal"):
            return self.page_section_centered(idx, page)
        elif variant in ("split", "overview", "side"):
            return self.page_section_split(idx, page)
        elif variant in ("progress", "stepper", "timeline"):
            return self.page_section_progress(idx, page)

        # 默认采用 contrast 模式自适应反转，形成明暗交替的破局呼吸感
        cv = self._new_canvas(default_bg_mode="contrast")
        pal = cv.pal
        sw, sh = self.sw, self.sh
        accents = pal.get("accents") or [pal.get("accent")]
        sec_acc = accents[idx % len(accents)]

        L.chrome(cv, sw, sh)

        # 章节序号处理
        sec_raw = str(page.get("number") or page.get("chapter") or (idx + 1))
        if sec_raw.isdigit():
            num_display = "%02d" % int(sec_raw)
        else:
            num_display = sec_raw.strip()

        # 1. 巨大半透明罗马/阿拉伯数字水印（杂志封面景深感，不遮挡主体）
        is_dark = pal.get("is_dark", False)
        if is_dark:
            watermark_fill = pal.get("card_bg", pal.get("primary_mid"))
        else:
            watermark_fill = pal.get("card_subtle", pal.get("bg_alt"))
        cv.text(sw - L.MARGIN - 4.6, 1.10, 4.6, 3.8, num_display,
                size=120, bold=True, color=watermark_fill, align="right", line=0.88)

        # 2. 章节微标胶囊
        lx = L.MARGIN + 0.35
        badge_text = "CHAPTER %s" % num_display if num_display.isdigit() else num_display
        cat = page.get("category")
        if cat:
            badge_text += " · %s" % cat
        badge_tc = T.rgb("FFFFFF") if is_dark else sec_acc
        badge_dot = sec_acc if not is_dark else T.rgb("FFFFFF")
        cv.badge(lx, 1.60, badge_text, dot=True,
                 dot_color=badge_dot, text_color=badge_tc, size=10.5, h=0.34)

        # 3. 章节主标题（字号自适应长短文本，杜绝孤字成行）
        sec_title = str(page.get("title", ""))
        title_len = sum(1.8 if ord(ch) > 127 else 1.0 for ch in sec_title)
        sec_sz = 35 if title_len <= 16 else (29 if title_len <= 26 else 24)

        lw = 8.5
        cv.text(lx, 2.20, lw, 1.25, sec_title,
                size=sec_sz, bold=True, color=pal.get("text_title"), line=1.12)

        # 4. 章节副标题
        subtitle_text = page.get("subtitle")
        if subtitle_text:
            cv.text(lx, 3.55, lw, 0.75, str(subtitle_text),
                    size=15, color=pal.get("text_subtitle", pal["text_light"]), line=1.25)

        # 5. 核心品牌强调微线
        pill_y = 4.40 if subtitle_text else 3.75
        cv.pill(lx, pill_y, 1.6, 0.05, fill=sec_acc)

        # 6. 下部内容区：议题清单 (agenda / topics / items) 或 章节导读 (lead / summary)
        topics = page.get("agenda") or page.get("topics") or page.get("items") or []
        lead_text = page.get("lead") or page.get("summary")

        if topics and isinstance(topics, list):
            n = min(len(topics), 4)
            card_gap = 0.22
            total_w = sw - lx - L.MARGIN - 0.35
            card_w = (total_w - card_gap * (n - 1)) / float(n)
            card_y = 4.85
            card_h = 1.65

            for i, topic in enumerate(topics[:n]):
                cx = lx + i * (card_w + card_gap)
                cv.card(cx, card_y, card_w, card_h,
                        fill=pal.get("card_bg"), line=pal.get("card_border"), lw=0.6, radius=0.08)

                if isinstance(topic, dict):
                    t_title = topic.get("title", "")
                    t_desc = topic.get("desc", "")
                else:
                    t_title = str(topic)
                    t_desc = ""

                sub_acc = accents[i % len(accents)]
                badge_num_tc = T.rgb("FFFFFF") if is_dark else sub_acc
                cv.badge(cx + 0.25, card_y + 0.22, "%02d" % (i + 1),
                         dot=False, bg=pal.get("badge_bg"), text_color=badge_num_tc, size=9.5, h=0.28)

                cv.text(cx + 0.25, card_y + 0.62, card_w - 0.50, 0.45,
                        t_title, size=11.5, bold=True, color=pal.get("text_title"), line=1.15)
                if t_desc:
                    cv.text(cx + 0.25, card_y + 1.10, card_w - 0.50, 0.45,
                            t_desc, size=9.5, color=pal.get("text_muted", pal.get("text_subtitle")), line=1.2)

        elif lead_text:
            lead_y = 4.75
            lead_w = sw - lx - L.MARGIN - 0.35
            lead_h = 1.45
            cv.card(lx, lead_y, lead_w, lead_h,
                    fill=pal.get("card_bg"), line=pal.get("card_border"), lw=0.6, radius=0.08)
            cv.badge(lx + 0.35, lead_y + 0.22, "本章导读 · SECTION SUMMARY",
                     dot=True, dot_color=sec_acc, size=9.5, h=0.28)
            cv.text(lx + 0.35, lead_y + 0.62, lead_w - 0.70, 0.68,
                    rich(lead_text, color=sec_acc),
                    size=12, color=pal.get("text_body"), line=1.3)

        self._add_page_tracker(cv, idx)
        return cv

    def page_agenda(self, idx, page):
        """全局大目录导航大盘页：全篇 3~6 章节全景骨架导航。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)

        title = page.get("title") or "目录与议程导览"
        subtitle = page.get("subtitle") or "AGENDA & TABLE OF CONTENTS"
        L.header(cv, title, subtitle)

        chapters = (page.get("chapters") or page.get("sections") or
                    page.get("agenda") or page.get("items") or [
                        {"title": "背景概述与核心目标", "desc": "全局背景脉络梳理与核心诉求目标界定"},
                        {"title": "总体方案与核心设计", "desc": "端到端系统架构设计与核心模块全景图"},
                        {"title": "关键举措与具体实施", "desc": "重点推进方案分解与分阶段实施路径"},
                        {"title": "总结复盘与延伸思考", "desc": "核心成果价值沉淀与后续长效推进方向"}
                    ])

        y = 1.35
        h = sh - y - 0.65
        L.agenda_grid(cv, y, chapters, h=h)

        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_section_centered(self, idx, page):
        """沉浸式极简居中章节过渡页：Contrast 反差底色，纯粹聚焦。"""
        cv = self._new_canvas(default_bg_mode="contrast")
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        L.chrome(cv, sw, sh)

        sec_raw = str(page.get("number") or page.get("chapter") or (idx + 1))
        num_display = "%02d" % int(sec_raw) if sec_raw.isdigit() else sec_raw.strip()

        # 居中章节徽标
        badge_text = "CHAPTER %s" % num_display if num_display.isdigit() else num_display
        cat = page.get("category")
        if cat:
            badge_text += " · %s" % cat
        approx_w = max(1.6, sum(2 if ord(c) > 127 else 1 for c in badge_text) * 0.09 + 0.6)
        cv.badge((sw - approx_w) / 2.0, 1.80, badge_text, dot=True,
                 dot_color=acc, text_color=T.rgb("FFFFFF") if pal.get("is_dark") else acc, size=10.5, h=0.34)

        # 居中大主标题
        sec_title = str(page.get("title", ""))
        cv.text(L.MARGIN, 2.45, sw - 2 * L.MARGIN, 1.45, sec_title,
                size=38 if len(sec_title) <= 12 else 32, bold=True,
                color=pal.get("text_title"), align="center", line=1.12)

        # 居中副标题
        subtitle_text = page.get("subtitle")
        sub_y = 4.05
        if subtitle_text:
            cv.text(L.MARGIN + 1.0, sub_y, sw - 2 * L.MARGIN - 2.0, 0.60, str(subtitle_text),
                    size=16, color=pal.get("text_subtitle", pal["text_light"]), align="center", line=1.2)
            sub_y += 0.70

        cv.pill((sw - 1.8) / 2.0, sub_y, 1.8, 0.045, fill=acc)

        # 居中金句/导读盒
        quote = page.get("quote") or page.get("lead") or page.get("summary") or page.get("question") or ""
        if quote:
            box_w = min(9.2, sw - 2 * L.MARGIN)
            cv.card((sw - box_w) / 2.0, sub_y + 0.35, box_w, 1.40,
                    fill=pal.get("card_bg"), line=pal.get("card_border"), lw=0.6, radius=0.08)
            cv.text((sw - box_w) / 2.0 + 0.40, sub_y + 0.50, box_w - 0.80, 1.10,
                    rich(quote, color=acc), size=12.5, bold=True,
                    color=pal.get("text_body"), align="center", line=1.3)

        self._add_page_tracker(cv, idx)
        return cv

    def page_section_split(self, idx, page):
        """左右分栏式章节全景转场页：左侧命题，右侧核心成果与研讨议题。"""
        cv = self._new_canvas(default_bg_mode="contrast")
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        L.chrome(cv, sw, sh)

        left_w = 4.6
        right_w = sw - left_w - 2 * L.MARGIN - 0.40
        right_x = sw - L.MARGIN - right_w

        sec_raw = str(page.get("number") or page.get("chapter") or (idx + 1))
        num_display = "%02d" % int(sec_raw) if sec_raw.isdigit() else sec_raw.strip()

        # 左栏：章节标号 + 标题 + 导读
        cv.badge(L.MARGIN, 1.40, "CHAPTER %s" % num_display, dot=True, dot_color=acc, size=10.5, h=0.32)

        sec_title = str(page.get("title", ""))
        cv.text(L.MARGIN, 2.00, left_w, 1.35, sec_title,
                size=30, bold=True, color=pal.get("text_title"), line=1.15)

        if page.get("subtitle"):
            cv.text(L.MARGIN, 3.50, left_w, 0.60, str(page["subtitle"]),
                    size=14.5, color=pal.get("text_subtitle"), line=1.2)

        cv.pill(L.MARGIN, 4.25, 1.5, 0.045, fill=acc)

        lead_str = page.get("lead") or page.get("summary") or "本章重点剖析核心技术架构与实施落地细则，确保各项战术指标高效达成。"
        cv.text(L.MARGIN, 4.55, left_w, 2.0, rich(lead_str, color=acc),
                size=11, color=pal.get("text_body"), line=1.35)

        # 右栏：两大精装卡片（上：成果交付，下：议题清单）
        card_h = (sh - 1.40 - 0.70 - 0.25) / 2.0

        # 上卡：本章核心成果
        top_y = 1.40
        cv.card(right_x, top_y, right_w, card_h, fill=pal.get("card_bg"), line=pal.get("card_border"), lw=0.6, radius=0.08)
        cv.pill(right_x + 0.22, top_y + 0.05, min(1.2, right_w - 0.44), 0.035, fill=acc)
        cv.badge(right_x + 0.22, top_y + 0.18, "★ 本章核心目标与收获 / KEY DELIVERABLES", dot=False,
                 bg=pal.get("badge_bg"), text_color=acc, size=9, h=0.28)

        deliverables = page.get("deliverables") or page.get("outcomes") or page.get("highlights") or [
            "清晰掌握全链路系统交互时序与瓶颈根因",
            "完成核心模块架构重构与高可用方案评审",
            "沉淀标准化的线上工程化实施规范"
        ]
        L.bullet_list(cv, right_x + 0.22, top_y + 0.62, right_w - 0.44, card_h - 0.75,
                      deliverables, size=10.5, gap=6, line=1.26, dot_color=acc)

        # 下卡：议题清单
        bot_y = top_y + card_h + 0.25
        cv.card(right_x, bot_y, right_w, card_h, fill=pal.get("card_subtle"), line=pal.get("card_border"), lw=0.6, radius=0.08)
        cv.badge(right_x + 0.22, bot_y + 0.18, "本章重点议题 / AGENDA TOPICS", dot=True,
                 dot_color=acc, text_color=pal.get("text_title"), size=9, h=0.28)

        topics = page.get("topics") or page.get("agenda") or page.get("items") or [
            "关键架构选型与权衡比对",
            "性能压测与全链路容灾指标",
            "阶段里程碑与排期计划"
        ]
        L.bullet_list(cv, right_x + 0.22, bot_y + 0.62, right_w - 0.44, card_h - 0.75,
                      topics, size=10.5, gap=6, line=1.26, dot_color=acc)

        self._add_page_tracker(cv, idx)
        return cv

    def page_section_progress(self, idx, page):
        """全景进度点亮章节转场页：贯穿全篇章节进度轴，高亮当前阶段。"""
        cv = self._new_canvas(default_bg_mode="contrast")
        pal = cv.pal
        sw, sh = self.sw, self.sh
        acc = pal.get("accent", pal["primary"])
        L.chrome(cv, sw, sh)

        sec_raw = str(page.get("number") or page.get("chapter") or (idx + 1))
        num_display = "%02d" % int(sec_raw) if sec_raw.isdigit() else sec_raw.strip()
        curr_i = int(sec_raw) - 1 if sec_raw.isdigit() else idx

        # 上半部：主标题与导读
        cv.badge(L.MARGIN, 1.35, "CHAPTER %s · 章节进展" % num_display, dot=True, dot_color=acc, size=10, h=0.32)

        sec_title = str(page.get("title", ""))
        cv.text(L.MARGIN, 1.95, sw - 2 * L.MARGIN, 1.15, sec_title,
                size=32, bold=True, color=pal.get("text_title"), line=1.12)

        subtitle_text = page.get("subtitle") or page.get("lead") or page.get("summary")
        if subtitle_text:
            cv.text(L.MARGIN, 3.20, sw - 2 * L.MARGIN, 0.55, str(subtitle_text),
                    size=15, color=pal.get("text_subtitle"), line=1.22)

        cv.pill(L.MARGIN, 3.85, 1.8, 0.045, fill=acc)

        # 下半部：全景章节进度轴
        chapters = page.get("chapters") or page.get("sections") or page.get("roadmap") or [
            {"title": "宏观态势与现状复盘"},
            {"title": "战略定位与核心突破"},
            {"title": "架构设计与技术攻坚"},
            {"title": "实施落地与资源协同"}
        ]

        stepper_y = 4.30
        stepper_h = sh - stepper_y - 0.70
        L.section_stepper(cv, stepper_y, chapters, current_idx=curr_i, h=stepper_h)

        self._add_page_tracker(cv, idx)
        return cv

    def page_dashboard(self, idx, page):
        """综合数据看板页：导读摘要 + Bento KPI + 复合多面板。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        y = 1.15
        lead = page.get("lead") or page.get("summary")
        if lead:
            L.lead_bar(cv, rich(lead, color=pal.get("accent", pal["primary"])), y=y, h=0.56)
            y += 0.68

        kpis = page.get("kpis") or page.get("metrics") or []
        if kpis:
            kpi_h = float(page.get("kpi_height", 1.42))
            y = L.kpi_row(cv, kpis, y=y, h=kpi_h) + 0.16

        panels = (page.get("panels") or page.get("cards") or
                  page.get("columns") or page.get("items") or [])
        if panels:
            bottom_limit = sh - 0.48
            ph = max(bottom_limit - y, 1.5)
            gap = 0.22
            n_p = len(panels)
            pw = (L.CONTENT_W - gap * (n_p - 1)) / float(n_p)

            for i, p in enumerate(panels):
                px = L.MARGIN + i * (pw + gap)
                panel_acc = accents[i % len(accents)]
                chart = p.get("chart")
                if chart:
                    cats = chart.get("categories") or chart.get("labels") or []
                    ser = chart.get("series")
                    if isinstance(ser, list) and ser and isinstance(ser[0], (int, float)):
                        ser = [{"name": chart.get("series_name", "数值"), "values": ser}]
                    elif not ser and chart.get("values"):
                        ser = [{"name": chart.get("series_name", "数值"), "values": chart["values"]}]

                    chart_spec = dict(chart)
                    chart_spec["categories"] = cats
                    chart_spec["series"] = ser

                    L.panel_with_chart(
                        cv, px, y, pw, ph, p.get("title", ""), p.get("subtitle"),
                        p.get("items"),
                        categories=cats,
                        series=ser,
                        chart_h=float(chart.get("height", 1.45)),
                        accent_color=panel_acc,
                        chart_spec=chart_spec)

                else:
                    bx, by, bw = L.panel(cv, px, y, pw, ph, p.get("title", ""),
                                         p.get("subtitle"), accent_color=panel_acc)

                    # 1. 优先检查并流式渲染顶置进度条（杜绝与下方列表重叠）
                    if p.get("progress") is not None:
                        from .charts import progress_bar
                        prog_obj = p["progress"]
                        prog_label = p.get("progress_label")
                        header_lbl = ""

                        if isinstance(prog_obj, dict):
                            val = prog_obj.get("val", prog_obj.get("value", 0))
                            total = prog_obj.get("total", 100)
                            prog_ratio = float(val) / float(total) if total > 0 else 0.5
                            header_lbl = str(prog_obj.get("label", ""))
                            prog_label = prog_label or ("%.0f%%" % (prog_ratio * 100))
                        elif isinstance(prog_obj, (int, float)):
                            prog_ratio = float(prog_obj) if prog_obj <= 1.0 else (float(prog_obj) / 100.0 if prog_obj <= 100 else 1.0)
                            prog_label = prog_label or ("%.0f%%" % (prog_ratio * 100))
                        else:
                            prog_ratio = 0.5
                            prog_label = prog_label or "50%"

                        if header_lbl:
                            cv.text(bx, by, bw, 0.22, header_lbl, size=9.8, bold=True,
                                    color=pal.get("text_title", pal["primary"]))
                            by += 0.24

                        progress_bar(cv, bx, by, bw, 0.16, prog_ratio,
                                     fill=panel_acc, label=prog_label)
                        by += 0.32

                    # 2. 渲染面板项目列表
                    items = p.get("items") or p.get("bullets") or []
                    if items:
                        is_all_metrics = all(isinstance(it, dict) and "label" in it and "value" in it for it in items)
                        if is_all_metrics and len(items) <= 5:
                            from .charts import progress_bar
                            item_space_h = (ph - (by - y) - 0.20) / float(len(items))

                            vals = []
                            for it in items:
                                v = it.get("value")
                                if isinstance(v, (int, float)):
                                    vals.append(float(v))
                                elif isinstance(v, str):
                                    v_clean = v.replace("%", "").strip()
                                    try:
                                        vals.append(float(v_clean))
                                    except ValueError:
                                        pass

                            panel_max = p.get("max") or p.get("max_value")
                            if panel_max is not None:
                                base_max = float(panel_max)
                            elif vals:
                                max_v = max(vals)
                                title_str = (p.get("title") or "") + (p.get("subtitle") or "")
                                is_explicit_pct = ("%" in title_str or "占比" in title_str or "比例" in title_str or "率" in title_str)
                                sum_v = sum(vals)
                                is_sum_100 = (95 <= sum_v <= 105)

                                if (is_explicit_pct or is_sum_100) and max_v <= 100:
                                    base_max = 100.0
                                elif max_v <= 1.0:
                                    base_max = 1.0
                                else:
                                    base_max = max_v if max_v > 0 else 1.0
                            else:
                                base_max = 1.0

                            for j, it in enumerate(items):
                                iy = by + j * item_space_h
                                it_val = it["value"]
                                num_val = float(it_val) if isinstance(it_val, (int, float)) else (
                                    float(str(it_val).replace("%", "").strip()) if str(it_val).replace("%", "").strip().replace(".", "", 1).isdigit() else 0.0
                                )
                                it_ratio = max(0.04, min(1.0, num_val / base_max)) if base_max > 0 else 0.5
                                if isinstance(it_val, float) and it_val <= 1.0 and base_max == 1.0:
                                    it_label = "%.0f%%" % (it_val * 100)
                                else:
                                    it_label = str(it_val)

                                cv.text(bx, iy, bw, 0.22, str(it["label"]), size=10, bold=True,
                                        color=pal.get("text_title", pal["primary"]))
                                item_acc = accents[j % len(accents)]
                                progress_bar(cv, bx, iy + 0.24, bw, 0.14, it_ratio,
                                             fill=item_acc, label=it_label)
                        else:
                            avail_h = max(0.40, ph - (by - y) - 0.15)
                            L.bullet_list(cv, bx, by, bw, avail_h, items,
                                          size=p.get("size", 10.5), dot_color=panel_acc)

        if page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_bullets(self, idx, page):
        """多栏结构化要点页：自适应砖块布局（Item Blocks），彻底消除大片空白。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh
        accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        cols = page.get("columns") or page.get("cols") or []
        start_y = 1.25

        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")

        if not cols:
            items = page.get("items") or page.get("bullets") or page.get("cards") or []
            n_items = len(items)
            # 自适应高度：字少时提供更充实大气的卡片高度与砖块间距，杜绝空洞
            if has_takeaway:
                card_h = min(4.2, max(3.0, 1.2 + n_items * 0.95))
            else:
                card_h = min(5.4, max(4.6 if n_items <= 3 else 3.8, 1.2 + n_items * 0.95))

            y_pos = start_y
            cv.card(L.MARGIN, y_pos, L.CONTENT_W, card_h)

            if n_items <= 6 and n_items > 0:
                block_w = L.CONTENT_W - 0.60
                item_h = (card_h - 0.60) / float(n_items)
                b_sz = 12.5 if n_items <= 3 else (11.5 if n_items <= 4 else 10.5)
                for i, it in enumerate(items):
                    it_acc = accents[i % len(accents)]
                    L.item_block(cv, L.MARGIN + 0.30, y_pos + 0.30 + i * item_h,
                                 block_w, max(0.40, item_h - 0.12), it, index=i + 1, size=b_sz,
                                 accent_color=it_acc)
            elif n_items > 6:
                avail_h = card_h - 0.60
                dyn_gap = max(6, min(14, int((avail_h * 72 - n_items * 14) / max(1, n_items - 1))))
                L.bullet_list(cv, L.MARGIN + 0.35, y_pos + 0.30, L.CONTENT_W - 0.70,
                              card_h - 0.50, items, size=page.get("size", 10.5), gap=dyn_gap)

            if has_takeaway:
                takeaway_y = min(sh - 1.35, y_pos + card_h + 0.22)
                takeaway_h = min(1.20, sh - takeaway_y - 0.50)
                L.takeaway_card(cv, takeaway_text, takeaway_y, title="落地保障与执行洞察", h=takeaway_h)
            elif page.get("footer"):
                L.footer_note(cv, sh, page["footer"])
        else:
            n = min(len(cols), 4)
            gap = 0.26
            cw = (L.CONTENT_W - gap * (n - 1)) / float(n)

            # 计算各列最大要点数
            max_items = max(len((c.get("items") or c.get("bullets") if isinstance(c, dict) else c) or []) for c in cols)
            card_h = min(5.4, max(4.6 if max_items <= 3 else 3.8, 1.2 + max_items * 0.88))
            y_pos = start_y

            for i, c in enumerate(cols[:n]):
                cx = L.MARGIN + i * (cw + gap)
                col_acc = accents[i % len(accents)]
                if isinstance(c, dict):
                    c_items = c.get("items") or c.get("bullets") or []
                    bx, by, bw = L.panel(cv, cx, y_pos, cw, card_h,
                                         c.get("title", ""), c.get("subtitle"),
                                         accent_color=col_acc)
                    if len(c_items) <= 4:
                        avail_h = card_h - (by - y_pos) - 0.20
                        it_h = avail_h / float(len(c_items)) if c_items else 0.8
                        b_sz = 12.0 if len(c_items) <= 2 else 10.5
                        for j, it in enumerate(c_items):
                            it_acc = accents[(i + j) % len(accents)]
                            L.item_block(cv, bx, by + j * it_h, bw, it_h - 0.10,
                                         it, index=j + 1, size=b_sz, accent_color=it_acc)
                    else:
                        L.bullet_list(cv, bx, by, bw, card_h - (by - y_pos) - 0.20,
                                      c_items, size=c.get("size", 10.5), gap=6, dot_color=col_acc)
                else:
                    cv.card(cx, y_pos, cw, card_h)
                    cv.pill(cx + 0.25, y_pos + 0.04, min(1.2, cw - 0.50), 0.03, fill=col_acc)
                    L.bullet_list(cv, cx + 0.25, y_pos + 0.30, cw - 0.50, card_h - 0.50,
                                  c, size=page.get("size", 11), gap=8, dot_color=col_acc)

            if has_takeaway:
                takeaway_y = min(sh - 1.35, y_pos + card_h + 0.22)
                takeaway_h = min(1.20, sh - takeaway_y - 0.50)
                L.takeaway_card(cv, takeaway_text, takeaway_y, title="落地保障与执行洞察", h=takeaway_h)
            elif page.get("note") or page.get("footer"):
                L.footer_note(cv, sh, page.get("note") or page.get("footer"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_grid(self, idx, page):
        """全新 2x2 或 1x3 Bento 矩阵网格页：4 大核心战略支柱或能力板块。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        cards = (page.get("cards") or page.get("items") or
                 page.get("columns") or page.get("panels") or
                 page.get("blocks") or [])
        cols = int(page.get("cols", 2))
        y = 1.25

        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")

        rows_cnt = (len(cards) + cols - 1) // max(1, cols) if cards else 1
        card_h = page.get("card_height")
        if not card_h:
            if rows_cnt == 1:
                card_h = 3.60 if has_takeaway else 4.80
            else:
                card_h = 2.15 if has_takeaway else 2.50

        grid_end_y = L.grid_cards(cv, y, cards, cols=cols, h=card_h)

        if has_takeaway:
            takeaway_y = min(sh - 1.35, max(grid_end_y + 0.22, sh - 1.45))
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="核心战略策略")
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_comparison(self, idx, page):
        """全新对比页：Before vs After / 挑战 vs 突破方案。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        left = page.get("left") or page.get("before") or {"title": "原有模式 / 痛点挑战", "items": []}
        right = page.get("right") or page.get("after") or {"title": "新代架构 / 核心突破", "items": []}
        y = 1.30

        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        comp_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.75)

        L.comparison_cards(cv, y, left, right, h=comp_h)

        if has_takeaway:
            takeaway_y = y + comp_h + 0.22
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="演进策略建议", h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_kpi(self, idx, page):
        """专属核心指标页：大数字震撼冲击力，完美自适应单行/多行，居中排版。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        y = 1.25
        lead = page.get("lead") or page.get("summary")
        if lead:
            L.lead_bar(cv, rich(lead, color=pal.get("accent")), y=y, h=0.58)
            y += 0.72

        kpis = (page.get("kpis") or page.get("metrics") or
                page.get("cards") or page.get("items") or [])
        cols = page.get("cols") or min(len(kpis), 4) or 1
        rows = [kpis[i:i + cols] for i in range(0, len(kpis), cols)]
        row_count = len(rows)

        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")

        if row_count == 1:
            t_title = page.get("takeaway_title") or "核心指标深度洞察与分析"
            if has_takeaway:
                kpi_h = float(page.get("kpi_height") or 2.10)
                y_kpi = y + 0.15
                L.kpi_row(cv, rows[0], y=y_kpi, h=kpi_h, cols=cols)
                takeaway_y = y_kpi + kpi_h + 0.28
                takeaway_h = min(1.65, sh - takeaway_y - 0.55)
                L.takeaway_card(cv, takeaway_text, takeaway_y, title=t_title, h=takeaway_h)
            else:
                kpi_h = float(page.get("kpi_height") or 2.40)
                y_kpi = y + 0.25
                L.kpi_row(cv, rows[0], y=y_kpi, h=kpi_h, cols=cols)
                default_insight = "指标综合分析：核心监测指标运行稳健，关键成效达成良好，建议持续保持优势模块并强化薄弱环节。"
                takeaway_y = y_kpi + kpi_h + 0.28
                takeaway_h = min(1.35, sh - takeaway_y - 0.55)
                L.takeaway_card(cv, default_insight, takeaway_y, title=t_title, h=takeaway_h)
        else:
            avail_h = (sh - 1.55 - y) if has_takeaway else (sh - 0.75 - y)
            kpi_h = float(page.get("kpi_height") or min(2.3, max(1.5, (avail_h / float(row_count)) - 0.22)))
            for r, chunk in enumerate(rows):
                L.kpi_row(cv, chunk, y=y + r * (kpi_h + 0.22), h=kpi_h, cols=cols)

            if has_takeaway:
                takeaway_y = y + row_count * (kpi_h + 0.22) + 0.15
                takeaway_h = min(1.40, sh - takeaway_y - 0.55)
                t_title = page.get("takeaway_title") or "核心指标深度洞察与分析"
                L.takeaway_card(cv, takeaway_text, takeaway_y, title=t_title, h=takeaway_h)

        if not has_takeaway and page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_table(self, idx, page):
        """卡片表格页：行高根据数据规模智能自适应，告别大片留白。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        headers = page.get("headers") or page.get("columns") or page.get("cols")
        raw_rows = page.get("rows") or page.get("data") or page.get("table") or []
        rows = list(raw_rows)
        if headers and isinstance(headers, list):
            if not rows or rows[0] != headers:
                rows = [headers] + rows
        n_rows = len(rows)

        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")

        if n_rows <= 3:
            row_h = 0.75 if not has_takeaway else 0.65
            header_size = 13
            size = 12
        elif n_rows <= 5:
            row_h = 0.62 if not has_takeaway else 0.54
            header_size = 12
            size = 11.5
        elif n_rows <= 8:
            row_h = 0.46
            header_size = 11.5
            size = 11
        else:
            row_h = page.get("row_h", 0.38)
            header_size = page.get("header_size", 10.5)
            size = page.get("size", 10)

        total_table_h = n_rows * row_h
        table_y = 1.25

        table_end_y = L.table(cv, L.MARGIN, table_y, L.CONTENT_W, rows,
                              col_widths=page.get("col_widths"),
                              row_h=row_h, size=size, header_size=header_size)

        t_title = page.get("takeaway_title") or "核心数据分析与洞察"
        if has_takeaway:
            takeaway_y = min(sh - 1.45, table_end_y + 0.25)
            takeaway_h = min(1.20, sh - takeaway_y - 0.50)
            L.takeaway_card(cv, takeaway_text, takeaway_y, title=t_title, h=takeaway_h)
        elif table_end_y < 3.8 and not page.get("footer"):
            default_insight = "数据洞察：从各项关键指标来看，整体分布均衡可控，核心数据表现符合预期，建议结合实际应用场景进一步开展针对性分析。"
            L.takeaway_card(cv, default_insight, table_end_y + 0.28, title=t_title, h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_timeline(self, idx, page):
        """里程碑与发展时间线页：大气流程卡片。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        items = (page.get("items") or page.get("steps") or
                 page.get("milestones") or page.get("stages") or
                 page.get("events") or [])
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        card_h = 2.80 if has_takeaway else 3.80
        end_y = L.timeline(cv, 1.45, items, h=card_h)

        if has_takeaway:
            takeaway_y = end_y + 0.25
            takeaway_h = min(1.20, sh - takeaway_y - 0.50)
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="推进保障机制", h=takeaway_h)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_text(self, idx, page):
        """正文演讲/说明页：双层卡片排版。"""
        cv = self._new_canvas()
        pal = cv.pal
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        card_y = 1.30
        card_h = sh - card_y - 0.65
        cv.card(L.MARGIN, card_y, L.CONTENT_W, card_h, radius=0.08)

        # 顶部精致装饰条
        cv.pill(L.MARGIN + 0.30, card_y + 0.04, 1.40, 0.035, fill=pal.get("accent", pal["primary"]))

        lead_text = page.get("lead") or page.get("summary") or page.get("quote")
        content_y = card_y + 0.35

        if lead_text:
            # 渲染导言引言微卡片
            lead_h = 0.72
            cv.card(L.MARGIN + 0.35, content_y, L.CONTENT_W - 0.70, lead_h,
                    fill=pal.get("card_subtle", pal.get("badge_bg")),
                    line=pal.get("card_border"), radius=0.06)
            cv.rect(L.MARGIN + 0.35, content_y, 0.06, lead_h, fill=pal.get("accent", pal["primary"]))
            lead_tf = cv.tbox(L.MARGIN + 0.55, content_y + 0.12, L.CONTENT_W - 1.10, lead_h - 0.24)
            cv.para(lead_tf, rich(str(lead_text), color=pal.get("accent")),
                    size=12.5, bold=True, color=pal.get("text_title", pal["primary"]),
                    line=1.25, first=True)
            content_y += lead_h + 0.25

        # 提取正文段落
        paras = (page.get("paragraphs") or page.get("body") or
                 page.get("items") or page.get("text") or page.get("content") or [])
        if isinstance(paras, str):
            paras = [p.strip() for p in paras.split("\n\n") if p.strip()] or [paras]

        remain_h = max(1.2, (card_y + card_h) - content_y - 0.30)
        tf = cv.tbox(L.MARGIN + 0.40, content_y, L.CONTENT_W - 0.80, remain_h)

        p_size = 12.0 if len(paras) >= 3 else 13.0
        p_gap = 10 if len(paras) >= 3 else 16

        for i, p_val in enumerate(paras):
            text_str = str(p_val)
            cv.para(tf, rich(text_str, color=pal.get("accent")), size=p_size,
                    color=pal.get("text_body", pal["text"]), line=1.45,
                    before=(0 if i == 0 else p_gap), first=(i == 0))

        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_stack(self, idx, page):
        """分层架构图页：自适应纵向技术栈/组织架构分层。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        layers = page.get("layers") or page.get("items") or page.get("steps") or []
        y = 1.30
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        stack_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.85)

        L.architecture_stack(cv, y, layers, h=stack_h)

        if has_takeaway:
            takeaway_y = y + stack_h + 0.22
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="架构设计原则", h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_hero_statement(self, idx, page):
        """核心战略论断/震撼大字报页：大字号直击核心结论，打破常规卡片沉闷。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        stmt = page.get("statement") or page.get("quote") or page.get("lead") or page.get("title") or ""
        pts = page.get("points") or page.get("items") or page.get("pillars") or []
        y = 1.30
        h = sh - y - 0.85
        L.hero_statement(cv, y, stmt, points=pts, h=h)

        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_process_flow(self, idx, page):
        """有向流程流转图页：带清晰粗箭头指示的有向步骤演进。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        steps = page.get("steps") or page.get("items") or page.get("stages") or []
        y = 1.35
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        flow_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.85)

        L.process_flow(cv, y, steps, h=flow_h)

        if has_takeaway:
            takeaway_y = y + flow_h + 0.22
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="执行关键控制点", h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_poetry(self, idx, page):
        """东方水墨诗词与韵律美学页：词牌名出处、对仗诗句与战略寓意解读。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        poem = page.get("poem") or page.get("lines") or page.get("statement") or page.get("content") or ""
        y = 1.30
        h = sh - y - 0.85
        L.poetry_card(cv, y, poem, title=page.get("tag") or page.get("cipai"),
                      author=page.get("author") or page.get("dynasty"),
                      interpretation=page.get("interpretation") or page.get("desc"),
                      h=h)

        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_diagram(self, idx, page):
        """分布式架构拓扑图页：服务节点方阵 + 调用流向箭头。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        nodes = page.get("nodes") or page.get("items") or []
        edges = page.get("edges") or page.get("links") or []
        y = 1.30
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        diag_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.85)

        L.topology_diagram(cv, y, nodes, edges=edges, h=diag_h)

        if has_takeaway:
            takeaway_y = y + diag_h + 0.22
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="拓扑调用与解耦机制", h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_team(self, idx, page):
        """团队与组织架构介绍页：核心专家/组织成员画像卡片。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        members = page.get("members") or page.get("team") or page.get("items") or []
        y = 1.30
        h = sh - y - 0.70
        L.team_grid(cv, y, members, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_pyramid(self, idx, page):
        """金字塔层级 / 转化漏斗递进页：多层居中递进渐变色块。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        tiers = page.get("tiers") or page.get("layers") or page.get("stages") or page.get("items") or []
        y = 1.30
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        pyr_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.75)
        end_y = L.pyramid_funnel(cv, y, tiers, h=pyr_h)
        if has_takeaway:
            L.takeaway_card(cv, takeaway_text, end_y + 0.20, title="层级递进与转化洞察", h=1.0)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])
        self._add_page_tracker(cv, idx)
        return cv

    def page_swot(self, idx, page):
        """SWOT 战略态势分析矩阵页：优势、劣势、机会、威胁 4 大象限。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        swot_data = page.get("swot") or page
        y = 1.30
        h = sh - y - 0.70
        L.swot_matrix(cv, y, swot_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_roadmap(self, idx, page):
        """战略演进路线图：3~4 个阶段跨周期里程碑与任务清单。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        phases = page.get("phases") or page.get("stages") or page.get("milestones") or page.get("items") or []
        y = 1.30
        h = sh - y - 0.70
        L.roadmap_milestones(cv, y, phases, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_pros_cons(self, idx, page):
        """方案利弊得失权衡页：优势收益 vs 风险成本权衡卡。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        pros = page.get("pros") or page.get("advantages") or page.get("gains") or []
        cons = page.get("cons") or page.get("disadvantages") or page.get("risks") or []
        takeaway = page.get("takeaway") or page.get("decision") or page.get("note")
        y = 1.30
        h = sh - y - 0.70
        L.pros_cons(cv, y, pros, cons, takeaway=takeaway, h=h)
        if page.get("footer"):
            L.footer_note(cv, sh, page["footer"])
        self._add_page_tracker(cv, idx)
        return cv

    def page_faq(self, idx, page):
        """常见疑虑与核心问答页：3~4 组优雅手风琴式 Q&A 卡片。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        faqs = page.get("faqs") or page.get("questions") or page.get("items") or []
        y = 1.30
        h = sh - y - 0.70
        L.faq_accordion(cv, y, faqs, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_case_study(self, idx, page):
        """标杆客户案例故事页：痛点挑战 + 落地措施 + ROI 量化成果 + 证言。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        case_data = page.get("case") or page
        y = 1.30
        h = sh - y - 0.70
        L.case_study(cv, y, case_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_pricing(self, idx, page):
        """商业定价与产品套餐页：3 档卡片与高光推荐卡。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        packages = page.get("packages") or page.get("plans") or page.get("tiers") or page.get("items") or []
        y = 1.30
        h = sh - y - 0.70
        L.pricing_packages(cv, y, packages, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_summary(self, idx, page):
        """高管总结与行动计划页：战略核心结论 + 下一步落地行动方案。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        summary_data = page.get("summary") or page
        y = 1.30
        h = sh - y - 0.70
        L.summary_next_steps(cv, y, summary_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_matrix_quadrant(self, idx, page):
        """二维四象限矩阵分析页：战略研报/投资收益/价值难度矩阵。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        matrix_data = page.get("matrix") or page
        y = 1.30
        h = sh - y - 0.70
        L.matrix_quadrant(cv, y, matrix_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_org_tree(self, idx, page):
        """组织架构与决策树：中枢决策 + 分支业务 + 底层支撑。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        tree_data = page.get("tree") or page
        y = 1.30
        h = sh - y - 0.70
        L.org_tree(cv, y, tree_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_radar(self, idx, page):
        """能力雷达与竞品全景评估页：多维双轨打分 + 综合评级卡。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        radar_data = page.get("radar") or page.get("capability") or page
        y = 1.30
        h = sh - y - 0.70
        L.capability_radar(cv, y, radar_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_funnel(self, idx, page):
        """转化漏斗专版：逐级下沉转化率阶梯。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        funnel_data = page.get("funnel") or page
        y = 1.30
        h = sh - y - 0.70
        L.funnel_stages(cv, y, funnel_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_chat_flow(self, idx, page):
        """AI 对话流与用户访谈原声：气泡流对话。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        chat_data = page.get("chat") or page.get("dialogue") or page
        y = 1.30
        h = sh - y - 0.70
        L.chat_dialogue(cv, y, chat_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_split_showcase(self, idx, page):
        """左右分屏产品/技术展示页：代码视窗 + 3 个 Bento 特性砖块。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        showcase_data = page.get("showcase") or page
        y = 1.30
        h = sh - y - 0.70
        L.split_showcase(cv, y, showcase_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_metric_grid(self, idx, page):
        """高密度微指标监控大盘：6~8 个紧凑指标卡。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        metrics_data = page.get("metrics") or page
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")

        y = 1.30
        h = (sh - y - 1.70) if has_takeaway else (sh - y - 0.70)
        grid_end_y = L.metric_grid(cv, y, metrics_data, h=h)

        if has_takeaway:
            takeaway_y = max(y + h + 0.20, sh - 1.40)
            takeaway_h = min(1.10, sh - takeaway_y - 0.45)
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="监控分析结论", h=takeaway_h)
        elif page.get("footer"):
            L.footer_note(cv, sh, page.get("footer"))

        self._add_page_tracker(cv, idx)
        return cv

    def page_quote_focus(self, idx, page):
        """沉浸式高管观点与大引用页：大号双引号 + 人物独白。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        quote_data = page.get("quote_card") or page
        y = 1.30
        h = sh - y - 0.70
        L.quote_focus(cv, y, quote_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_chart_analytics(self, idx, page):
        """数据洞察与原生图表分析页：左侧高保真大图表（折线/面积/环形/柱状） + 右侧核心洞察卡。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        analytics_data = page.get("analytics") or page
        y = 1.30
        h = sh - y - 0.70
        L.chart_analytics(cv, y, analytics_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_code_walkthrough(self, idx, page):
        """程序员源码深度剖析与算法走读页：代码高亮视窗 + 时序调用 + 复杂度 + 避坑指南。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        code_data = page.get("walkthrough") or page.get("code_data") or page
        y = 1.30
        h = sh - y - 0.70
        L.code_walkthrough(cv, y, code_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    def page_classroom_quiz(self, idx, page):
        """大中小学课堂教学互动测验与例题解析页：题目卡片 + 4 个选项卡片 + 名师思路点拨。"""
        cv = self._new_canvas()
        sw, sh = self.sw, self.sh
        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))
        quiz_data = page.get("quiz") or page.get("exercise") or page
        y = 1.30
        h = sh - y - 0.70
        L.classroom_quiz(cv, y, quiz_data, h=h)
        if page.get("footer") or page.get("note"):
            L.footer_note(cv, sh, page.get("footer") or page.get("note"))
        self._add_page_tracker(cv, idx)
        return cv

    LAYOUTS = {
        # 封面多元矩阵（多元场景适配）
        "title_slide": page_title_slide,
        "cover": page_title_slide,
        "title": page_title_slide,
        "cover_split_card": page_title_slide,
        "cover_centered": page_cover_centered,
        "title_centered": page_cover_centered,
        "cover_keynote": page_cover_centered,
        "cover_summit": page_cover_centered,
        "cover_pillars": page_cover_pillars,
        "title_pillars": page_cover_pillars,
        "cover_cards": page_cover_pillars,
        "cover_bottom_cards": page_cover_pillars,
        "cover_minimal": page_cover_minimal,
        "title_minimal": page_cover_minimal,
        "cover_editorial": page_cover_minimal,
        "cover_split": page_cover_split,
        "title_split": page_cover_split,
        "cover_hero": page_cover_split,
        "cover_contrast": page_cover_split,
        # 全局大目录导航大盘页（全篇架构导航）
        "agenda": page_agenda,
        "toc": page_agenda,
        "table_of_contents": page_agenda,
        "contents": page_agenda,
        "catalog": page_agenda,
        # 章节与过渡页多元矩阵
        "section": page_section,
        "chapter": page_section,
        "divider": page_section,
        "section_centered": page_section_centered,
        "chapter_centered": page_section_centered,
        "section_minimal": page_section_centered,
        "section_split": page_section_split,
        "chapter_split": page_section_split,
        "section_side": page_section_split,
        "section_progress": page_section_progress,
        "chapter_progress": page_section_progress,
        "section_timeline": page_section_progress,
        "chapter_stepper": page_section_progress,
        "dashboard": page_dashboard,
        "bullets": page_bullets,
        "grid": page_grid,
        "cards": page_grid,
        "comparison": page_comparison,
        "kpi": page_kpi,
        "table": page_table,
        "timeline": page_timeline,
        "text": page_text,
        "stack": page_stack,
        "architecture": page_stack,
        "layers": page_stack,
        "statement": page_hero_statement,
        "hero_statement": page_hero_statement,
        "quote": page_hero_statement,
        "focus": page_hero_statement,
        "process_flow": page_process_flow,
        "flow": page_process_flow,
        "pipeline": page_process_flow,
        "poetry": page_poetry,
        "poem": page_poetry,
        "verse": page_poetry,
        "diagram": page_diagram,
        "topology": page_diagram,
        "architecture_diagram": page_diagram,
        "team": page_team,
        "team_grid": page_team,
        "members": page_team,
        "org": page_team,
        "pyramid": page_pyramid,
        "pyramid_funnel": page_pyramid,
        "swot": page_swot,
        "swot_matrix": page_swot,
        "roadmap": page_roadmap,
        "roadmap_milestones": page_roadmap,
        "milestones": page_roadmap,
        "pros_cons": page_pros_cons,
        "tradeoff": page_pros_cons,
        "faq": page_faq,
        "faq_accordion": page_faq,
        "qa": page_faq,
        "case_study": page_case_study,
        "case": page_case_study,
        "pricing": page_pricing,
        "pricing_packages": page_pricing,
        "packages": page_pricing,
        "plans": page_pricing,
        "summary": page_summary,
        "summary_next_steps": page_summary,
        "next_steps": page_summary,
        "action_plan": page_summary,
        # 高质感矩阵、组织与评估版式
        "matrix_quadrant": page_matrix_quadrant,
        "matrix": page_matrix_quadrant,
        "quadrant": page_matrix_quadrant,
        "org_tree": page_org_tree,
        "tree": page_org_tree,
        "radar": page_radar,
        "capability": page_radar,
        "capability_radar": page_radar,
        "funnel": page_funnel,
        "funnel_stages": page_funnel,
        "chat_flow": page_chat_flow,
        "chat": page_chat_flow,
        "dialogue": page_chat_flow,
        "split_showcase": page_split_showcase,
        "showcase": page_split_showcase,
        "metric_grid": page_metric_grid,
        "metrics_grid": page_metric_grid,
        "quote_focus": page_quote_focus,
        "executive_quote": page_quote_focus,
        # 原生图表深度分析
        "chart_analytics": page_chart_analytics,
        "chart": page_chart_analytics,
        "analytics": page_chart_analytics,
        "data_analytics": page_chart_analytics,
        # 程序员技术深度走读
        "code_walkthrough": page_code_walkthrough,
        "code_review": page_code_walkthrough,
        "tech_talk": page_code_walkthrough,
        "tech_sharing": page_code_walkthrough,
        # 大中小学课堂教学与随堂测验
        "classroom_quiz": page_classroom_quiz,
        "quiz": page_classroom_quiz,
        "exam": page_classroom_quiz,
        "exercise": page_classroom_quiz,
    }


    # ---------------- 整体构建与体检 ----------------
    def render_pages(self, pages):
        self.total_pages = len(pages)
        self._prev_layout = None
        for idx, page in enumerate(pages):
            self._current_page = page
            # 单页独立主题切换：若页面显式指定 theme，则针对该页切至该主题
            page_theme = (page or {}).get("theme")
            base_pal = T.get_palette(page_theme) if page_theme else self._default_pal
            self.pal = base_pal

            layout = (page or {}).get("layout", "bullets")
            fn = self.LAYOUTS.get(layout)
            if fn is None:
                raise ValueError("未知 layout: %r（可选 %s）"
                                 % (layout, ", ".join(sorted(self.LAYOUTS))))
            fn(self, idx, page)
            self._validate_page(idx, page)
        self._current_page = None
        self.pal = self._default_pal
        if not pages:
            self._new_canvas()

    def _validate_page(self, idx, page):
        """轻量体检与大模型设计建议。"""
        layout = (page or {}).get("layout")
        prev_layout = getattr(self, "_prev_layout", None)
        if prev_layout and prev_layout == layout and layout in ("grid", "bullets", "cards"):
            self._warn(idx, "连续两页使用 %r 版式，易造成视觉疲劳；建议交替使用 stack 分层架构、comparison 对比或 hero_statement 核心论断增加节奏感" % layout)
        self._prev_layout = layout

        if layout == "dashboard":
            n = len(page.get("panels") or [])
            if n > 4:
                self._warn(idx, "panel 数量 %d 超过 4，建议拆页" % n)
        if layout == "bullets":
            cols = page.get("columns") or []
            if cols and len(cols) == 2:
                items_cnt = sum(len(c.get("items", [])) if isinstance(c, dict) else len(c) for c in cols)
                if items_cnt < 6:
                    self._warn(idx, "2 栏要点较少（共 %d 条），建议扩充说明或使用 grid 布局" % items_cnt)

    def save(self, path):
        """保存 presentation 至指定文件路径。"""
        out_dir = os.path.dirname(os.path.abspath(path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        self.prs.save(path)
        return path


def build_from_spec(spec):
    """按 spec 生成高精美度 pptx，返回 (deck, saved_path, warnings)。"""
    spec = spec or {}
    deck = Deck(theme=spec.get("theme", "gov_blue"),
                size=spec.get("size", "16:9"),
                font_cjk=spec.get("font_cjk"),
                font_latin=spec.get("font_latin"),
                meta=spec.get("meta"))
    pages = spec.get("pages") or []
    deck.render_pages(pages)
    out = spec.get("output") or {}
    path = out.get("path")
    if not path:
        name = out.get("name") or "deck.pptx"
        if not name.lower().endswith(".pptx"):
            name += ".pptx"
        path = os.path.join(out.get("dir") or os.getcwd(), name)
    deck.save(path)
    return deck, path, deck.warnings
