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

_TAG_RE = re.compile(r"\[\[(.*?)\]\]")


def rich(text, color=None, size=None, bold=None):
    """将 '普通[[强调]]文本' 解析为 Canvas.para 混排 runs。"""
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
        runs.append((m.group(1), opt))
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
        self.pal = T.get_palette(theme)
        self.font_cjk = font_cjk or T.FONT_CJK_DEFAULT
        self.font_latin = font_latin or T.FONT_LATIN_DEFAULT
        self.sw, self.sh = T.SLIDE_SIZES.get(size, T.SLIDE_SIZES["16:9"])
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.sw)
        self.prs.slide_height = Inches(self.sh)
        self.warnings = []
        self.total_pages = 1
        self.meta = meta or {}
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

    def _new_canvas(self, bg=True, bg_color=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        fill_color = bg_color or self.pal.get("bg")
        if bg and fill_color:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = T.rgb(fill_color)
        return Canvas(slide, self.pal, self.font_cjk, self.font_latin)

    def _warn(self, page_idx, msg):
        self.warnings.append({"page": page_idx, "message": msg})

    def _add_page_tracker(self, cv, idx):
        """右下角轻量页码胶囊。"""
        if self.total_pages > 1 and idx > 0:
            text = "%02d / %02d" % (idx + 1, self.total_pages)
            px = self.sw - L.MARGIN - 1.05
            py = self.sh - 0.44
            cv.badge(px, py, text, dot=False,
                     bg=self.pal.get("badge_bg"),
                     text_color=self.pal.get("badge_text", self.pal.get("text_subtitle")),
                     size=8.5, h=0.25, pad_x=0.10)

    # ---------------- 核心版式渲染器 ----------------

    def page_title_slide(self, idx, page):
        """现代化非对称高管封面页：左右黄金分割，完美兼顾要点卡片与空态自适应。"""
        cv = self._new_canvas()
        pal = self.pal
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
        cat = page.get("category") or self.meta.get("subject") or "战略工作汇报"
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
            meta_line = "汇报部门：%s" % self.meta.get("author")
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
                cv.badge(right_x + 0.45, by + (sub_card_h - 0.28) / 2.0, num_str,
                         dot=False, bg=pal.get("badge_bg"), text_color=item_acc,
                         size=9, h=0.28)

                cv.text(right_x + 1.15, by + 0.10, right_w - 1.55, sub_card_h - 0.20,
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

    def page_section(self, idx, page):
        """章节过渡页：大牌感章节标号 + 议题导航概览。"""
        cv = self._new_canvas()
        pal = self.pal
        sw, sh = self.sw, self.sh
        accents = pal.get("accents") or [pal.get("accent")]
        sec_acc = accents[idx % len(accents)]

        L.chrome(cv, sw, sh)

        # 居中大浮岛卡片
        cw, ch = 9.8, 4.0
        cx = (sw - cw) / 2.0
        cy = (sh - ch) / 2.0

        cv.card(cx, cy, cw, ch, radius=0.10)

        # 章节序号标号
        sec_num = "%02d" % (idx + 1)
        cv.badge(cx + 0.6, cy + 0.55, "CHAPTER %s" % sec_num, dot=True,
                 dot_color=sec_acc, text_color=sec_acc, size=10.5, h=0.32)

        # 章节主标题
        cv.text(cx + 0.6, cy + 1.15, cw - 1.2, 0.95, page.get("title", ""),
                size=33, bold=True, color=pal.get("text_title", pal["primary"]), line=1.1)

        # 章节副标题
        if page.get("subtitle"):
            cv.text(cx + 0.6, cy + 2.25, cw - 1.2, 0.6, page["subtitle"],
                    size=14, color=pal.get("text_subtitle", pal["text_light"]), line=1.3)

        cv.pill(cx + 0.6, cy + 3.10, 1.5, 0.05, fill=sec_acc)

        self._add_page_tracker(cv, idx)
        return cv

    def page_dashboard(self, idx, page):
        """综合数据看板页：导读摘要 + Bento KPI + 复合多面板。"""
        cv = self._new_canvas()
        pal = self.pal
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
                    if not ser and chart.get("values"):
                        ser = [{"name": chart.get("series_name", "能力"), "values": chart["values"]}]
                    L.panel_with_chart(
                        cv, px, y, pw, ph, p.get("title", ""), p.get("subtitle"),
                        p.get("items"),
                        categories=cats,
                        series=ser,
                        chart_h=float(chart.get("height", 1.35)),
                        accent_color=panel_acc)
                else:
                    bx, by, bw = L.panel(cv, px, y, pw, ph, p.get("title", ""),
                                         p.get("subtitle"), accent_color=panel_acc)
                    items = p.get("items") or p.get("bullets") or []
                    if items:
                        is_all_metrics = all(isinstance(it, dict) and "label" in it and "value" in it for it in items)
                        if is_all_metrics and len(items) <= 5:
                            from .charts import progress_bar
                            item_space_h = (ph - (by - y) - 0.20) / float(len(items))

                            # 严格提取数值并计算最大基准，确保色块长度与数值严格对应（数值越大色块越长）
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
                                    # 绝对性能/指标数值（如 80、120、200、409.6），以最大值为 100% 满条基准，色块长度严格对应数值大小！
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
                            L.bullet_list(cv, bx, by, bw, ph - (by - y) - 0.15, items,
                                          size=p.get("size", 10.8), dot_color=panel_acc)
                    if p.get("progress") is not None:
                        from .charts import progress_bar
                        prog_val = p["progress"]
                        if isinstance(prog_val, (int, float)):
                            prog_ratio = float(prog_val) if prog_val <= 1.0 else (float(prog_val) / 100.0 if prog_val <= 100 else 1.0)
                        else:
                            prog_ratio = 0.5
                        prog_lbl = p.get("progress_label") or ("%.0f%%" % (prog_ratio * 100))
                        progress_bar(cv, bx, by + 0.25, bw, 0.16,
                                     prog_ratio, fill=panel_acc, label=prog_lbl)

        if page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_bullets(self, idx, page):
        """多栏结构化要点页：自适应砖块布局（Item Blocks），彻底消除大片空白。"""
        cv = self._new_canvas()
        pal = self.pal
        sw, sh = self.sw, self.sh
        accents = pal.get("accents") or [pal.get("accent", pal["primary"])]

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        cols = page.get("columns") or page.get("cols") or []
        start_y = 1.25

        if not cols:
            items = page.get("items") or page.get("bullets") or page.get("cards") or []
            n_items = len(items)
            card_h = min(5.4, max(3.0, 0.8 + n_items * 0.95))
            y_pos = start_y + max(0, ((sh - 0.8) - start_y - card_h) * 0.30)
            cv.card(L.MARGIN, y_pos, L.CONTENT_W, card_h)

            if n_items <= 5:
                # 采用精致独立砖块卡片
                block_w = L.CONTENT_W - 0.60
                item_h = (card_h - 0.60) / float(n_items)
                for i, it in enumerate(items):
                    it_acc = accents[i % len(accents)]
                    L.item_block(cv, L.MARGIN + 0.30, y_pos + 0.30 + i * item_h,
                                 block_w, item_h - 0.12, it, index=i + 1, size=11.5,
                                 accent_color=it_acc)
            else:
                L.bullet_list(cv, L.MARGIN + 0.35, y_pos + 0.30, L.CONTENT_W - 0.70,
                              card_h - 0.50, items, size=page.get("size", 11.5), gap=8)
        else:
            n = min(len(cols), 4)
            gap = 0.26
            cw = (L.CONTENT_W - gap * (n - 1)) / float(n)

            # 计算各列最大要点数
            max_items = max(len((c.get("items") or c.get("bullets") if isinstance(c, dict) else c) or []) for c in cols)
            card_h = min(5.4, max(3.4, 1.2 + max_items * 0.88))
            y_pos = start_y + max(0, ((sh - 0.8) - start_y - card_h) * 0.30)

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
                        for j, it in enumerate(c_items):
                            it_acc = accents[(i + j) % len(accents)]
                            L.item_block(cv, bx, by + j * it_h, bw, it_h - 0.10,
                                         it, index=j + 1, size=10.5, accent_color=it_acc)
                    else:
                        L.bullet_list(cv, bx, by, bw, card_h - (by - y_pos) - 0.20,
                                      c_items, size=c.get("size", 10.5), gap=6, dot_color=col_acc)
                else:
                    cv.card(cx, y_pos, cw, card_h)
                    cv.pill(cx + 0.25, y_pos + 0.04, min(1.2, cw - 0.50), 0.03, fill=col_acc)
                    L.bullet_list(cv, cx + 0.25, y_pos + 0.30, cw - 0.50, card_h - 0.50,
                                  c, size=page.get("size", 11), gap=8, dot_color=col_acc)

        if page.get("note") or page.get("footer"):
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

        card_h = page.get("card_height")
        if not card_h:
            card_h = 2.15 if (has_takeaway and len(cards) > 2) else (2.4 if len(cards) > 2 else 3.4)

        grid_end_y = L.grid_cards(cv, y, cards, cols=cols, h=card_h)

        if has_takeaway:
            takeaway_y = min(sh - 1.35, max(grid_end_y + 0.25, sh - 1.45))
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
        comp_h = (sh - y - 1.55) if has_takeaway else (sh - y - 0.85)

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
        pal = self.pal
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
            kpi_h = float(page.get("kpi_height") or 2.10)
            if has_takeaway:
                # 垂直黄金比例分布：KPI 居中偏上，下方紧跟充实的战略洞察大卡片
                y_kpi = y + 0.20
                L.kpi_row(cv, rows[0], y=y_kpi, h=kpi_h, cols=cols)
                takeaway_y = y_kpi + kpi_h + 0.35
                takeaway_h = min(1.65, sh - takeaway_y - 0.55)
                L.takeaway_card(cv, takeaway_text, takeaway_y, title="核心战略洞察与成果分析", h=takeaway_h)
            else:
                y_kpi = y + max(0.2, (sh - 1.0 - y - kpi_h) * 0.40)
                L.kpi_row(cv, rows[0], y=y_kpi, h=kpi_h, cols=cols)
        else:
            avail_h = (sh - 1.55 - y) if has_takeaway else (sh - 0.75 - y)
            kpi_h = float(page.get("kpi_height") or min(2.3, max(1.5, (avail_h / float(row_count)) - 0.22)))
            for r, chunk in enumerate(rows):
                L.kpi_row(cv, chunk, y=y + r * (kpi_h + 0.22), h=kpi_h, cols=cols)

            if has_takeaway:
                takeaway_y = y + row_count * (kpi_h + 0.22) + 0.15
                takeaway_h = min(1.40, sh - takeaway_y - 0.55)
                L.takeaway_card(cv, takeaway_text, takeaway_y, title="核心战略洞察与成果分析", h=takeaway_h)

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

        rows = page.get("rows") or page.get("data") or page.get("table") or []
        n_rows = len(rows)

        if n_rows <= 5:
            row_h = 0.54
            header_size = 12
            size = 11.5
        elif n_rows <= 8:
            row_h = 0.44
            header_size = 11.5
            size = 11
        else:
            row_h = page.get("row_h", 0.38)
            header_size = page.get("header_size", 10.5)
            size = page.get("size", 10)

        total_table_h = n_rows * row_h
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        table_y = 1.25

        table_end_y = L.table(cv, L.MARGIN, table_y, L.CONTENT_W, rows,
                              col_widths=page.get("col_widths"),
                              row_h=row_h, size=size, header_size=header_size)

        if has_takeaway:
            takeaway_y = max(table_end_y + 0.25, sh - 1.45)
            takeaway_h = min(1.20, sh - takeaway_y - 0.50)
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="数据洞察与结论", h=takeaway_h)
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
                 page.get("milestones") or page.get("stages") or [])
        has_takeaway = bool(page.get("takeaway") or page.get("note"))
        takeaway_text = page.get("takeaway") or page.get("note")
        card_h = 2.80 if has_takeaway else 3.40
        L.timeline(cv, 1.65, items, h=card_h)

        if has_takeaway:
            takeaway_y = 1.65 + card_h + 0.30
            takeaway_h = min(1.15, sh - takeaway_y - 0.50)
            L.takeaway_card(cv, takeaway_text, takeaway_y, title="推进保障机制", h=takeaway_h)
        elif page.get("footer"):
            L.footer_note(cv, sh, page["footer"])

        self._add_page_tracker(cv, idx)
        return cv

    def page_text(self, idx, page):
        """正文演讲/说明页：双层卡片排版。"""
        cv = self._new_canvas()
        pal = self.pal
        sw, sh = self.sw, self.sh

        L.chrome(cv, sw, sh)
        L.header(cv, page.get("title", ""), page.get("subtitle"))

        card_y = 1.30
        card_h = sh - card_y - 0.65
        cv.card(L.MARGIN, card_y, L.CONTENT_W, card_h)

        tf = cv.tbox(L.MARGIN + 0.40, card_y + 0.35, L.CONTENT_W - 0.80, card_h - 0.70)
        paras = page.get("body") or page.get("items") or []
        if isinstance(paras, str):
            paras = paras.split("\n\n")

        for i, para_text in enumerate(paras):
            cv.para(tf, rich(para_text, color=pal.get("accent")), size=12.5,
                    color=pal.get("text_body", pal["text"]), line=1.45,
                    before=(0 if i == 0 else 12), first=(i == 0))

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

    LAYOUTS = {
        "title_slide": page_title_slide,
        "section": page_section,
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
        "funnel": page_pyramid,
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
    }

    # ---------------- 整体构建与体检 ----------------
    def render_pages(self, pages):
        self.total_pages = len(pages)
        self._prev_layout = None
        for idx, page in enumerate(pages):
            layout = (page or {}).get("layout", "bullets")
            fn = self.LAYOUTS.get(layout)
            if fn is None:
                raise ValueError("未知 layout: %r（可选 %s）"
                                 % (layout, ", ".join(sorted(self.LAYOUTS))))
            fn(self, idx, page)
            self._validate_page(idx, page)
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
