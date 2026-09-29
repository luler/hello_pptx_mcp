# -*- coding: utf-8 -*-
"""现代化绘图原语系统：rect / rrect / card / badge / tbox / para / bullet / icon_box。

全面升级为符合现代高管汇报与科技发布会标准的视觉原语，
自带东亚中文字体保护（a:ea 标签）与高对比度排版引擎。
"""
import math
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor

ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}


def to_rgb(col):
    """自动将 RGBColor 或 16 进制颜色字符串（如 '#FF5F56'）归一化为 RGBColor。"""
    if col is None:
        return None
    if isinstance(col, RGBColor):
        return col
    if isinstance(col, str):
        c = col.lstrip("#").strip()
        if len(c) == 6:
            try:
                return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))
            except ValueError:
                pass
    return col


class Canvas:
    """幻灯片矢量绘制上下文（单位统一为英寸）。"""

    def __init__(self, slide, pal, font_cjk="微软雅黑", font_latin="Arial"):
        self.slide = slide
        self.pal = pal
        self.font_cjk = font_cjk
        self.font_latin = font_latin

    # ---------------- 基础几何形状 ----------------
    def rect(self, x, y, w, h, fill=None, line=None, lw=0.75,
             shape=MSO_SHAPE.RECTANGLE, radius=None):
        """绘制矩形或圆角矩形。"""
        s = self.slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        s.shadow.inherit = False
        if fill is None:
            s.fill.background()
        else:
            s.fill.solid()
            s.fill.fore_color.rgb = to_rgb(fill)

        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = to_rgb(line)
            s.line.width = Pt(lw)

        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            try:
                # adjustments[0] 为圆角相对较短边的比例 (0.0~0.5)
                s.adjustments[0] = radius
            except Exception:
                pass
        s.text_frame.word_wrap = True
        return s

    def rrect(self, x, y, w, h, fill=None, line=None, lw=0.75, radius=0.06):
        """绘制细腻微圆角矩形。"""
        return self.rect(x, y, w, h, fill=fill, line=line, lw=lw,
                         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)

    def pill(self, x, y, w, h, fill=None, line=None, lw=0.5):
        """绘制胶囊状药丸形状（全圆角）。"""
        return self.rect(x, y, w, h, fill=fill, line=line, lw=lw,
                         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)

    def circle(self, x, y, r, fill=None, line=None, lw=0.75):
        """以 (x, y) 为左上角外接矩形绘制正圆。"""
        return self.rect(x, y, r, r, fill=fill, line=line, lw=lw,
                         shape=MSO_SHAPE.OVAL)

    def card(self, x, y, w, h, fill=None, line=None, lw=0.75, radius=0.06):
        """绘制现代浮岛卡片容器，默认取当前主题的卡片背景色与精细边框。"""
        f = fill or self.pal.get("card_bg")
        l = line or self.pal.get("card_border")
        return self.rrect(x, y, w, h, fill=f, line=l, lw=lw, radius=radius)

    def badge(self, x, y, text, dot=True, dot_color=None, bg=None,
              text_color=None, size=9.5, h=0.32, pad_x=0.18):
        """绘制高质感胶囊标签（如 '● 核心指标汇报'）。"""
        bg_col = bg or self.pal.get("badge_bg", self.pal["card_bg"])
        tx_col = text_color or self.pal.get("badge_text", self.pal["primary"])
        dot_col = dot_color or self.pal.get("accent", self.pal["primary"])

        # 智能规避：如果文本自身以特殊前缀符号开头，则自动不重复添加圆点
        clean_text = str(text).strip()
        has_symbol_prefix = any(clean_text.startswith(sym) for sym in ["●", "★", "✔", "▲", "▪", "▸", "◆", "■"])
        actual_dot = dot and not has_symbol_prefix

        # 智能紧凑宽度：数字或极短胶囊自适应紧凑尺寸，常规胶囊保持丰满大气
        char_count = sum(2 if ord(c) > 127 else 1 for c in clean_text)
        if char_count <= 2 and not actual_dot:
            effective_pad = min(pad_x, 0.10)
            min_w = 0.44
        elif char_count <= 4 and not actual_dot:
            effective_pad = min(pad_x, 0.14)
            min_w = 0.60
        else:
            effective_pad = pad_x
            min_w = 0.85
        w = max(min_w, char_count * 0.082 + (0.34 if actual_dot else 0.14) + effective_pad * 2)

        # 深色模式下为徽章增加柔和微光细边框，提升在暗底卡片上的悬浮质感与辨识度
        line_col = self.pal.get("card_border") if self.pal.get("is_dark") else None
        self.pill(x, y, w, h, fill=bg_col, line=line_col)
        tf = self.tbox(x, y, w, h, anchor=MSO_ANCHOR.MIDDLE)

        runs = []
        if actual_dot:
            runs.append(("●  ", {"c": dot_col, "b": True, "sz": size - 1.5}))
        runs.append((clean_text, {"c": tx_col, "b": True, "sz": size}))

        self.para(tf, runs, first=True, align="center")
        return x + w

    def icon_box(self, x, y, size=0.34, char="▸", fill=None, color=None):
        """在卡片标题前绘制小巧精致的图标容器或标号方块。"""
        bg_col = fill or self.pal.get("badge_bg", self.pal["card_subtle"])
        tx_col = color or self.pal.get("accent", self.pal["primary"])
        self.rrect(x, y, size, size, fill=bg_col, line=None, radius=0.25)
        tf = self.tbox(x, y, size, size, anchor=MSO_ANCHOR.MIDDLE)
        self.para(tf, [(str(char), {"c": tx_col, "b": True, "sz": 11.5})],
                  first=True, align="center")

    def divider(self, x, y, w, color=None, h=0.015):
        """绘制纤细典雅的水平分割线。"""
        c = color or self.pal.get("divider", self.pal["card_border"])
        return self.rect(x, y, w, h, fill=c, line=None)

    # ---------------- 文本排版系统 ----------------
    def tbox(self, x, y, w, h, anchor=MSO_ANCHOR.TOP, margins=0):
        """添加文本框并清除冗余边距，记录尺寸边界并注入原生自动缩放防护。"""
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        try:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
        m = Inches(margins) if margins else 0
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
        tf.vertical_anchor = anchor
        tf._box_w = float(w)
        tf._box_h = float(h)
        tf._box_x = float(x)
        tf._box_y = float(y)
        tf._est_used_h = 0.0
        return tf

    def para(self, tf, runs, size=11, color=None, bold=False, align="left",
             before=0, after=0, line=1.2, first=False, fit=True):
        """段落构建器，支持多 run 混排、中文字体强制声明与字多/字少双向自适应排版引擎。"""
        def_color = color or self.pal.get("text_body", self.pal["text"])
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.alignment = ALIGN.get(align, PP_ALIGN.LEFT)

        if first:
            tf._est_used_h = 0.0

        if isinstance(runs, str):
            runs = [(runs, {})]

        # ---------------- 双向自适应字号与防溢出算法 ----------------
        box_w = getattr(tf, "_box_w", None)
        box_h = getattr(tf, "_box_h", None)
        used_h = getattr(tf, "_est_used_h", 0.0)

        eff_size = float(size)
        eff_line = float(line)
        eff_before = float(before)
        eff_after = float(after)

        if fit and box_w and box_h and box_w > 0.3 and box_h > 0.2:
            avail_h = max(0.1, box_h - used_h)
            full_text = "".join(str(t) for t, _ in runs)
            # 计算等效中文字符数（CJK 占 1.0 字符位，ASCII/英文占 0.52 字符位）
            eff_chars = sum(1.0 if ord(c) > 127 else 0.52 for c in full_text)

            if eff_chars > 0:
                # 1. 字少自适应饱满提号（Anti-Hollow Boost）：当单段落文字较短且文本框高大时
                if first and box_h >= 1.1 and avail_h >= 0.9 and eff_chars <= 45 and size <= 11.5:
                    cpl_base = max(1.0, (box_w * 72.0) / (size * 1.05))
                    lines_base = math.ceil(eff_chars / cpl_base)
                    h_base = (lines_base * size * line + before + after) / 72.0
                    if h_base < avail_h * 0.45:
                        boost = min(2.0, max(0.5, (avail_h - h_base) * 1.5))
                        eff_size = min(14.0, size + boost)
                        eff_line = min(1.42, line + 0.12)

                # 2. 字多自适应降号防御（Anti-Overflow Downscaling）
                cpl = max(1.0, (box_w * 72.0) / (eff_size * 1.05))
                lines = math.ceil(eff_chars / cpl)
                h_needed = (lines * eff_size * eff_line + eff_before + eff_after) / 72.0

                if h_needed > avail_h:
                    fitted = False
                    for try_s in [eff_size - 1.0, eff_size - 1.8, eff_size - 2.5, eff_size - 3.2, 8.5, 7.5]:
                        if try_s < 7.5:
                            break
                        try_l = max(1.10, min(eff_line, 1.16))
                        try_cpl = max(1.0, (box_w * 72.0) / (try_s * 1.05))
                        try_lines = math.ceil(eff_chars / try_cpl)
                        try_h = (try_lines * try_s * try_l + eff_before * 0.6 + eff_after * 0.6) / 72.0
                        if try_h <= avail_h * 1.03:
                            eff_size = try_s
                            eff_line = try_l
                            eff_before *= 0.6
                            eff_after *= 0.6
                            fitted = True
                            break

                    # 3. 极度拥挤熔断截断（防止溢出穿透破坏整体卡片/遮挡其他模块）
                    if not fitted:
                        eff_size = 7.5
                        eff_line = 1.10
                        max_lines = max(1, int((avail_h * 72.0) / (7.5 * 1.10)))
                        max_cpl = (box_w * 72.0) / (7.5 * 1.05)
                        max_eff = max(6.0, max_lines * max_cpl * 0.90)

                        cur_eff = 0.0
                        new_runs = []
                        truncated = False
                        for t_str, o_dict in runs:
                            t_s = str(t_str)
                            if truncated:
                                break
                            r_eff = sum(1.0 if ord(c) > 127 else 0.52 for c in t_s)
                            if cur_eff + r_eff <= max_eff:
                                new_runs.append((t_s, o_dict))
                                cur_eff += r_eff
                            else:
                                remain_eff = max(0, max_eff - cur_eff)
                                sub_chars = []
                                sub_count = 0.0
                                for ch in t_s:
                                    ch_w = 1.0 if ord(ch) > 127 else 0.52
                                    if sub_count + ch_w <= remain_eff:
                                        sub_chars.append(ch)
                                        sub_count += ch_w
                                    else:
                                        break
                                cut_str = "".join(sub_chars).rstrip() + "..."
                                new_runs.append((cut_str, o_dict))
                                truncated = True
                                break
                        if new_runs:
                            runs = new_runs

        scale = eff_size / float(size) if size else 1.0
        p.space_before = Pt(eff_before)
        p.space_after = Pt(eff_after)
        p.line_spacing = eff_line

        # 估算实际占用高度
        cpl_final = max(1.0, (box_w * 72.0) / (eff_size * 1.05)) if box_w else 40.0
        full_t = "".join(str(t) for t, _ in runs)
        eff_c = sum(1.0 if ord(c) > 127 else 0.52 for c in full_t)
        lines_final = math.ceil(eff_c / cpl_final) if eff_c > 0 else 1
        tf._est_used_h = used_h + (lines_final * eff_size * eff_line + eff_before + eff_after) / 72.0

        for text, o in runs:
            r = p.add_run()
            r.text = str(text)
            r.font.name = o.get("font", self.font_latin)
            run_sz = o.get("sz", size) * scale
            r.font.size = Pt(max(7.5, run_sz))
            run_color = o.get("c")
            r.font.bold = o.get("b", bold)
            r.font.color.rgb = def_color if run_color is None else run_color

            # 强制注入东亚文字字体属性，杜绝 PPT 回退成宋体或乱码
            rpr = r.font._rPr
            rpr.set("altLang", "zh-CN")
            rpr.set("lang", "zh-CN")
            rpr.append(rpr.makeelement(qn("a:ea"),
                                       {"typeface": o.get("ea", self.font_cjk)}))
        return p

    def bullet(self, tf, runs, first=False, size=11, before=8,
               line=1.25, dot=None, mark="•  "):
        """现代精致项目符号，采用微调圆点。"""
        dot_color = dot or self.pal.get("accent", self.pal["primary"])
        prefix = [(mark, {"c": dot_color, "b": True, "sz": size + 1})]
        content = list(runs) if isinstance(runs, (list, tuple)) else [(runs, {})]
        return self.para(tf, prefix + content, size=size, before=before,
                         line=line, first=first)

    def text(self, x, y, w, h, s, size=11, color=None, bold=False, align="left",
             anchor=MSO_ANCHOR.TOP, line=1.1):
        """在指定矩形区域快速输出文本。"""
        tf = self.tbox(x, y, w, h, anchor=anchor)
        self.para(tf, s, size=size, color=color, bold=bold,
                  align=align, line=line, first=True)
        return tf

    def picture(self, path, x, y, w=None, h=None):
        """插入图片。"""
        kw = {}
        if w is not None:
            kw["width"] = Inches(w)
        if h is not None:
            kw["height"] = Inches(h)
        return self.slide.shapes.add_picture(path, Inches(x), Inches(y), **kw)
