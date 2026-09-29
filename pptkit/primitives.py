# -*- coding: utf-8 -*-
"""现代化绘图原语系统：rect / rrect / card / badge / tbox / para / bullet / icon_box。

全面升级为符合现代高管汇报与科技发布会标准的视觉原语，
自带东亚中文字体保护（a:ea 标签）与高对比度排版引擎。
"""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}


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
            s.fill.fore_color.rgb = fill

        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = line
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

        # 估算宽度：汉字约按 0.16 字符宽，英文字母 0.09
        char_count = sum(2 if ord(c) > 127 else 1 for c in clean_text)
        w = max(0.9, char_count * 0.082 + (0.34 if actual_dot else 0.20) + pad_x * 2)

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
        """添加文本框并清除冗余边距。"""
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        m = Inches(margins) if margins else 0
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
        tf.vertical_anchor = anchor
        return tf

    def para(self, tf, runs, size=11, color=None, bold=False, align="left",
             before=0, after=0, line=1.2, first=False):
        """段落构建器，支持多 run 混排与中文字体强制声明。"""
        def_color = color or self.pal.get("text_body", self.pal["text"])
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.alignment = ALIGN.get(align, PP_ALIGN.LEFT)
        p.space_before = Pt(before)
        p.space_after = Pt(after)
        p.line_spacing = line

        if isinstance(runs, str):
            runs = [(runs, {})]

        for text, o in runs:
            r = p.add_run()
            r.text = text
            r.font.name = o.get("font", self.font_latin)
            r.font.size = Pt(o.get("sz", size))
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
