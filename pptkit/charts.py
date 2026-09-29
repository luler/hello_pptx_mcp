# -*- coding: utf-8 -*-
"""现代化图表构件：轻量级矢量柱状图与高质感进度条。

原生基于 PowerPoint 矢量形状拼装，体积极小、排版自适应主题色彩、
在各类屏幕与投屏器上均保持绝对锐利。
"""
from pptx.enum.text import MSO_ANCHOR


def bar_chart(cv, x, y, w, h, categories, series, max_value=None,
              legend=None, value_labels=True):
    """现代化精致分组柱状图。

    categories: ["2024级", "2025级", "2026级"]
    series: [{"name": "390分以上", "values": [2, 4, 19]},
             {"name": "400分以上", "values": [0, 2, 7]}]
    """
    pal = cv.pal
    colors = pal.get("chart_colors") or pal.get("accents") or [
        pal.get("primary", pal["primary"]),
        pal.get("accent2", pal["accent"]),
        pal.get("accent", pal["accent"]),
        pal.get("primary_mid", pal["primary"]),
    ]
    n_cat, n_ser = len(categories), len(series)
    if n_cat == 0 or n_ser == 0:
        return

    flat = [float(v) for s in series for v in s.get("values", [])]
    top = float(max_value or (max(flat) if flat else 1) or 1)
    if top <= 0:
        top = 1.0

    # 单系列图表默认不展示多余的通用图例
    show_legend = (n_ser > 1) if legend is None else bool(legend)
    legend_h = 0.28 if show_legend else 0.0
    label_h = 0.24
    plot_h = max(h - legend_h - label_h - 0.15, 0.45)
    base = y + legend_h + plot_h

    # 1. 图例（自适应现代胶囊排版）
    if show_legend:
        lx = x
        for i, s in enumerate(series):
            c_fill = colors[i % len(colors)]
            cv.circle(lx, y + 0.05, 0.09, fill=c_fill)
            name_str = str(s.get("name", ""))
            cv.text(lx + 0.14, y + 0.01, 1.6, 0.22, name_str,
                    size=9, color=pal.get("text_muted", pal["text"]), line=1.0)
            char_len = sum(2 if ord(ch) > 127 else 1 for ch in name_str)
            lx += max(1.2, char_len * 0.08 + 0.35)

    # 2. 微光基准底线
    cv.rect(x, base, w, 0.015, fill=pal.get("divider", pal["card_border"]))

    gw = w / float(n_cat)
    # 计算柱宽与间距
    bar_w = min(0.42, (gw * 0.70) / float(n_ser))
    gap = min(0.06, bar_w * 0.25)
    pair_w = bar_w * n_ser + gap * (n_ser - 1)

    for ci, cat in enumerate(categories):
        gx = x + gw * ci + (gw - pair_w) / 2.0
        for si, s in enumerate(series):
            vals = s.get("values", [])
            val = float(vals[ci]) if ci < len(vals) else 0.0
            bx = gx + si * (bar_w + gap)
            # 单系列时柱子按分类分配多彩 accent，多系列时按系列区分色彩
            color = colors[ci % len(colors)] if n_ser == 1 else colors[si % len(colors)]

            if val > 0:
                bh = max(0.04, plot_h * (val / top))
                # 柱子带顶部细腻圆角
                cv.rrect(bx, base - bh, bar_w, bh, fill=color, radius=0.20)
                ty = base - bh - 0.20
            else:
                ty = base - 0.20

            if value_labels:
                if val >= 100000000:
                    display_val = ("%.1f亿" % (val / 100000000.0)).rstrip("0").rstrip(".")
                elif val >= 10000:
                    display_val = ("%.1f万" % (val / 10000.0)).rstrip("0").rstrip(".")
                elif val.is_integer():
                    display_val = str(int(val))
                else:
                    display_val = "%.1f" % val

                cv.text(bx - 0.16, ty, bar_w + 0.32, 0.18, display_val,
                        size=9.5, bold=True, color=color, align="center")

        # X 轴分类标签
        cv.text(x + gw * ci, base + 0.06, gw, label_h, str(cat),
                size=9.5, color=pal.get("text_muted", pal["text"]), align="center")


def progress_bar(cv, x, y, w, h, ratio, fill=None, bg=None, label=None,
                 label_size=9.5):
    """现代化胶囊进度条，右侧配有精美数据徽章。"""
    pal = cv.pal
    ratio = max(0.0, min(1.0, float(ratio)))
    has_label = bool(label)
    label_str = str(label) if has_label else ""
    # 动态根据数值文字字符数分配合适的右侧宽度，防止类似 409.6、100% 出现换行或拥挤
    label_w = max(0.75, len(label_str) * 0.095 + 0.18) if has_label else 0.0
    track_w = max(0.5, w - label_w)

    track_bg = bg or pal.get("card_subtle", pal["bar_bg"])
    fill_col = fill or pal.get("primary", pal["primary"])

    # 轨道底色
    cv.pill(x, y, track_w, h, fill=track_bg)

    # 填充进度条：长度严格与 ratio 线性对应（数值越大色块越长）
    if ratio > 0.001:
        fill_w = max(h * 0.7, track_w * ratio)
        cv.pill(x, y, fill_w, h, fill=fill_col)

    # 标签数值
    if has_label:
        cv.text(x + track_w + 0.08, y - 0.04, label_w, h + 0.08, label_str,
                size=label_size, bold=True, color=fill_col,
                anchor=MSO_ANCHOR.MIDDLE)

