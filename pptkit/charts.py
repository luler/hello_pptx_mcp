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
    n_cat = len(categories) if categories else 0
    if not series or n_cat == 0:
        return

    # 兼容处理：若 series 为字典或一维数字列表，自动归一化为列表
    if isinstance(series, dict):
        series = [series]
    elif isinstance(series, list) and len(series) > 0 and isinstance(series[0], (int, float)):
        series = [{"name": "数值", "values": series}]

    n_ser = len(series)
    if n_ser == 0:
        return

    flat = [float(v) for s in series if isinstance(s, dict) for v in s.get("values", [])]
    # 顶部预留 18% 缓冲区，防止柱顶数值标签与图例或面板标题冲突碰撞
    top = float(max_value or (max(flat) * 1.18 if flat else 1) or 1)
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


def line_chart(cv, x, y, w, h, categories, series, legend=None, is_area=False, data_labels=None):
    """原生平滑折线图与面积渐变图（支持多系列、坐标轴美化与主题调色）。"""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from .primitives import to_rgb

    pal = cv.pal
    colors = pal.get("chart_colors") or pal.get("accents") or [
        pal.get("primary", pal["primary"]),
        pal.get("accent2", pal["accent"]),
        pal.get("accent", pal["accent"]),
        pal.get("primary_mid", pal["primary"]),
    ]
    rgb_colors = [to_rgb(c) for c in colors]

    cd = CategoryChartData()
    cd.categories = [str(c) for c in categories]

    if isinstance(series, dict):
        series = [series]
    elif isinstance(series, list) and len(series) > 0 and isinstance(series[0], (int, float)):
        series = [{"name": "趋势", "values": series}]

    for s in series:
        if isinstance(s, dict):
            cd.add_series(str(s.get("name", "系列")), s.get("values", []))

    chart_type = XL_CHART_TYPE.AREA if is_area else XL_CHART_TYPE.LINE
    shape = cv.slide.shapes.add_chart(chart_type, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    chart = shape.chart
    chart.has_title = False

    text_muted_rgb = to_rgb(pal.get("text_muted", pal["text"]))
    text_body_rgb = to_rgb(pal.get("text_body", pal["text"]))

    # 轴刻度文本颜色与字体自适应主题（深色模式下自动转为通透白/冷灰，杜绝发黑）
    try:
        chart.category_axis.tick_labels.font.color.rgb = text_muted_rgb
        chart.category_axis.tick_labels.font.size = Pt(8.5)
        chart.value_axis.tick_labels.font.color.rgb = text_muted_rgb
        chart.value_axis.tick_labels.font.size = Pt(8.5)
    except Exception:
        pass

    show_legend = (len(series) > 1) if legend is None else bool(legend)
    chart.has_legend = show_legend
    if show_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(8.5)
        try:
            chart.legend.font.color.rgb = text_muted_rgb
        except Exception:
            pass

    # 面积图默认不打数据标签以免重叠，折线图根据参数或系列规模自适应
    if data_labels is None:
        show_labels = (not is_area) and (len(categories) <= 6) and (len(series) <= 2)
    else:
        show_labels = bool(data_labels)

    chart.plots[0].has_data_labels = show_labels
    if show_labels:
        try:
            chart.plots[0].data_labels.font.size = Pt(8.5)
            chart.plots[0].data_labels.font.color.rgb = text_body_rgb
        except Exception:
            pass

    for i, s_obj in enumerate(chart.series):
        col = rgb_colors[i % len(rgb_colors)]
        try:
            s_obj.format.line.color.rgb = col
            s_obj.format.line.width = Pt(2.5)
            if is_area:
                s_obj.format.fill.solid()
                s_obj.format.fill.fore_color.rgb = col
        except Exception:
            pass


def pie_chart(cv, x, y, w, h, categories, series_or_values, legend=True, is_doughnut=False, data_labels=None):
    """原生环形图（Donut）与饼图（Pie），自适应主题色调与数据标签。"""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from .primitives import to_rgb

    pal = cv.pal
    colors = pal.get("chart_colors") or pal.get("accents") or [
        pal.get("primary", pal["primary"]),
        pal.get("accent2", pal["accent"]),
        pal.get("accent", pal["accent"]),
        pal.get("primary_mid", pal["primary"]),
    ]
    rgb_colors = [to_rgb(c) for c in colors]

    cd = CategoryChartData()
    cd.categories = [str(c) for c in categories]

    if isinstance(series_or_values, dict):
        vals = series_or_values.get("values", [])
        s_name = series_or_values.get("name", "占比")
    elif isinstance(series_or_values, list) and len(series_or_values) > 0:
        if isinstance(series_or_values[0], dict):
            vals = series_or_values[0].get("values", [])
            s_name = series_or_values[0].get("name", "占比")
        else:
            vals = series_or_values
            s_name = "占比"
    else:
        vals = []
        s_name = "占比"

    cd.add_series(s_name, vals)

    chart_type = XL_CHART_TYPE.DOUGHNUT if is_doughnut else XL_CHART_TYPE.PIE
    shape = cv.slide.shapes.add_chart(chart_type, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    chart = shape.chart
    chart.has_title = False

    text_muted_rgb = to_rgb(pal.get("text_muted", pal["text"]))
    text_body_rgb = to_rgb(pal.get("text_body", pal["text"]))

    chart.has_legend = bool(legend)
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(8.5)
        try:
            chart.legend.font.color.rgb = text_muted_rgb
        except Exception:
            pass

    show_labels = True if data_labels is None else bool(data_labels)
    chart.plots[0].has_data_labels = show_labels
    if show_labels:
        try:
            from pptx.dml.color import RGBColor
            chart.plots[0].data_labels.font.size = Pt(9.0)
            chart.plots[0].data_labels.font.bold = True
            chart.plots[0].data_labels.font.color.rgb = RGBColor(255, 255, 255)
        except Exception:
            pass

    try:
        for i, pt in enumerate(chart.plots[0].series[0].points):
            col = rgb_colors[i % len(rgb_colors)]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = col
    except Exception:
        pass


def render_chart(cv, x, y, w, h, chart_spec):
    """统一图表分发引擎：根据 spec.type 分发至柱状图、折线图、面积图、饼图或环形图。"""
    if not isinstance(chart_spec, dict):
        return
    ctype = str(chart_spec.get("type", "bar")).lower()
    cats = chart_spec.get("categories") or chart_spec.get("labels") or []
    series = chart_spec.get("series") or chart_spec.get("values") or []
    legend = chart_spec.get("legend")
    data_labels = chart_spec.get("data_labels")

    if ctype in ("line", "trend"):
        line_chart(cv, x, y, w, h, cats, series, legend=legend, is_area=False, data_labels=data_labels)
    elif ctype in ("area", "area_chart"):
        line_chart(cv, x, y, w, h, cats, series, legend=legend, is_area=True, data_labels=data_labels)
    elif ctype in ("pie", "pie_chart"):
        pie_chart(cv, x, y, w, h, cats, series, legend=True if legend is None else bool(legend), is_doughnut=False, data_labels=data_labels)
    elif ctype in ("doughnut", "donut", "ring"):
        pie_chart(cv, x, y, w, h, cats, series, legend=True if legend is None else bool(legend), is_doughnut=True, data_labels=data_labels)
    else:
        bar_chart(cv, x, y, w, h, cats, series, legend=legend)


