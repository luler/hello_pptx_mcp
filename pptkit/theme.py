# -*- coding: utf-8 -*-
"""配色主题与版式设计系统。

支持多套精美高级、多场景自适应的设计主题（浅色商务、深色极客、自然生态、奢华黑金等），
全面解耦语义颜色，确保无论切换何种主题，文字、卡片、背景与高亮元素均具备极致对比度与视觉美感。
"""
from pptx.dml.color import RGBColor


def rgb(h):
    """将 '#1B4F9C' 或 '1B4F9C' 转换为 RGBColor 对象。"""
    if isinstance(h, RGBColor):
        return h
    h = str(h).lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


SLIDE_SIZES = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
}

# --------------------------------------------------------------------------
# 高级主题调色板（包含浅色商业、深色科技、政务庄重、活力互联网、极简北欧等）
# --------------------------------------------------------------------------
PALETTES = {
    # 1. 经典政务蓝（庄重规范、严谨高质感，适用于国企、政务、事业单位汇报）
    "gov_blue": {
        "label": "政务蓝 · 权威规范 / 国企党政汇报",
        "is_dark": False,
        "bg": "F4F7FC",                # 柔和高雅的浅蓝灰底色，彻底告别刺眼死白
        "bg_alt": "E8F0FA",            # 次级背景
        "card_bg": "FFFFFF",           # 纯白卡片，形成悬浮层次
        "card_subtle": "EEF4FB",       # 浅蓝灰弱强调卡片
        "card_border": "D0E0F2",       # 细腻柔和边框
        "primary": "124184",           # 权威深政务蓝
        "primary_mid": "1E5AA8",       # 中饱和度主色
        "primary_dark": "0B2853",      # 深沉暗蓝
        "primary_tint": "E2EEFC",      # 极浅主色背景
        "accent": "C52026",            # 经典中国红强调色（突出核心数据）
        "accent_tint": "FDE8E9",       # 浅红高光背景
        "accent2": "E68A00",           # 暖金橙辅助色
        "text_title": "0D264A",        # 极高对比度的深蓝黑标题
        "text_subtitle": "3B5270",     # 次级说明文本
        "text_body": "2C3B4E",         # 正文深灰蓝（比纯黑更护眼高雅）
        "text_muted": "667D99",        # 弱化注释
        "badge_bg": "E3EFFC",          # 徽章背景
        "badge_text": "124184",        # 徽章文字
        "divider": "D0E0F2",           # 分割线
        "table_header_bg": "124184",   # 表头深蓝
        "table_header_text": "FFFFFF", # 表头文字
        "table_alt_row": "F3F8FD",     # 斑马纹
        "chart_colors": ["124184", "C52026", "E68A00", "2E86DE", "10AC84"],
        "accents": ["124184", "C52026", "E68A00", "0D9488", "6D28D9", "2563EB"],
        # 兼容旧字段
        "text": "2C3B4E", "text_light": "667D99",
        "border": "D0E0F2", "bar_bg": "EEF4FB", "white": "FFFFFF",
    },

    # 2. 现代科技黑（深邃极客、高对比荧光青，专为科技发布会、AI创新方案打造）
    "tech_dark": {
        "label": "深空黑 · 极客科技 / AI 与产品发布",
        "is_dark": True,
        "bg": "0A0F1D",                # 深邃午夜极光黑
        "bg_alt": "10182C",            # 次级暗底
        "card_bg": "131F37",           # 半透质感磨砂蓝黑卡片
        "card_subtle": "182644",       # 弱强调卡片
        "card_border": "24385E",       # 优雅发光感细边框
        "primary": "38BDF8",           # 亮天青色（发光感）
        "primary_mid": "0284C7",
        "primary_dark": "0369A1",
        "primary_tint": "11284A",      # 极光暗蓝底
        "accent": "00F0FF",            # 赛博荧光青（核心数据吸睛爆发力）
        "accent_tint": "0E384D",
        "accent2": "FBBF24",           # 亮琥珀金橙（高对比度）
        "text_title": "FFFFFF",        # 纯白大标题，极致通透
        "text_subtitle": "E2E8F0",     # 亮冷银白次标题（高对比度，清晰醒目，告别发暗）
        "text_body": "F8FAFC",         # 纯净极白正文（对比度 15:1，极其清晰通透）
        "text_muted": "94A3B8",        # 优雅银灰（比原 64748B 大幅提升对比度，清晰易读）
        "badge_bg": "1E293B",          # 科技胶囊背景
        "badge_text": "38BDF8",        # 科技胶囊文本
        "divider": "203254",           # 微光分割线
        "table_header_bg": "1A2B4C",   # 深蓝紫表头
        "table_header_text": "00F0FF", # 发光青表头字
        "table_alt_row": "0E1729",     # 表格深色斑马纹
        "chart_colors": ["00F0FF", "38BDF8", "FBBF24", "C084FC", "34D399"],
        "accents": ["00F0FF", "C084FC", "FBBF24", "34D399", "38BDF8", "F472B6"],
        # 兼容旧字段
        "text": "F8FAFC", "text_light": "E2E8F0",
        "border": "24385E", "bar_bg": "182644", "white": "FFFFFF",
    },

    # 3. 商务咨询蓝（顶尖投行麦肯锡/波士顿咨询风格，极简理性精英感）
    "business_blue": {
        "label": "咨询蓝 · 顶尖投行 / 商业计划与财报",
        "is_dark": False,
        "bg": "F8FAFC",                # 现代极简冷灰底色
        "bg_alt": "EDF2F7",
        "card_bg": "FFFFFF",
        "card_subtle": "F1F5F9",
        "card_border": "E2E8F0",
        "primary": "0F2748",           # 藏青深海蓝
        "primary_mid": "1D4ED8",       # 皇家亮宝蓝
        "primary_dark": "09182E",
        "primary_tint": "EEF4FD",
        "accent": "2563EB",            # 饱和钴蓝
        "accent_tint": "E0EDFE",
        "accent2": "D97706",           # 优雅金棕
        "text_title": "0A101D",        # 极致对比度石板黑
        "text_subtitle": "2D3B4E",     # 中深石板灰
        "text_body": "1E293B",         # 深灰高清晰正文
        "text_muted": "475569",        # 清晰易读注释
        "badge_bg": "EFF6FF",
        "badge_text": "1D4ED8",
        "divider": "E2E8F0",
        "table_header_bg": "0F2748",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F8FAFC",
        "chart_colors": ["1D4ED8", "0F2748", "D97706", "3B82F6", "64748B"],
        "accents": ["1D4ED8", "0D9488", "D97706", "7C3AED", "BE185D", "059669"],
        # 兼容旧字段
        "text": "1E293B", "text_light": "475569",
        "border": "E2E8F0", "bar_bg": "F1F5F9", "white": "FFFFFF",
    },

    # 4. 黑金尊享（高端战略领袖、年度盛典、峰会路演）
    "executive_gold": {
        "label": "黑金尊享 · 高端战略 / 领袖峰会与盛典",
        "is_dark": True,
        "bg": "121215",                # 曜石哑光黑
        "bg_alt": "1A1A1F",
        "card_bg": "1D1D24",           # 深灰黑浮岛卡片
        "card_subtle": "25252E",
        "card_border": "3E382A",       # 隐约的古铜金色边线
        "primary": "E0BA53",           # 奢华香槟金
        "primary_mid": "F3CF72",
        "primary_dark": "9E7D23",
        "primary_tint": "2B2414",
        "accent": "E0BA53",            # 耀目真金数值
        "accent_tint": "362C16",
        "accent2": "F3CF72",
        "text_title": "FFFFFF",        # 纯白大标题，极致通透
        "text_subtitle": "EADFC9",     # 亮象牙米白次标题（告别昏暗）
        "text_body": "FBF8EF",         # 暖白高对比正文
        "text_muted": "C5BBA6",        # 清晰浅暖灰注释（大幅提升对比度）
        "badge_bg": "2B2414",
        "badge_text": "F3CF72",
        "divider": "3D3627",
        "table_header_bg": "2A2417",
        "table_header_text": "E0BA53",
        "table_alt_row": "16161B",
        "chart_colors": ["E0BA53", "F3CF72", "38BDF8", "FBBF24", "34D399"],
        "accents": ["E0BA53", "38BDF8", "FBBF24", "34D399", "C084FC", "F472B6"],
        # 兼容旧字段
        "text": "FBF8EF", "text_light": "EADFC9",
        "border": "3E382A", "bar_bg": "25252E", "white": "FFFFFF",
    },

    # 5. 清新薄荷绿（生命科学、环保双碳、教育与医疗健康）
    "fresh_green": {
        "label": "清新绿 · 医疗健康 / ESG 双碳与教育",
        "is_dark": False,
        "bg": "F4FAF6",                # 柔润薄荷雾白
        "bg_alt": "EAF5EE",
        "card_bg": "FFFFFF",
        "card_subtle": "EDF7F1",
        "card_border": "CCE6D6",
        "primary": "0F7642",           # 沉稳森林绿
        "primary_mid": "16A34A",       # 活力翠绿
        "primary_dark": "0B512D",
        "primary_tint": "E2F5EA",
        "accent": "059669",            # 翡翠绿核心强调
        "accent_tint": "D1FAE5",
        "accent2": "D97706",           # 暖琥珀橙（点缀活跃度）
        "text_title": "08361E",        # 极深墨绿黑大标题
        "text_subtitle": "1C4B33",
        "text_body": "142E20",         # 墨草深黑正文（高对比度）
        "text_muted": "3D6650",
        "badge_bg": "E1F5E9",
        "badge_text": "0F7642",
        "divider": "CCE6D6",
        "table_header_bg": "0F7642",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F2F9F5",
        "chart_colors": ["0F7642", "16A34A", "D97706", "3B82F6", "10B981"],
        "accents": ["059669", "0284C7", "D97706", "10B981", "7C3AED", "EC4899"],
        # 兼容旧字段
        "text": "142E20", "text_light": "3D6650",
        "border": "CCE6D6", "bar_bg": "EDF7F1", "white": "FFFFFF",
    },

    # 6. 紫罗兰幻境（品牌营销、文化创意、新消费路演）
    "violet": {
        "label": "紫罗兰 · 品牌营销 / 创意消费与先锋演说",
        "is_dark": False,
        "bg": "F8F7FD",                # 淡雅薰衣草白
        "bg_alt": "EFEBFB",
        "card_bg": "FFFFFF",
        "card_subtle": "F3EFFD",
        "card_border": "DED5F9",
        "primary": "53229E",           # 贵族深紫罗兰
        "primary_mid": "7C3AED",       # 梦幻电紫
        "primary_dark": "3B1277",
        "primary_tint": "EFEAFE",
        "accent": "D92688",            # 摩登洋红（时尚感十足）
        "accent_tint": "FDE8F4",
        "accent2": "E68A00",           # 金色提亮
        "text_title": "1A0836",        # 极深紫黑大标题
        "text_subtitle": "3D2466",
        "text_body": "22153B",         # 深紫黑清晰正文
        "text_muted": "534073",
        "badge_bg": "EFEBFC",
        "badge_text": "53229E",
        "divider": "DED5F9",
        "table_header_bg": "53229E",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F4F1FD",
        "chart_colors": ["53229E", "7C3AED", "D92688", "E68A00", "06B6D4"],
        "accents": ["7C3AED", "06B6D4", "D92688", "E68A00", "10B981", "3B82F6"],
        # 兼容旧字段
        "text": "22153B", "text_light": "534073",
        "border": "DED5F9", "bar_bg": "F3EFFD", "white": "FFFFFF",
    },

    # 7. 活力珊瑚橙（创业路演、互联网产品、敏捷运营）
    "warm_coral": {
        "label": "珊瑚橙 · 创业路演 / 互联网增长与运营",
        "is_dark": False,
        "bg": "FDF9F7",                # 极暖象牙白
        "bg_alt": "F9ECE5",
        "card_bg": "FFFFFF",
        "card_subtle": "FCF0EA",
        "card_border": "F6D5C4",
        "primary": "C2410C",           # 稳重红褐
        "primary_mid": "EA580C",       # 活力火红橙
        "primary_dark": "9A3412",
        "primary_tint": "FFEDE5",
        "accent": "F95738",            # 爆裂活力珊瑚色
        "accent_tint": "FFE9E4",
        "accent2": "2563EB",           # 互补对比宝蓝
        "text_title": "2E0A04",        # 暖浓黑棕大标题
        "text_subtitle": "5E271B",
        "text_body": "2B1611",         # 浓郁深褐黑正文
        "text_muted": "6B4339",
        "badge_bg": "FFECE5",
        "badge_text": "C2410C",
        "divider": "F6D5C4",
        "table_header_bg": "C2410C",
        "table_header_text": "FFFFFF",
        "table_alt_row": "FDF3ED",
        "chart_colors": ["F95738", "EA580C", "2563EB", "10B981", "6B7280"],
        "accents": ["F95738", "EA580C", "2563EB", "10B981", "06B6D4", "D946EF"],
        # 兼容旧字段
        "text": "2B1611", "text_light": "6B4339",
        "border": "F6D5C4", "bar_bg": "FCF0EA", "white": "FFFFFF",
    },

    # 8. 数字科技青（数字孪生、新零售、敏捷技术方案）
    "modern_teal": {
        "label": "科技青 · 数字化转型 / 前沿技术与物联网",
        "is_dark": False,
        "bg": "F5FAF9",                # 冰爽微青冷白
        "bg_alt": "E8F4F2",
        "card_bg": "FFFFFF",
        "card_subtle": "EEF7F5",
        "card_border": "C8E5E0",
        "primary": "0E7490",           # 沉静深青
        "primary_mid": "06B6D4",       # 跃动电青
        "primary_dark": "155E75",
        "primary_tint": "E0F6F9",
        "accent": "0891B2",            # 高饱和碧水蓝
        "accent_tint": "CFFAFE",
        "accent2": "F97316",           # 暖亮橙对比
        "text_title": "0B3B4B",        # 极深墨青黑
        "text_subtitle": "164E63",
        "text_body": "11303D",         # 深青黑清晰正文
        "text_muted": "3B6473",
        "badge_bg": "E0F4F2",
        "badge_text": "0E7490",
        "divider": "C8E5E0",
        "table_header_bg": "0E7490",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F0F7F6",
        "chart_colors": ["0E7490", "06B6D4", "F97316", "10B981", "8B5CF6"],
        "accents": ["0E7490", "06B6D4", "F97316", "10B981", "8B5CF6", "E11D48"],
        # 兼容旧字段
        "text": "11303D", "text_light": "3B6473",
        "border": "C8E5E0", "bar_bg": "EEF7F5", "white": "FFFFFF",
    },

    # 9. 东方水墨诗词（国风美学、传统文化、诗书典故、借诗言志）
    "chinese_poetry": {
        "label": "东方丹青 · 诗词水墨 / 传统国风美学与文化典雅",
        "is_dark": False,
        "bg": "F8F5EE",                # 宣纸暖玉素白，典雅温润
        "bg_alt": "EEE8DC",            # 澄浆古纸微暗底
        "card_bg": "FFFFFF",           # 纯白雅致浮岛卡片
        "card_subtle": "F3EFE6",       # 浅米宣纸微强调
        "card_border": "D8CFC0",       # 细腻古朴竹丝边线
        "primary": "2B2E33",           # 沉稳松烟水墨黑
        "primary_mid": "4A4E57",
        "primary_dark": "181A1D",
        "primary_tint": "EFE9DD",      # 仿古笺纸淡黄底
        "accent": "B9382E",            # 经典朱砂红（印章点睛与核心数据）
        "accent_tint": "F8E7E5",       # 浅朱印泥淡光
        "accent2": "3D665E",           # 宋代汝窑天青/黛绿
        "text_title": "111315",        # 极深水墨标题
        "text_subtitle": "33363D",     # 次级烟墨色
        "text_body": "181A1D",         # 正文纯墨黑
        "text_muted": "575B63",
        "badge_bg": "EFE9DD",          # 仿古笺纸胶囊底
        "badge_text": "B9382E",        # 朱砂字色
        "divider": "D8CFC0",
        "table_header_bg": "2B2E33",
        "table_header_text": "F8F5EE",
        "table_alt_row": "F7F3EA",
        "chart_colors": ["B9382E", "3D665E", "C89B3C", "355070", "6D597A"],
        "accents": ["B9382E", "3D665E", "C89B3C", "355070", "8A5A44", "495867"],
        # 兼容旧字段
        "text": "181A1D", "text_light": "575B63",
        "border": "D8CFC0", "bar_bg": "F3EFE6", "white": "FFFFFF",
    },
}

FONT_CJK_DEFAULT = "微软雅黑"
FONT_LATIN_DEFAULT = "Arial"
FONT_CJK_FALLBACK = "Noto Sans CJK SC"


def get_palette(name):
    """根据主题名称或自定义配色字典获取调色板，自动将 hex 色值解析为 RGBColor，并支持未知主题智能回退。"""
    aliases = {
        "business": "business_blue",
        "gold": "executive_gold",
        "dark": "tech_dark",
        "green": "fresh_green",
        "purple": "violet",
        "blue": "gov_blue",
        "coral": "warm_coral",
        "teal": "modern_teal",
        "poetry": "chinese_poetry",
        "ink": "chinese_poetry",
        "guofeng": "chinese_poetry",
        "chinese": "chinese_poetry",
    }

    if isinstance(name, dict):
        base_key = str(name.get("base", "gov_blue")).strip().lower()
        base_key = aliases.get(base_key, base_key)
        raw = PALETTES.get(base_key, PALETTES["gov_blue"]).copy()
        for k, v in name.items():
            if k != "base":
                raw[k] = v
        raw.setdefault("label", "大模型定制配色 · 自定义风格")
    else:
        key = str(name).strip().lower()
        if key not in PALETTES:
            key = aliases.get(key, "gov_blue")
        raw = PALETTES.get(key, PALETTES["gov_blue"])

    out = {}
    for k, v in raw.items():
        if k in ("label", "is_dark"):
            out[k] = v
        elif k in ("chart_colors", "accents"):
            out[k] = [rgb(c) for c in v]
        else:
            try:
                out[k] = rgb(v)
            except Exception:
                out[k] = v
    if "accents" not in out:
        out["accents"] = out.get("chart_colors") or [out["accent"]]
    return out

