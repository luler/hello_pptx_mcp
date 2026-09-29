# -*- coding: utf-8 -*-
"""开箱即用的 spec 模板。

每个模板都是完整可用的 spec，可直接 create_presentation(spec) 出片。
`get_template` MCP 工具会读取这里。
"""

WORK_REPORT = {
    "theme": "gov_blue",
    "size": "16:9",
    "meta": {"title": "研究生工作成效", "author": "研究生院",
             "subject": "招生·培养·学位全链条管理"},
    "pages": [
        {
            "layout": "title_slide",
            "title": "研究生工作成效汇报",
            "subtitle": "招生 · 培养 · 学位  全链条过程管理",
            "meta_line": "汇报单位：研究生院    汇报时间：2025 年",
            "bullets": ["生源质量逆势提升", "培养关键指标全面达标", "学位授予质量获省级认可"],
        },
        {
            "layout": "dashboard",
            "title": "研究生工作成效",
            "subtitle": "招生 · 培养 · 学位   全链条过程管理",
            "lead": "坚持分类培养、分类施策，建立全链条过程管理 —— [[生源质量逆势提升]]，"
                    "培养与学位关键指标全面达标，人才培养质量再获省级认可",
            "kpis": [
                {"label": "招  生", "value": 19, "unit": "人",
                 "captions": ["一志愿硕士高分生源", "（390分以上）2 → 4 → 19 人"]},
                {"label": "培  养", "value": 100, "unit": "%",
                 "captions": ["临床专硕住培首考通过率", "继去年理论首考100%再获突破"]},
                {"label": "科  研", "value": 21, "unit": "篇",
                 "captions": ["毕业年级发表 SCI 论文", "占发表文章总数 36.1%"]},
                {"label": "学  位", "value": 100, "unit": "%",
                 "captions": ["学位论文盲审通过率（93人）", "获省优秀学位论文 2 篇"]},
            ],
            "panels": [
                {"title": "招生工作",
                 "subtitle": "报考人数总体下降、竞争激烈，生源质量逆势提升",
                 "items": ["建立招生指标综合评价分配机制，优化招生布局",
                           "分类考核、精准选才"],
                 "chart": {"height": 1.05, "categories": ["2024级", "2025级", "2026级"],
                           "series": [{"name": "390分以上", "values": [2, 4, 19]},
                                      {"name": "400分以上", "values": [0, 2, 7]}]}},
                {"title": "培养工作",
                 "subtitle": "坚持分类培养、分类施策，强化全过程管理",
                 "items": [[("临床专硕住培首考通过率 ", {}), ("100%", {"b": True, "c": None}), ("，继去年理论首考100%再获突破", {})],
                           [("毕业年级发表 ", {}), ("SCI 论文 21 篇", {"b": True, "c": None}), ("，占发表文章总数 36.1%", {})],
                           [("专硕：", {"b": True}), ("住培并轨、集中备考，从临床实践中提炼科研课题", {})],
                           [("学硕：", {"b": True}), ("月度检查实验记录、开展学术沙龙、自主创新立项 29 项", {})]]},
                {"title": "学位工作",
                 "subtitle": "严把学位授予质量关，以“严”守住质量底线",
                 "items": [[("上半年 ", {}), ("93 名", {"b": True}), ("研究生学位论文盲审通过率 ", {}), ("100%", {"b": True})],
                           [("斩获河南省优秀学位论文 ", {}), ("2 篇", {"b": True}), ("（优秀博士、优秀硕士论文各1篇）", {})],
                           [("规范前置：", {"b": True}), ("学术规范、科研伦理、论文写作指导贯穿培养前中期", {})],
                           [("严格审核：", {"b": True}), ("对盲审通过但成绩低于80分的论文再次质量复核", {})]]},
            ],
            "footer": "数据来源：研究生院统计年报",
        },
    ],
}

PRODUCT_LAUNCH = {
    "theme": "tech_dark",
    "size": "16:9",
    "meta": {"title": "产品发布", "author": "产品团队"},
    "pages": [
        {"layout": "title_slide", "title": "数据平台 V3.0 发布",
         "subtitle": "更快 · 更稳 · 更开放",
         "meta_line": "2025 年度产品发布会", "bullets": []},
        {"layout": "kpi", "title": "核心指标", "subtitle": "V3.0 相较 V2.0",
         "cols": 3, "kpi_height": 1.7,
         "kpis": [
             {"label": "响应速度", "value": "3.2", "unit": "x", "captions": ["P95 延迟下降 68%"]},
             {"label": "可用性", "value": "99.99", "unit": "%", "captions": ["全年无重大故障"]},
             {"label": "开放接口", "value": 128, "unit": "个", "captions": ["较上版新增 54 个"]},
         ],
         "note": "指标口径：生产环境近 90 天滚动数据。"},
        {"layout": "table", "title": "版本对比",
         "rows": [["能力项", "V2.0", "V3.0", "提升"],
                  ["并发连接", "5,000", "20,000", "4x"],
                  ["冷启动耗时", "8.5s", "2.1s", "-75%"],
                  ["开放接口", "74", "128", "+54"],
                  ["SLA", "99.9%", "99.99%", "+0.09pp"]],
         "col_widths": [3, 2, 2, 2], "row_h": 0.42,
         "note": "测试环境：8C16G x 3 节点集群。"},
    ],
}

TRAINING_PLAN = {
    "theme": "violet",
    "size": "16:9",
    "meta": {"title": "培训方案", "author": "人力资源部"},
    "pages": [
        {"layout": "title_slide", "title": "新员工培训方案",
         "subtitle": "30 天融入计划", "meta_line": "人力资源部 · 2025"},
        {"layout": "timeline", "title": "培训节奏", "subtitle": "四个阶段，逐步融入",
         "items": [
             {"title": "第 1 周", "desc": "公司介绍、制度宣讲、环境熟悉"},
             {"title": "第 2 周", "desc": "业务通识、产品架构、工具链上手"},
             {"title": "第 3 周", "desc": "岗位实操、导师带教、结对开发"},
             {"title": "第 4 周", "desc": "结业考核、课题汇报、转正评估"},
         ],
         "note": "全程配备一对一导师，每周提交学习周报。"},
        {"layout": "bullets", "title": "考核标准",
         "columns": [
             {"title": "过程", "items": ["出勤率不低于 95%", "周报提交及时完整", "导师评价合格及以上"]},
             {"title": "结果", "items": ["结业笔试 ≥ 80 分", "课题汇报通过评审", "岗位实操无重大失误"]},
         ]},
    ],
}

TEMPLATES = {
    "work_report": {"label": "工作成效汇报（政务蓝 · 单页看板）", "spec": WORK_REPORT},
    "product_launch": {"label": "产品发布（深空蓝 · 指标+对比表）", "spec": PRODUCT_LAUNCH},
    "training_plan": {"label": "培训方案（紫罗兰 · 时间线+双栏）", "spec": TRAINING_PLAN},
}


def get(name):
    if name not in TEMPLATES:
        raise KeyError("未知模板 %r，可选: %s" % (name, ", ".join(TEMPLATES)))
    return TEMPLATES[name]["spec"]


def listing():
    return [{"name": k, "label": v["label"],
             "pages": len(v["spec"]["pages"]), "theme": v["spec"]["theme"]}
            for k, v in TEMPLATES.items()]
