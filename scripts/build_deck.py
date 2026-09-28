# -*- coding: utf-8 -*-
"""
示例 deck：一份 15 页的中文科研组会汇报
========================================

这份 deck 有三个作用：
  1. 当**页型样例**——十五页覆盖了 references/page-recipes.md 里的全部页型，
     改内容不要改结构，照着抄结构最省事；
  2. 当**回归测试**——`verify_all.py` 拿它跑四道闸门。改 kit / 母版 / 闸门之后
     先跑一遍这里，全绿说明没把基线搞坏；
  3. 当**纯排版路线的示范**——通篇无图，全部靠栅格、卡片、数据大数字和原生图表撑版面。

内容全是中性虚构的（XX 课题组 / 示例数据集），不指向任何真实课题，
数字也全部是编的——照抄时请换成真实口径，见 references/narration.md 的口径纪律。

用法：
    python scripts/build_deck.py --help              # 参数说明
    python scripts/build_deck.py                     # → out/示例汇报_中文科研组会.pptx
    python scripts/build_deck.py --zh-kaiti          # 中文改楷体，产物另存一份
    python scripts/build_deck.py --template D:/xx/我的模板.pptx
产物：
    out/示例汇报_中文科研组会.pptx
    out/示例汇报_中文科研组会_layout.json     版面清单（闸门一的输入）
    out/示例汇报_中文科研组会_讲稿.md          讲稿（与 PPT 内嵌备注同源）
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import kit
from kit import (BAND, BODY, BORDER, BROWN, INK, M, MUTED, SKY, TINT, WHITE, CW,
                 cap, chip, column_chart, fig_fit, foot, head, headbar, new_slide, notes,
                 panel, rect, rule, table, txt)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

MT = MSO_ANCHOR.MIDDLE
CT = PP_ALIGN.CENTER
L = PP_ALIGN.LEFT
RT = PP_ALIGN.RIGHT

TITLE = "示例汇报 · 三层检索架构的落地与评测"
EYEBROW = "组会汇报 · 示例课题组"
FOOT = "示例汇报 · 三层检索架构 · 内部草稿"
MEET = "第 12 周 · 内部场次"
PAGES_N = 15

# 三等分 / 四等分栅格：宽度与间距是算出来的，不是量出来的，改 CW 时自动跟着走
BW3, GAP3 = 3.90, 0.155
X3 = [M + i * (BW3 + GAP3) for i in range(3)]
BW4, GAP4 = 2.875, 0.17
X4 = [M + i * (BW4 + GAP4) for i in range(4)]


# ────────────────────────────── 页面 ──────────────────────────────
def cover(pn):
    """封面。左下几何装饰由版式自带，右侧白底排字，一行装饰都不加。"""
    s = new_slide("cover", pn)
    x0 = 6.95
    txt(s, x0, 1.50, 5.72, 0.30, [(EYEBROW, 13.5, True)], color=BAND, spc=4, anchor=MT)
    txt(s, x0, 1.92, 5.72, 0.80, [("Retrieval = Index + Graph", 30, True, INK)],
        latin=kit.FORM, anchor=MT, ls=1.20)
    txt(s, x0, 2.88, 5.72, 0.72, [("三层检索架构的落地与评测", 30, True)], color=INK, anchor=MT)
    rule(s, x0, 3.76, 2.60, color=kit.ACC, h=0.07)
    txt(s, x0, 3.96, 5.72, 0.66,
        [("把一套检索方法从论文搬进生产：本讲汇报实验设计、三组结果，以及还差哪三件事。", 15.5)],
        color=BODY, anchor=MT, ls=1.30)
    txt(s, x0, 4.74, 5.72, 0.62,
        [("示例数据集：1,337 条查询 × 6,551 条候选；本页所有数字均为演示用途的虚构值。", 12.5)],
        color=MUTED, anchor=MT, ls=1.32)
    rule(s, x0, 5.76, 5.72, color=BORDER)
    for i, (k, v) in enumerate([("研究模块", "检索与评测模块"),
                                ("对比对象", "关键词 · 语义 · 图谱 · 融合"),
                                ("组会时间", MEET)]):
        y = 5.92 + i * 0.36
        txt(s, x0, y, 2.20, 0.26, [(k, 10.5)], color=MUTED, spc=3, anchor=MT)
        txt(s, x0 + 2.30, y, 3.42, 0.26, [(v, 12.5, True)], color=INK, anchor=MT)


def roadmap(pn):
    """路线页：四个编号标签 + 一句话总纲 + 口径先说清。

    底部那条不是预告，是**口径的自我限制**（样本量多大、标注到什么程度）——
    先说在前面，免得结论被误读。写「下一步做 A、B、C」是正常交代，
    写「下一讲的结果是 X」才是自相矛盾。
    """
    s = new_slide("content", pn)
    head(s, "路线 · 四步", "四步走完：问题、实验、结果、还差什么")
    steps = [
        ("01", "问题", "把一个方法从论文搬进生产，先要回答的是：它到底比基线好在哪一档。"),
        ("02", "实验设计", "同一批查询、同一套指标，三种检索路径逐层叠加，每层贡献单独算清。"),
        ("03", "三组结果", "命中率、排序质量、端到端时延各给一组数，赢的照实写赢、输的照实写输。"),
        ("04", "还差什么", "只报已经算完的数，最后交代还差哪三件事。"),
    ]
    for i, (no, t, d) in enumerate(steps):
        x = M + i * (BW4 + GAP4)
        chip(s, x, 1.74, BW4, 0.62, no, t, size=16)
        txt(s, x, 2.50, BW4, 1.14, [(d, 13.5)], color=BODY, ls=1.34)
    rect(s, M, 3.72, CW, 1.58, fill=BAND, line=None)
    txt(s, M + 0.34, 3.72, CW - 0.68, 1.58,
        [("一句话总纲　", 15.5, True, SKY),
         ("检索这件事该分给确定性工程：", 15.5, False, WHITE),
         ("索引管召回，图谱管关系，模型只管把答案组织成话。", 15.5, True, WHITE)],
        anchor=MT, ls=1.36)
    rect(s, M, 5.48, CW, 1.06, fill=TINT, line=BORDER, lw=1.5)
    txt(s, M + 0.34, 5.48, CW - 0.68, 1.06,
        [("口径先说清　", 15, True, BAND),
         ("标注集只有 19 条，是两个人独立标完再对齐的结果，没有第三方终审；"
          "样本量小，所以这一讲只看相对排序、不看绝对值。定价与硬件成本都是假设区间。", 14.5)],
        color=BODY, anchor=MT, ls=1.34)
    foot(s, pn)


def thesis(pn):
    """总纲页：一句公式 + 两栏对照 + 底部结论带。"""
    s = new_slide("content", pn)
    head(s, "总纲", "检索栈拆成三层：各自最贵的那部分资源都不一样")
    txt(s, M, 1.72, CW, 0.70, [("Retrieval = 索引召回 + 图谱扩展 + 语义重排", 28, True)],
        color=BAND, latin=kit.FORM, anchor=MT)
    cols = [
        ("索引层", "确定性工程",
         ["结构化倒排，毫秒级返回", "可复现、可审计、每条能溯源",
          "短板：同义说法召回不到"]),
        ("图谱层", "关系工程",
         ["多跳关联，把词面不接的两端接上", "代价是入库与维护成本",
          "短板：冷启动实体对不齐"]),
    ]
    for i, (t, sub, lines) in enumerate(cols):
        x, w = M + i * 6.02, 5.99
        panel(s, x, 2.56, w, 2.62)
        txt(s, x + 0.34, 2.74, w - 0.68, 0.48, [(t, 24, True)], color=BAND, anchor=MT)
        txt(s, x + 0.34, 3.30, w - 0.68, 0.30, [(sub, 13.5)], color=MUTED, anchor=MT)
        rule(s, x + 0.34, 3.72, 1.10, color=kit.ACC)
        for j, ln in enumerate(lines):
            txt(s, x + 0.34, 3.88 + j * 0.42, w - 0.68, 0.40,
                [("·　" + ln, 14.5)], color=BODY, anchor=MT, ls=1.26)
    rect(s, M, 5.38, CW, 1.44, fill=BAND, line=None)
    txt(s, M + 0.34, 5.38, CW - 0.68, 1.44,
        [("这条结论是：", 15, True, SKY),
         ("大模型不是不能用，是不能放在召回位置上。", 19, True, WHITE),
         ("  召回要可复现、要能溯源，生成式模型给不了这个承诺。", 15, None, SKY)],
        anchor=MT, ls=1.38)
    foot(s, pn)


def compare(pn):
    """对照表页：左表右评，底部一句收口。表格必须有底板，否则版面会显得很空。"""
    s = new_slide("content", pn)
    head(s, "对照", "同一批查询，四条路径的同一把尺")
    hd = ["路径", "召回数", "MRR", "P95 时延", "可溯源"]
    rows = [
        ["仅倒排", "5.0", "0.318", "18 ms", "是"],
        ["倒排 + 语义", "8.0", "0.463", "96 ms", "是"],
        ["倒排 + 图谱", "8.0", "0.512", "24 ms", "是"],
        ["三条融合", "8.0", "0.677", "143 ms", "是"],
    ]
    tw = [2.10, 1.05, 1.05, 1.35, 1.05]
    panel(s, M, 1.74, sum(tw) + 0.44, 0.50 * (len(rows) + 1) + 0.44, fill=WHITE)
    body = []
    for i, r in enumerate(rows):
        last = (i == len(rows) - 1)
        body.append([(r[0], 13, last, BAND if last else INK, L),
                     (r[1], 13, last, BAND if last else BODY, CT),
                     (r[2], 13, last, BAND if last else BODY, CT),
                     (r[3], 13, False, BODY, CT),
                     (r[4], 13, False, BODY, CT)])
    table(s, M + 0.22, 1.96, tw, [list(hd)] + body, rowH=0.50)

    ins = [("融合赢在排序", "召回条数三条路径都是 8 条，差别全在顺序：MRR 从 0.463 抬到 0.677。"),
           ("图谱赢在召回", "倒排 + 图谱比纯倒排多召回 3 条，且 P95 只多 6 ms。"),
           ("代价写在时延", "融合路径 P95 到 143 ms，是纯倒排的 近 12 倍；这是实付的价钱。")]
    x, w = 9.18, 3.49
    for i, (t, d) in enumerate(ins):
        y = 1.74 + i * 1.52
        txt(s, x, y, w, 0.38, [(t, 16, True)], color=BAND if i == 0 else BORDER, anchor=MT)
        txt(s, x, y + 0.44, w, 0.92, [(d, 13.5)], color=BODY, ls=1.32)
        if i < 2:
            rule(s, x, y + 1.36, w)
    txt(s, M, 5.62, 8.30, 0.86,
        [("三条路径用的是同一批查询、同一份标注、同一套评测脚本，脚本一个字没改。", 13.5)],
        color=BODY, ls=1.34)
    rect(s, M, 6.10, CW, 0.72, fill=TINT, line=BORDER, lw=1.5)
    txt(s, M + 0.32, 6.10, CW - 0.64, 0.72,
        [("表里最该看的不是 0.677，是它比 0.463 高出来的那一段。", 15, True)], color=BAND, anchor=MT)
    foot(s, pn)


def datanums(pn):
    """数据大数字页：四格等宽卡。数字用 Times New Roman 粗体，视觉上更稳。"""
    s = new_slide("content", pn)
    head(s, "数据", "实验跑在什么数据、什么硬件上")
    st = [("1,337", "条", "查询集", "全部来自示例数据集，人工挑选并逐条标注相关文档"),
          ("6,551", "条", "候选池", "四条路径在同一份候选池上召回，不换数据、不挑样本"),
          ("19", "条", "标注集", "两人独立标注后对齐，分歧全部留痕，未经第三方终审"),
          ("6", "GB", "显存预算", "全程本地推理，单条 P95 143 ms，未触发降频")]
    for i, (n, u, t, d) in enumerate(st):
        x = X4[i]
        panel(s, x, 1.74, BW4, 2.46)
        txt(s, x + 0.26, 1.90, BW4 - 0.52, 1.00,
            [(n, 30, True, BAND), (" " + u, 15, True, MUTED)], anchor=MT)
        txt(s, x + 0.26, 2.98, BW4 - 0.52, 0.36, [(t, 15, True)], color=INK, anchor=MT)
        txt(s, x + 0.26, 3.40, BW4 - 0.52, 0.72, [(d, 12.5)], color=BODY, ls=1.32)
    headbar(s, M, 4.36, CW, "一个意外收获：把评测脚本顺手做成了可复现实验", h=0.46, size=16)
    txt(s, M, 4.96, 8.30, 1.30,
        [("原来的评测脚本靠手工改参数跑，换一组数据就要重写一遍。改成配置文件驱动之后，"
          "同一份脚本能跑全部配置，实验记录里能查到每一次的随机种子与环境版本。", 13.5)],
        color=BODY, ls=1.36)
    rect(s, 9.22, 4.96, 3.45, 1.30, fill=TINT, line=BORDER, lw=1.5)
    txt(s, 9.48, 4.96, 2.93, 1.30,
        [("这件事本身就该单独立项", 14.5, True, BAND),
         ("\n可复现不是加分项，是让别人的结论能被质疑的前提。", 12.5)],
        color=BODY, anchor=MT, ls=1.34)
    txt(s, M, 6.44, CW, 0.40,
        [("标注集 19 条是分母，不是查询条数；样本量小，后面所有数只看相对排序。", 15, True)],
        color=BORDER, anchor=MT)
    foot(s, pn)


def layers(pn):
    """分层页：左侧四层 + 右侧深色论断面板。"""
    s = new_slide("content", pn)
    head(s, "分层", "四层结构：大模型只进最上面一层", 28)
    layers = [
        ("4", "生成与交互", "本地小模型 / 云端大模型", "把候选组织成答案", "概率系统", True),
        ("3", "语义重排", "交叉编码器", "把顺序排对", "亚线性、毫秒级", False),
        ("2", "结构与关系", "知识图谱 / 关系库", "多跳关联、实体对齐", "确定性、可审计", False),
        ("1", "数据资产", "查询 / 文档 / 实体", "带溯源的原始数据", "单一事实来源", False),
    ]
    x0, w, rh = M, 8.35, 1.02
    for i, (no, t, tool, job, trait, top) in enumerate(layers):
        y = 1.74 + i * rh
        if top:
            rect(s, x0, y, w, rh - 0.16, fill=BAND, line=None)
        else:
            panel(s, x0, y, w, rh - 0.16)
        txt(s, x0 + 0.12, y, 0.60, rh - 0.16, [(no, 30, True)],
            color=SKY if top else BORDER, italic=True, latin=kit.NUMF, align=CT, anchor=MT)
        txt(s, x0 + 0.82, y, 1.55, rh - 0.16, [(t, 17, True)],
            color=WHITE if top else BAND, anchor=MT)
        txt(s, x0 + 2.45, y, 1.95, rh - 0.16, [(tool, 13.5)],
            color=SKY if top else BODY, anchor=MT, ls=1.26)
        txt(s, x0 + 4.48, y, 2.30, rh - 0.16, [(job, 13)],
            color=SKY if top else BODY, anchor=MT, ls=1.26)
        txt(s, x0 + 6.86, y, 1.49, rh - 0.16, [(trait, 12.5, True)],
            color=BORDER if not top else WHITE, align=RT, anchor=MT)
    kit.note(s, x0, 5.98, w, "每一层都有各自最便宜的资源，数据自下而上流动。", size=12.5, h=0.32)
    x1, w1 = 9.18, 3.49
    rect(s, x1, 1.74, w1, 4.98, fill=BAND, line=None)
    txt(s, x1 + 0.30, 1.98, w1 - 0.60, 0.86, [("只有第 4 层才留给大模型", 21, True)],
        color=WHITE, anchor=MT, ls=1.30)
    rule(s, x1 + 0.30, 2.94, 1.00, color=kit.ACC)
    txt(s, x1 + 0.30, 3.12, w1 - 0.60, 1.90,
        [("大模型不是不能用，是只能放在最上面一层。每一层干自己最划算的事，"
          "这条分界线要拿真实数据去验证，而不是靠说法。", 14)],
        color=SKY, ls=1.38)
    txt(s, x1 + 0.30, 5.22, w1 - 0.60, 1.34,
        [("下面三层", 15.5, True, WHITE),
         ("是这套系统最该自己做的部分。", 15.5, True, SKY)], ls=1.36)
    foot(s, pn)


def chart(pn):
    """图表页：左解释右原生柱状图，底部两个大数字卡。"""
    s = new_slide("content", pn)
    head(s, "发现", "图谱补的是字面上找不到的那一跳")
    txt(s, M, 1.76, 5.42, 0.46,
        [("查询原文只有「设备巡检」四个字，没有「传感器」。", 16, True)], color=BAND,
        anchor=MT, ls=1.30)
    panel(s, M, 2.36, 5.42, 1.30)
    txt(s, M + 0.28, 2.50, 4.86, 0.36,
        [("图谱把「设备巡检」与电力、电网连在一起", 14.5, True)], color=BORDER, anchor=MT)
    txt(s, M + 0.28, 2.94, 4.86, 0.58,
        [("一跳就是从这个查询词跨一次关系边到候选词，于是召回了「传感器状态监测」这条文档。",
          13.5)], color=BODY, anchor=MT, ls=1.32)
    txt(s, M, 3.84, 5.42, 1.10,
        [("纯倒排只能回到查询原文的词面，图谱存的是实体间的关系边，"
          "所以这一跳只能由图谱补上。", 15, True)], color=BORDER, anchor=MSO_ANCHOR.TOP, ls=1.36)
    column_chart(s, 6.26, 1.74, 6.41, 2.72,
                 ["纯倒排", "＋语义重排", "＋图谱扩展", "＋融合排序"],
                 (0.318, 0.463, 0.512, 0.677),
                 [kit.ACC, BORDER, BAND, INK], series_name="MRR", max_val=0.82)
    kit.note(s, 6.26, 4.56, 6.41,
             "同一候选池、同一份 19 条标注、同一套指标，仅切换「是否接入图谱扩展」。",
             size=11, h=0.34)
    cards = [("+3", "条召回", "图谱扩展单独贡献的召回数", BAND, True),
             ("+0.052", "MRR", "再加融合排序的额外贡献", TINT, False)]
    for i, (n, u, d, fill, dark) in enumerate(cards):
        x = 6.26 + i * 3.30
        rect(s, x, 4.98, 3.09, 1.24, fill=fill, line=None if dark else BORDER, lw=1.25)
        txt(s, x + 0.24, 5.10, 2.61, 0.58,
            [(n, 28, True, WHITE if dark else BORDER),
             (" " + u, 12.5, True, SKY if dark else MUTED)], anchor=MT)
        txt(s, x + 0.24, 5.72, 2.61, 0.42, [(d, 12.5)],
            color=SKY if dark else BODY, anchor=MT, ls=1.28)
    txt(s, M, 6.38, CW, 0.44,
        [("图谱的价值不在于多存了数据，在于补上查询词与候选词之间那条字面上不存在的边。",
          14, True)], color=BORDER, anchor=MT)
    foot(s, pn)


def findings(pn):
    """三栏卡片页：每栏 = 一个面板 + 顶栏 + 一句话结论。"""
    s = new_slide("content", pn)
    head(s, "三组发现", "三个发现，一页说完")
    finds = [
        ("01", "词面对齐做不到",
         "查询写「水位监测」，文档写「雷达水位测量」，两边根本不共用同一个词。"),
        ("02", "图谱补的是那一跳",
         "从「设备巡检」一跳到「传感器状态监测」，多召回 3 条候选，P95 只多 6 ms。"),
        ("03", "融合赢在排序不在召回",
         "三条路径召回条数一样，MRR 从 0.463 抬到 0.677，差别全在顺序。"),
    ]
    for i, (no, t, d) in enumerate(finds):
        x = X3[i]
        panel(s, x, 1.74, BW3, 4.42)
        headbar(s, x, 1.74, BW3, f"{no}　{t}", h=0.52, size=16)
        txt(s, x + 0.30, 2.50, BW3 - 0.60, 2.20, [(d, 14.5)], color=BODY, ls=1.36)
        rule(s, x + 0.30, 5.10, 1.00, color=kit.ACC)
        tail = ["所以同义召回必须靠语义层补。", "所以关系型数据值得单独建一层。",
                "所以时延预算要按最贵的那条路径定。"][i]
        txt(s, x + 0.30, 5.30, BW3 - 0.60, 0.72, [(tail, 13.5, True)], color=BORDER, ls=1.32)
    txt(s, M, 6.36, CW, 0.44,
        [("这三条不是「方法不行」，是「哪一层该放什么」——分层是从量级算出来的，不是说法。",
          14.5, True)], color=BORDER, anchor=MT)
    foot(s, pn)


def summary(pn):
    """小结页：三条已站住的结论 + 接下来计划做的三件事。"""
    s = new_slide("content", pn)
    head(s, "小结 · 下一步", "现在站住的三条，接下来计划做的三件事")
    finds = [
        ("01", "语义层不可省", "纯倒排漏掉的三条里，有两条靠语义重排就回来了，成本 53 ms。"),
        ("02", "图谱层性价比高", "多召回 3 条、P95 只多 6 ms，代价是入库与实体对齐的工程量。"),
        ("03", "融合的收益在排序", "召回数不变、MRR 抬 0.165，但端到端 P95 翻了 近 12 倍。"),
    ]
    for i, (no, t, d) in enumerate(finds):
        x = X3[i]
        panel(s, x, 1.74, BW3, 2.44)
        headbar(s, x, 1.74, BW3, f"{no}　{t}", h=0.48, size=16)
        txt(s, x + 0.28, 2.42, BW3 - 0.56, 1.62, [(d, 13.5)], color=BODY, ls=1.34)
    rect(s, M, 4.52, CW, 1.60, fill=BAND, line=None)
    txt(s, M + 0.34, 4.64, CW - 0.68, 0.30,
        [("接下来计划做的三件事　", 15.5, True, SKY)], anchor=MT)
    nxt = ["补第三方标注终审——19 条对齐之后才能写进论文，这是最紧的一件",
           "把图谱层的冷启动实体对齐单独拆一版，看召回能再抬多少",
           "把融合路径的时延拆开测，看 143 ms 花在哪一段"]
    for i, t in enumerate(nxt):
        txt(s, M + 0.34, 5.00 + i * 0.34, CW - 0.68, 0.30,
            [(f"{i + 1}．　", 14.5, True, SKY), (t, 14.5, None, WHITE)], anchor=MT)
    txt(s, M, 6.36, CW, 0.44,
        [("现在站住的结论：召回该留给确定性工程，生成式模型只做最后一层。", 15, True)],
        color=BORDER, anchor=MT)
    foot(s, pn)


def twocol(pn):
    """两栏对照页（`new_slide("two_col", pn)`）。

    这一页存在的理由是**覆盖 two_col 版式分支**：两栏内容版式自带的两条内容占位符
    会被 new_slide 删掉，所以两栏的构图必须页面自己画——版式只负责标题带。
    栏宽 5.85 / 栏距 0.31 照版式自带的占位符几何来，右边界正好落在 R = M + CW。
    """
    s = new_slide("two_col", pn)
    head(s, "取舍 · 两条路线", "同样三个月的预算，两条路线各买到什么")
    cols = [
        ("A", "先补语义层", [
            ("召回立刻变好", "纯倒排漏掉的三条里两条回来了，端到端只多 53 ms。"),
            ("代价是可控的", "多一个交叉编码器服务，显存还在预算内，不动数据资产。"),
            ("风险是天花板", "词面完全接不上的那类查询，语义层也接不上。"),
        ]),
        ("B", "先补图谱层", [
            ("关系型查询能答", "从「设备巡检」跨一跳到「传感器状态监测」，字面搜不到也能召回。"),
            ("代价是冷启动", "实体对齐要人工确认，入库与维护是一笔持续投入。"),
            ("风险是数据债", "实体对不齐的时候，图谱层的结论只能写相对排序。"),
        ]),
    ]
    for i, (no, title, items) in enumerate(cols):
        x = M if i == 0 else 6.82
        w = 5.85
        panel(s, x, 1.74, w, 4.20)
        headbar(s, x, 1.74, w, f"{no}　{title}", h=0.50, size=16)
        for j, (t, d) in enumerate(items):
            y = 2.42 + j * 1.02
            txt(s, x + 0.30, y, w - 0.60, 0.34, [(t, 15.5, True)], color=BAND, anchor=MT)
            txt(s, x + 0.30, y + 0.38, w - 0.60, 0.58, [(d, 13.5)], color=BODY, ls=1.32)
    rect(s, M, 6.10, CW, 0.72, fill=TINT, line=BORDER, lw=1.5)
    txt(s, M + 0.32, 6.10, CW - 0.64, 0.72,
        [("两条不冲突：这一版先上 A，下一版补 B——顺序问题，不是取舍问题。", 15, True)],
        color=BAND, anchor=MT)
    foot(s, pn)


def timeline(pn):
    """时间线 / 里程碑页：横轴 + 四个节点，轴上下各挂一段文字。

    排法上有两处刻意的选择：
      · 轴线画成一条 `rule()`，节点画成小圆——它们都是装饰，**文字全部挂在轴的两侧**，
        这样纵向被文字和卡片连成一条线，闸门四才查不出大空隙；
      · 每个节点写「解决了什么 + 留下了什么」两件事。里程碑页只写产出不写遗留问题，
        听众会以为这条路没有代价。
    """
    s = new_slide("content", pn)
    head(s, "里程碑", "四个阶段：每一步解决了什么，又留下了什么")
    txt(s, M, 1.74, CW, 0.44,
        [("按依赖顺序走的路：后一阶段的数据资产是前一阶段的产出，跳步做不出上一段的数。",
          15.5, True)], color=BAND, anchor=MT)
    stages = [
        ("阶段一", "第 1–4 周", "打通词面召回",
         "倒排上线，示例查询集 1,337 条入库。产出是可复现的评测脚本。"),
        ("阶段二", "第 5–9 周", "补语义重排",
         "接入交叉编码器，排序质量抬升。遗留：词面完全不接的查询接不上。"),
        ("阶段三", "第 10–16 周", "建关系层",
         "实体对齐 + 一跳扩展。遗留：冷启动阶段要对齐靠人工确认。"),
        ("阶段四", "第 17–20 周", "三路融合",
         "召回并联、排序串联。代价是端到端 P95 到了 143 ms。"),
    ]
    rule(s, M, 3.02, CW, color=kit.ACC, h=0.022)
    for i, (ph, wk, t, d) in enumerate(stages):
        x = X4[i]
        cx = x + BW4 / 2
        txt(s, x, 2.28, BW4, 0.36, [(ph, 17, True)], color=BAND, align=CT, anchor=MT)
        txt(s, x, 2.66, BW4, 0.30, [(wk, 12.5)], color=MUTED, align=CT, anchor=MT)
        rect(s, cx - 0.11, 2.92, 0.22, 0.22, fill=BAND, line=None, shape=MSO_SHAPE.OVAL)
        txt(s, x, 3.20, BW4, 0.36, [(t, 15, True)], color=INK, align=CT, anchor=MT)
        txt(s, x, 3.60, BW4, 0.92, [(d, 13)], color=BODY, ls=1.32)
    cards = [("这一版实际只做到第三阶段", "第四阶段的融合结果来自上一版实验，尚未在本环境复跑。", TINT),
             ("为什么不用甘特图", "四个阶段有严格依赖、没有并行，横向节点比时间条更诚实。", TINT)]
    for i, (t, d, fill) in enumerate(cards):
        x = M + i * 6.02
        rect(s, x, 4.70, 5.99, 1.10, fill=fill, line=BORDER, lw=1.5)
        txt(s, x + 0.30, 4.70, 5.39, 0.42, [(t, 15, True, BAND)], anchor=MT)
        txt(s, x + 0.30, 5.16, 5.39, 0.58, [(d, 13)], color=BODY, anchor=MT, ls=1.30)
    rect(s, M, 6.00, CW, 0.72, fill=BAND, line=None)
    txt(s, M + 0.32, 6.00, CW - 0.64, 0.72,
        [("排期是承诺，遗留问题是事实：这一讲只承诺前三段。", 15, True)], color=WHITE, anchor=MT)
    foot(s, pn)


def quadrant(pn):
    """四象限矩阵页：2×2 面板 + 两条轴说明 + 底部一句决策。

    轴说明不画箭头也不做竖排文字——竖排中文在 PowerPoint 里要靠「文本框竖排」实现，
    一旦换了字体或行距就会歪。两句横排的「投入大/小」「收益小/大」信息量一样，稳得多。
    """
    s = new_slide("content", pn)
    head(s, "定位", "四个象限：这版该动哪一格")
    txt(s, M, 1.72, CW, 0.40,
        [("把候选工作按「投入」和「收益」两轴摊开，先补空格里最便宜的那一格。", 15, True)],
        color=BAND, anchor=MT)
    qx0, qw, qh, gap = M + 0.92, (CW - 0.92 - 0.22) / 2, 1.50, 0.22
    cells = [
        (0, 0, "高投入 · 高收益", "图谱层", "建库、实体对齐、持续维护都要人。收益是多跳关系型查询。", "暂缓"),
        (1, 0, "低投入 · 高收益", "语义重排", "一个模型服务就能上线，排序质量立刻抬一截。", "本版做"),
        (0, 1, "高投入 · 低收益", "端到端重训", "要重标数据、重跑全部实验，口径还不能完全对齐。", "不做"),
        (1, 1, "低投入 · 低收益", "词面同义词表", "手工维护，效果随语料变化，几个月后基本要重来。", "不做"),
    ]
    for cx, cy, t, name, d, mark in cells:
        x = qx0 + cx * (qw + gap)
        y = 2.24 + cy * (qh + gap)
        panel(s, x, y, qw, qh)
        headbar(s, x, y, qw, f"{name}　·　{mark}", h=0.46, size=15)
        txt(s, x + 0.28, y + 0.58, qw - 0.56, 0.80, [(d, 13.5)], color=BODY, ls=1.32)
    for i, lab in enumerate(("投入小", "投入大")):
        txt(s, M, 2.24 + i * (qh + gap), 0.86, qh, [(lab, 12.5, True)], color=MUTED,
            align=CT, anchor=MT)
    for i, lab in enumerate(("收益小", "收益高")):
        txt(s, qx0 + i * (qw + gap), 5.50, qw, 0.30, [(lab, 12.5, True)], color=MUTED,
            align=CT, anchor=MT)
    rect(s, M, 5.94, CW, 0.74, fill=BAND, line=None)
    txt(s, M + 0.32, 5.94, CW - 0.64, 0.74,
        [("先补右上那一格：投入最小的两格里，只有语义重排能在这一版做完。", 15, True)],
        color=WHITE, anchor=MT)
    foot(s, pn)


def kpiwall(pn):
    """KPI 数字墙：3×2 六格，每格「大数字 + 口径 + 一句变化」。

    和数据大数字页（四格）的区别在密度：六格版给的是「一屏看完一版实验的体检报告」，
    每格都必须能独立成立——单独截一格出来贴在周报里也读得懂。
    """
    s = new_slide("content", pn)
    head(s, "体检", "六格看完这一版：哪些数站得住，哪些还在动")
    cards = [
        ("1,337", "条", "查询集", "人工挑选并逐条标注相关文档", "较上一版 +112 条"),
        ("6,551", "条", "候选池", "四条路径共用同一份候选池，不换数据", "本版未变"),
        ("19", "条", "标注集", "两人独立标注后对齐，未经第三方终审", "本版未变"),
        ("0.463", "", "排序质量", "融合路径相对词面基线的相对提升", "较词面 +0.145"),
        ("143", "ms", "端到端 P95", "含重排与图谱扩展的完整链路", "较上版 +31 ms"),
        ("11", "GB", "显存预算", "全程本地推理，未触发降频", "本版未变"),
    ]
    for i, (n, u, t, d, delta) in enumerate(cards):
        x = X3[i % 3]
        y = 1.74 + (i // 3) * 2.52
        panel(s, x, y, BW3, 2.30)
        txt(s, x + 0.28, y + 0.20, BW3 - 0.56, 0.80,
            [(n, 30, True, BAND), (" " + u, 15, True, MUTED)], anchor=MT)
        txt(s, x + 0.28, y + 1.06, BW3 - 0.56, 0.34, [(t, 15, True)], color=INK, anchor=MT)
        rule(s, x + 0.28, y + 1.48, 0.90, color=kit.ACC)
        txt(s, x + 0.28, y + 1.60, BW3 - 0.56, 0.58, [(d, 12.5)], color=BODY, ls=1.30)
        txt(s, x + 0.28, y + 1.98, BW3 - 0.56, 0.26, [(delta, 12, True)], color=BORDER, anchor=MT)
    foot(s, pn)


def figpage(pn):
    """图版页：左解释 + 右图（matplotlib 出的位图）。

    这一页存在的理由是**给 `kit.fig_fit()` 一条回归线**：图片必须等比 + 居中，
    拉伸变形肉眼查不出来，但闸门一（清单里的实测 ar）和闸门二（PDF 里的落位）会查。
    图是 `python scripts/figstyle.py` 生成的示例图，比例完全由 figsize 决定。
    """
    s = new_slide("content", pn)
    head(s, "量级", "四段链路的时延差一个数量级")
    txt(s, M, 1.76, 4.55, 1.40,
        [("线性轴上，前三段的柱子几乎一样高——不是它们成本相同，"
          "是被最大那一根压平了。", 15, True)], color=BORDER, ls=1.36)
    panel(s, M, 3.34, 4.55, 1.56)
    txt(s, M + 0.28, 3.52, 3.99, 0.36, [("所以这张图用对数轴", 14.5, True)], color=BAND, anchor=MT)
    txt(s, M + 0.28, 3.94, 3.99, 0.82,
        [("对数轴上等距离代表等倍数，读者能看出「重排比词面贵 4 倍、"
          "融合比词面贵近 12 倍」。刻度标签已隐藏，只留网格。", 13)], color=BODY, ls=1.30)
    fig_fit(s, 5.55, 1.76, 7.12, 3.28, str(kit.PKG / "assets" / "figures" / "fig_demo_latency.png"))
    cap(s, 5.55, 5.12, 7.12, "图 1", "各段链路端到端 P95 时延（对数轴，单位 ms；示例数据）")
    rect(s, M, 5.66, CW, 0.94, fill=BAND, line=None)
    txt(s, M + 0.32, 5.66, CW - 0.64, 0.94,
        [("要讲的是倍数关系，不是绝对值：", 15, True, SKY),
         ("重排的 53 ms 买得到排序提升，融合的 143 ms 买的是整合，"
          "两者不能放在同一个预算里谈。", 15, True, WHITE)], anchor=MT, ls=1.34)
    foot(s, pn)


def close(pn):
    """结尾页。整块低饱和底 + 上下细线由版式自带，这里只排字。

    末尾那行落款放在底色之下（y 6.62）：压在底色上虽然也够对比度，
    但底色里已经有一行分隔细线，再压一行字会显得挤。
    """
    s = new_slide("close", pn)
    txt(s, 3.50, 2.56, 6.33, 1.34, [("谢谢聆听", 54, True)], color=INK, align=CT, anchor=MT)
    txt(s, 3.50, 4.32, 6.33, 0.50, [("欢迎批评指正", 18, True)], color=MUTED, align=CT, anchor=MT)
    txt(s, 3.50, 5.00, 6.33, 1.06,
        [("检索交给工程，生成交给模型。", 17, True, BAND),
         ("词面接不上的那一跳，图谱补得上。", 17, True, INK)],
        align=CT, anchor=MT, ls=1.40, latin=kit.FORM)
    txt(s, 3.50, 6.62, 6.33, 0.32, [(f"示例课题组 · 检索与评测模块 · {MEET}", 11.5)],
        color=MUTED, spc=2, align=CT, anchor=MT)


PAGES = [cover, roadmap, thesis, compare, datanums, layers, chart, findings, summary,
         twocol, timeline, quadrant, kpiwall, figpage, close]


# ────────────────────────────── 讲稿 ──────────────────────────────
# 讲稿的唯一来源：kit.SCRIPT。同一份内容同时产出 PPT 内嵌备注与 markdown，
# 不允许各写各的。kit.notes(pn=…) 与 kit.write_script_md() 都从它读。
# 写法约定见 references/narration.md。
kit.SCRIPT = {
    1: ("示例汇报 · 三层检索架构的落地与评测", "0:00–0:40",
        "各位老师好。今天汇报一件事：把一套三层检索架构从论文搬进生产，"
        "跑了三组实验，以及还差哪三件事。\n\n"
        "（指屏）封面这一行是整份工作的压缩式子：检索等于索引召回加图谱扩展加语义重排。\n\n"
        "（停顿）先说清：这一讲所有数字都是示例数据集上的虚构值，只用来演示汇报结构。开始。"),
    2: ("四步走完：问题、实验、结果、还差什么", "0:40–2:00",
        "先交一下底，今天四步。\n\n"
        "（指四个编号块）第一步，把一个方法从论文搬进生产，先要回答它到底比基线好在哪一档。"
        "第二步，实验设计：同一批查询、同一套指标，三种检索路径逐层叠加。"
        "第三步，三组结果。第四步，只报已经算完的数，最后交代还差哪三件事。\n\n"
        "（指蓝色总纲带）总纲一句：检索该分给确定性工程，索引管召回，图谱管关系，"
        "模型只管把答案组织成话。\n\n"
        "（指下方浅蓝框，这句特意说在前面）口径先说清：标注集只有 19 条，是两个人独立标完再对齐的，"
        "没有第三方终审。所以这一讲只看相对排序，不看绝对值。"),
    3: ("检索栈拆成三层：各自最贵的那部分资源都不一样", "2:00–3:40",
        "这一页把检索栈拆成三层，各自算账。\n\n"
        "（指左栏）索引层，结构化倒排，毫秒级返回，可复现可审计每条能溯源。"
        "短板是同义说法召回不到。\n\n"
        "（指右栏）图谱层，多跳关联，能把词面不接的两端接上。代价是入库与维护成本，"
        "短板是冷启动实体对不齐。\n\n"
        "（指底部）这条结论是：大模型不是不能用，是不能放在召回位置上。"),
    4: ("同一批查询，四条路径的同一把尺", "3:40–5:20",
        "这张表是今天的主表。四条路径，同一批查询，同一份标注，同一个评测脚本。\n\n"
        "（逐行指）纯倒排召回数 5 条，MRR 0.318。倒排加语义召回数 8 条，MRR 抬到 0.463。"
        "倒排加图谱也是 8 条，但 MRR 0.512，而且 P95 只多 6 毫秒。三条融合召回数没变，"
        "MRR 到 0.677。\n\n"
        "（停顿）表里最该看的不是 0.677，是它比 0.463 高出来的那一段——差别全在顺序。\n\n"
        "（指右下）代价也写在表里：融合路径 P95 到 143 毫秒，是纯倒排的 近 12 倍，这是实付的价钱。"),
    5: ("实验跑在什么数据、什么硬件上", "5:20–6:30",
        "数据侧：1,337 条查询，6,551 条候选池，标注集 19 条，硬件是 11 GB 显存的本地机器。\n\n"
        "（指右侧浅蓝框）顺便说一个意外收获：把评测脚本改成配置文件驱动之后，"
        "同一份脚本能跑全部配置，实验记录里能查到每一次的随机种子和环境版本。"
        "这件事本身就该单独立项——可复现不是加分项，是让别人的结论能被质疑的前提。"),
    6: ("四层结构：大模型只进最上面一层", "6:30–8:20",
        "这一页是我这套架构的分层。\n\n"
        "（自上往下指）第四层生成与交互，用本地小模型或云端大模型，把候选组织成答案，"
        "这一层是概率系统。第三层语义重排，用交叉编码器把顺序排对，代价是毫秒级。"
        "第二层结构与关系，知识图谱管多跳关联和实体对齐，确定性、可审计。"
        "第一层数据资产，带溯源的原始数据。\n\n"
        "（指右栏）大模型不是不能用，是只能放在最上面一层。每一层干自己最划算的事，"
        "这条分界线要拿真实数据去验证，不是靠说法。"),
    7: ("图谱补的是字面上找不到的那一跳", "8:20–10:00",
        "看这张柱状图。四根柱子依次是纯倒排、加语义重排、加图谱扩展、加融合排序。\n\n"
        "（指左栏）查询原文只有「设备巡检」四个字，没有「传感器」。纯倒排只能回到查询原文的词面。"
        "图谱把「设备巡检」和电力、电网连在一起，一跳就跨到「传感器状态监测」这条文档。\n\n"
        "（指下方两张卡）图谱扩展单独多召回 3 条，融合排序再贡献 0.054 的 MRR。\n\n"
        "（停顿）所以图谱的价值不在于多存了数据，在于补上查询词与候选词之间那条字面上不存在的边。"),
    8: ("三个发现，一页说完", "10:00–11:30",
        "三个发现一页说完。\n\n"
        "（指第一栏）词面对齐做不到。查询写「水位监测」，文档写「雷达水位测量」，两边不共用同一个词。\n\n"
        "（指第二栏）图谱补的是那一跳。多召回 3 条，P95 只多 6 毫秒。\n\n"
        "（指第三栏）融合赢在排序不在召回。召回数一样，MRR 差了一截。\n\n"
        "（指底部）这三条不是方法不行，是哪一层该放什么。分层是从量级算出来的，不是说法。"),
    9: ("现在站住的三条，接下来计划做的三件事", "11:30–13:00",
        "现在站住的三条：语义层不可省，纯倒排漏掉的三条里有两条靠语义重排就回来了；"
        "图谱层性价比高，多召回 3 条只多 6 毫秒；融合的收益在排序，召回数不变但时延翻了 近 12 倍。\n\n"
        "接下来计划做的三件事：补第三方标注终审，这是最紧的一件；"
        "把图谱层的冷启动实体对齐单独拆一版；把融合路径的时延拆开测，看 143 毫秒花在哪一段。\n\n"
        "（停顿）这三件是计划，不是已经做完的。"),
    10: ("同样三个月的预算，两条路线各买到什么", "13:00–14:00",
         "这一页是取舍。两条路线，各买到什么，各欠什么。\n\n"
         "（指左栏）先补语义层：纯倒排漏掉的三条里有两条回来了，端到端只多 53 毫秒，"
         "代价是一个交叉编码器服务。但词面完全接不上的那类查询，语义层也接不上，这是它的天花板。\n\n"
         "（指右栏）先补图谱层：从「设备巡检」跨一跳到「传感器状态监测」，字面搜不到的也能召回。"
         "代价是冷启动的实体对齐要人工确认，入库与维护是持续投入。\n\n"
         "（指底部）这两条并不冲突：这一版先上 A，下一版补 B。这是顺序问题，不是取舍问题。"),
    11: ("四个阶段：每一步解决了什么，又留下了什么", "14:00–15:00",
         "这一页是排期。按依赖顺序走的：后一阶段的数据资产是前一阶段的产出。\n\n"
         "（指第一段）阶段一打通词面召回，1,337 条查询入库，产出是可复现的评测脚本。"
         "阶段二补语义重排，排序质量抬升，遗留问题是词面完全不接的查询接不上。\n\n"
         "（指第三、四段）阶段三建关系层，冷启动要人工确认实体对齐；"
         "阶段四三路融合，代价是端到端 143 毫秒。\n\n"
         "（指下方两张卡，先把话说在前面）这一版实际只做到第三阶段，"
         "第四阶段的数来自上一版实验，尚未在本环境复跑。排期是承诺，遗留问题是事实。"),
    12: ("四个象限：这版该动哪一格", "15:00–16:00",
         "把候选工作按投入和收益两轴摊开，看该补哪一格。\n\n"
         "（指左下）高投入高收益是图谱层，要建库、要人工对齐、还要持续维护。"
         "低投入高收益是语义重排，一个服务就能上线。\n\n"
         "（指右侧两格）高投入低收益是端到端重训，要重标数据重跑全部实验；"
         "低投入低收益是同义词表，手工维护、几个月后基本要重来。\n\n"
         "所以先补右上那一格：投入最小的两格里，只有语义重排这一版做得完。"),
    13: ("六格看完这一版：哪些数站得住，哪些还在动", "16:00–17:00",
         "一页体检报告，六格。左边三格是数据规模：1,337 条查询、6,551 条候选池、19 条标注集。\n\n"
         "（指右边三格）排序质量 0.463，比词面基线抬 0.145；端到端 P95 143 毫秒，"
         "比上版多 31 毫秒；显存预算 11 GB，本地推理没有降频。\n\n"
         "（停顿）每格下面那行是相对上一版的变化，没变的我照实写没变，"
         "不拿没动过的数充业绩。"),
    14: ("四段链路的时延差一个数量级", "17:00–18:00",
         "这张图用对数轴，刻度标签已经隐藏，只留网格。\n\n"
         "（指图）线性轴上前三段几乎一样高，那不是它们成本相同，是被最大那一根压平了。"
         "对数轴上等距离是等倍数：重排比词面贵 4 倍，融合比词面贵近 12 倍。\n\n"
         "要讲的是倍数关系不是绝对值：重排的 53 毫秒买得到排序提升，"
         "融合的 143 毫秒买的是整合，两者不能放在同一个预算里谈。"),
    15: ("结尾", "18:00–18:30",
        "我的汇报到这里。\n\n"
        "一句话收住：检索交给工程，生成交给模型；词面接不上的那一跳，图谱补得上。\n\n"
        "（停顿）请老师指正。"),
}

QNA = [
    ("标注集才 19 条，这个数可信吗？",
     "绝对值不可信，我只看相对排序。19 条是两个人独立标完再对齐的结果，分歧全部留痕，"
     "但没有第三方终审，所以还没进论文。补终审是接下来第一件事。"),
    ("融合路径 P95 143 毫秒，能上线吗？",
     "单看这个数分场景。交互式检索按 P95 算就要压预算，这一版我承认压不住；"
     "异步批处理可以接受。时延拆解是接下来第三件事。"),
    ("图谱的实体对齐是哪来的？",
     "这一版的实体对是半自动的，人工确认过一部分，剩下的是模型候选加规则兜底。"
     "准确率我没测，所以图谱层的结论我也只敢写相对排序。"),
]


def build(out_dir: Path):
    kit.OUT_DIR = out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    deck = dict(out="示例汇报_中文科研组会.pptx",
                manifest="示例汇报_中文科研组会_layout.json",
                total=PAGES_N, foot=FOOT, title=TITLE)
    kit.start_deck(deck)
    kit.PRS = kit.open_base()
    for pn, fn in enumerate(PAGES, 1):
        fn(pn)
        kit.notes(kit.PRS.slides[-1], pn=pn)      # 台词取自 kit.SCRIPT，不在页面里再抄一份
    kit.save_deck(kit.PRS, deck)
    script_md = kit.write_script_md(
        deck,                                   # pages 省略 → 由 kit.SCRIPT 生成
        path=out_dir / "示例汇报_中文科研组会_讲稿.md",
        title="示例汇报 · 三层检索架构的落地与评测",
        meta=f"组会时间 {MEET}　·　{PAGES_N} 页　·　内容为示例数据集上的虚构值",
        qna=QNA)
    return script_md


def _rebind_palette():
    """把本模块 import 期绑死的色板 / 栅格常量重新指到 kit 当前生效的值。

    顶部那句 `from kit import BAND, BODY, ...` 在 import 期就把**值**抄进了本模块。
    `kit.load_theme()` 改的是 kit 自己的全局，抄进来的这一份不会跟着变——于是
    `--theme midnight` 出来的稿子是「母版深蓝、正文还是 academic-blue 的深灰」，
    投影上一片糊，闸门三一口气报三十几处对比度不足。

    无 `--theme` 时不调用它：学术蓝那组值与 kit 初值逐字相同，这条修复对默认
    构建零影响。
    """
    global BAND, BODY, BORDER, BROWN, INK, MUTED, SKY, TINT, WHITE, M, CW
    BAND, BODY, BORDER, BROWN, INK, MUTED, SKY, TINT, WHITE, M, CW = (
        kit.BAND, kit.BODY, kit.BORDER, kit.BROWN, kit.INK, kit.MUTED, kit.SKY,
        kit.TINT, kit.WHITE, kit.M, kit.CW)


def main():
    ap = argparse.ArgumentParser(
        description="生成示例 deck（15 页，中文科研组会）",
        epilog="参数拼错会直接报错退出——把 --tempalte 当成 --template 蒙混过去，"
               "会以为套了自定义模板、实际套的是内置母版，而四道闸门照样全绿。")
    ap.add_argument("--zh-kaiti", action="store_true",
                    help="中文改楷体，产物另存到 out/variant_kaiti/")
    ap.add_argument("--template", metavar="PPTX",
                    help="换基座模板（.pptx/.potx）。只换基座文件，"
                         "palette/fonts/grid/layouts 仍以已载入的 theme.json 为准")
    ap.add_argument("--out", metavar="DIR", help="产物目录（默认 out/，verify_all.py 用 OUT_DIR 注入）")
    ap.add_argument("--theme", default="", choices=sorted(kit.BUILTIN_THEMES),
                    help="切内置主题（母版与页面色板一起换），产物另存到 out/theme_<name>/")
    a = ap.parse_args()

    # OUT_DIR 由 verify_all.py 注入；直接手跑时落在当前目录的 out/
    out = Path(os.environ.get("OUT_DIR") or (Path.cwd() / "out"))
    if a.theme:
        # 必须在建 deck 之前换主题，而且调完还要 rebind：from kit import 来的
        # BAND / TINT 等名字在 import 期就把值绑死了（和 EAK 那条坑同源）。
        kit.load_theme(name=a.theme)
        _rebind_palette()
        out = out / f"theme_{a.theme}"
    if a.zh_kaiti:
        # 中文改楷体。必须在任何页面函数 import 之后改 kit.EAK 才生效——
        # 传默认参数（def txt(..., ea=EAK)）在 def 时就绑定了，改这里没用。
        kit.EAK = "楷体"
        out = out / "variant_kaiti"
    if a.template:
        kit.kit_config(template=a.template)
    if a.out:
        out = Path(a.out)
    print(f"== 示例汇报 · {PAGES_N} 页 ==")
    print(f"   主题 {kit.LAYOUTS}　模板 {kit.TPL}")
    print(f"   字体 西文 {kit.LAT} / 中文 {kit.EAK}\n")
    build(out)
    print(f"\n产物在 {out}")


if __name__ == "__main__":
    sys.exit(main() or 0)
