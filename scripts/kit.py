# -*- coding: utf-8 -*-
"""
中文科研汇报 PPT · 通用组件库
================================

设计意图
--------
这个文件只解决一件事：把「按模板排版」这件事里所有与具体项目无关的部分固定下来，
让写页面的人只关心内容本身。

三条硬约束（改动前先读 references/pitfalls.md）：

1. `run.font.name` 只写 `<a:latin>`。中文字形要另注入 `<a:ea typeface="..."/>`，
   而且 `<a:ea>` 必须排在 `<a:latin>` 之后——顺序反了 PowerPoint 读不到中文字体。
   `_style()` 已经把这件事做掉，页面代码不要绕开它自己设 font.name。

2. 中文字体不能用默认参数传递。`def txt(..., ea=EAK)` 会在 def 时把当时的 EAK
   绑定成默认参数，之后外部改 `kit.EAK` 一个字也改不到，字体名照样写进 XML、
   就是不换。所以一律写 `ea=None`，在函数体里 `ea = ea or EAK` 现取。
   同理 `latin=None` → 现取 LAT。

3. 版式装饰是「非占位符普通形状」，`add_slide()` 只克隆占位符，装饰会自动渲染在
   页面下层。所以页面函数一行装饰都不要加——任何自绘色块盖上去都会盖住版式自带的
   标题带或封面装饰。新增一页走 `new_slide(版式, 页码)`，它会删掉克隆来的占位符。

主题
----
本模块的所有色板 / 字体 / 栅格常量都可以被 `theme.json` 覆盖（见 `load_theme()`）。
没有 theme.json 时用内置中性主题（assets/theme-neutral.pptx 配套的那一套）。
用户自带模板的接入流程：scripts/analyze_template.py 反推 → theme.json → load_theme()。

用法
----
    import kit
    kit.load_theme("theme.json")        # 可选；不调用就用内置中性主题
    kit.kit_config(template="D:/xx/我的模板.pptx")   # 只换基座，保留已载入主题
    kit.PRS = kit.open_base()            # 清空模板示例页并返回 Presentation
    ...
    kit.SCRIPT[1] = ("标题", "0:00–0:40", "讲稿…")     # 讲稿唯一来源
    kit.notes(slide, pn=1)                            # 按页码取台词写备注
    kit.write_script_md(deck, qna=QNA)                # markdown 讲稿由 SCRIPT 生成
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ── 路径 ──
BASE = Path(__file__).resolve().parent
PKG = BASE.parent                      # skill 包根目录
ASSET = PKG / "assets"                 # 插图目录（kit.pic 默认从这里取）
OUT_DIR = Path.cwd()                   # 产物落盘目录
THEME_FILE = PKG / "theme.json"        # 可选的主题覆盖文件

EMU_IN = 914400
R = None                               # 右边界（M + CW），apply_theme 里算


def In(v):
    """英寸 → EMU。栅格里所有坐标单位都是英寸。"""
    return Emu(int(round(v * EMU_IN)))


# ── 内置中性主题 ──
# 对比度是它与白底（或在标注色块上）的 WCAG 2.1 比值，正文一律 ≥ 4.5，
# 大字（≥18pt 或 ≥14pt 粗体）≥ 3.0。禁用场景写在 comments 里，改色时一并看。
NEUTRAL_THEME = {
    "name": "theme-neutral",
    "template": "assets/theme-neutral.pptx",
    # 页面底色（六字符十六进制）。**刻意不放进 palette**：palette 的 WHITE 键是
    # 「深色块上的文字」，和「页面底」是两件事——深色主题里页面底是深蓝、
    # WHITE 仍是浅字（或者反过来，见 references/design-language.md §2c）。
    # 旧 theme.json 没有这个字段时回落到 "FFFFFF"，行为与旧版完全一致。
    "background": "FFFFFF",
    "palette": {
        "BAND": "006DB8",      # 标题带 / 结论带 / 强调数字：承白字 5.41:1
        "BORDER": "4472C4",    # 描边 2.25pt / 淡底上的强调字：白底 4.72:1
        "ACC": "6096E6",       # 只做装饰线：白底 3.00:1，承不住小字
        "SKY": "E8F2FB",       # 只放在 BAND 上（4.77:1），白底太浅不可用
        "TABBG": "D3D9E9",     # 编号标签底：承深字 12.99:1
        "NUM": "7B8DBE",       # 编号圆：白字 20pt 粗体 3.28:1
        "TINT": "F2F6FB",      # 淡底卡片：承 BAND 字 4.99:1
        "BROWN": "B39B77",     # 纯装饰：白字仅 2.67:1，绝不承文字
        "INK": "0F1423",       # 标题：18.35:1
        "BODY": "262626",      # 正文：15.13:1
        "MUTED": "56646E",     # 弱化文字：6.10:1
        "WHITE": "FFFFFF",
    },
    "fonts": {
        # 西文统一 Times New Roman：学术语境正式、可信。早期版本用的是
        # 装饰性手写体与无衬线体，都比它宽 6%–17%，会让按窄字体量出来的框溢出。
        # 选型理由与被否方案见 references/design-language.md。
        "latin": "Times New Roman",
        "ea": "微软雅黑",       # 中文默认；改成 "楷体" 走楷体变体
        "num": "Times New Roman",   # 大序号
        "formula": "Times New Roman",  # 公式；用 Bold 取得等宽感，见 design-language.md
    },
    "grid": {
        "W": 13.33, "H": 7.5,       # 页面尺寸（16:9）
        "M": 0.66,                  # 左右边距
        "CW": 12.01,                # 内容区宽
        "FTY": 6.96,                # 页脚文字基线
        "FOOT_TOP": 6.98,           # 页脚带起点：正文不许越进来
        "CT_TOP": 1.66,             # 内容区顶（标题带 + 章节标签之下）
        "CT_BOT": 6.94,             # 内容区底
    },
    "layouts": {
        "cover": "标题幻灯片",
        "content": "标题和内容",
        "two_col": "两栏内容",
        "section": "节标题",
        "close": "末尾幻灯片",
    },
}

# ── 内置六套主题 ──
# 母版由 make_neutral_template.py --theme <name> 生成在 assets/themes/<name>.pptx，
# 每套同时留一份 <name>.json 供本模块读。**色板改这里也要改那个脚本**——
# 两处不一致时，母版是蓝的、页面是红的，比只有一套色板更糟。
#
# 每套的实测对比度与禁用场景写在 references/design-language.md §2b（三套浅色）
# 与 §2c（midnight 深色 / forest-green / warm-sand）。
# 选型口径：默认 academic-blue；投影偏暗或场合偏正式用 scholar-red；
# 打印稿、外部评审、不想让人从颜色上先入为主时用 minimal-mono；
# 暗投影/路演/keynote/大屏用 midnight；环境、生态、材料、农业、地学用 forest-green；
# 人文、社科、医学、报告附件用 warm-sand。
BUILTIN_THEMES = {
    "academic-blue": {          # 内置中性主题的正式名（默认，向后兼容）
        "name": "academic-blue",
        "template": "assets/themes/academic-blue.pptx",
        "background": "FFFFFF",
        "palette": dict(NEUTRAL_THEME["palette"]),
    },
    "scholar-red": {
        "name": "scholar-red",
        "template": "assets/themes/scholar-red.pptx",
        "background": "FFFFFF",
        "palette": {
            "BAND": "8A2B34",      # 暗红标题带 / 结论带：承白字 8.48:1
            "BORDER": "9E3A44",    # 描边与淡底上的强调字：白底 6.68:1
            "ACC": "C06C74",       # 只做装饰线：白底 3.72:1，承不住小字
            "SKY": "F7E9EA",       # 只放在 BAND 上（7.19:1），白底太浅不可用
            "TABBG": "E3D2D4",     # 编号标签底：承深字 12.60:1
            "NUM": "A9505A",       # 编号圆：白字 20pt 粗体 5.27:1
            "TINT": "FAF1F2",      # 淡底卡片：承 BAND 字 7.65:1
            "BROWN": "8A6E52",     # 纯装饰，绝不承文字
            "INK": "1C1214",       # 标题：18.32:1
            "BODY": "2A2224",      # 正文：15.52:1
            "MUTED": "6B5B5E",     # 弱化文字：6.40:1
            "WHITE": "FFFFFF",
        },
    },
    "minimal-mono": {
        "name": "minimal-mono",
        "template": "assets/themes/minimal-mono.pptx",
        "background": "FFFFFF",
        "palette": {
            "BAND": "2F2F2F",      # 近黑灰标题带：承白字 13.39:1
            "BORDER": "3F3F3F",    # 描边与淡底上的强调字：白底 10.53:1
            "ACC": "B5532C",       # 全套唯一的强调色（陶土橙），白底 4.95:1
            "SKY": "EDEDED",       # 只放在 BAND 上（11.44:1）
            "TABBG": "DCDCDC",     # 编号标签底：承深字 13.77:1
            "NUM": "5F5F5F",       # 编号圆：白字 20pt 粗体 6.39:1
            "TINT": "F4F4F4",      # 淡底卡片：承 BAND 字 12.17:1
            "BROWN": "8C8C8C",     # 纯装饰（灰阶），白字仅 3.36:1，绝不承文字
            "INK": "111111",       # 标题：18.88:1
            "BODY": "262626",      # 正文：15.13:1
            "MUTED": "5A5A5A",     # 弱化文字：6.90:1
            "WHITE": "FFFFFF",
        },
    },
    # ── 新增三套 ──
    # midnight：深色底，第一个吃 `background` 字段的客户。整套色板做了**明度反转**：
    # INK/BODY/MUTED 变浅字、WHITE 键变成「面板底」（见 references/design-language.md §2c），
    # 所以 `panel()` 的默认填充必须是**调用时**才取的 WHITE，不能像原来那样在 def 时绑定
    # 成 "FFFFFF"（否则深底主题里卡片会整片变白，浅色文字直接糊掉）。
    "midnight": {
        "name": "midnight",
        "template": "assets/themes/midnight.pptx",
        "background": "101C2E",    # 深海军蓝页面底
        "palette": {
            "BAND": "9AC7EE",      # 浅蓝标题带/结论带：承白字 5.25:1，底色上 9.59:1
            "BORDER": "8FD3D0",    # 描边、headbar 底、淡底强调字：底色 10.10:1
            "ACC": "A8CDF0",       # 只做装饰线：底色 10.31:1
            "SKY": "304862",       # 只放在 BAND 上（5.28:1），底色上太暗不可用
            "TABBG": "34536F",     # 编号标签底：承深字 7.33:1
            "NUM": "7BA6CE",       # 编号圆：白字 20pt 粗体 3.65:1（大字，够）
            "TINT": "2A4560",      # 深色卡片：承 BAND 字 5.56:1 / MUTED 5.71:1
            "BROWN": "C9A86A",     # 纯装饰（暖金，深底上唯一暖色）：底色 7.57:1
            "INK": "F0F5FB",       # 标题：底色 15.61:1
            "BODY": "DCE6F2",      # 正文：底色 13.56:1
            "MUTED": "B2C7DE",     # 弱化文字：底色 9.87:1，面板上 5.39:1
            "WHITE": "31476B",     # 「面板底 / 深色块上的文字」：与 BAND 5.25:1
        },
    },
    "forest-green": {
        "name": "forest-green",
        "template": "assets/themes/forest-green.pptx",
        "background": "FFFFFF",
        "palette": {
            "BAND": "1B5E3A",      # 深绿标题带/结论带：承白字 7.75:1
            "BORDER": "2E7D4F",    # 描边、headbar 底：白底 5.05:1
            "ACC": "7A9A4E",       # 只做装饰线（橄榄）：白底 3.20:1，承不住小字
            "SKY": "DCEEE1",       # 只放在 BAND 上（6.41:1），白底太浅不可用
            "TABBG": "CFE0D5",     # 编号标签底：承深字 11.85:1
            "NUM": "3F8F5E",       # 编号圆：白字 20pt 粗体 3.96:1（大字，够）
            "TINT": "E8F2EC",      # 淡底卡片：承 BAND 字 6.77:1 / MUTED 4.94:1
            "BROWN": "6E5B3E",     # 纯装饰（土色，生态/材料的暖调），白底 6.50:1
            "INK": "10241A",       # 标题：白底 16.29:1
            "BODY": "22302A",      # 正文：白底 13.78:1
            "MUTED": "5A6B61",     # 弱化文字：白底 5.66:1
            "WHITE": "FFFFFF",
        },
    },
    "warm-sand": {
        "name": "warm-sand",
        "template": "assets/themes/warm-sand.pptx",
        "background": "FBF7F1",    # 很浅的暖白页面底（走 `background` 字段）
        "palette": {
            "BAND": "8A5A2B",      # 赭褐标题带/结论带：承白字 5.87:1
            "BORDER": "9A6633",    # 描边、headbar 底：暖底 4.55:1
            "ACC": "AE7C48",       # 只做装饰线：暖底 3.41:1，承不住小字
            "SKY": "F3E7D8",       # 只放在 BAND 上（4.82:1），暖底太浅不可用
            "TABBG": "E6D3BC",     # 编号标签底：承深字 11.71:1
            "NUM": "6B4A24",       # 编号圆：白字 20pt 粗体 7.98:1
            "TINT": "F6ECDD",      # 淡底卡片：承 BAND 字 5.02:1 / MUTED 5.28:1
            "BROWN": "A08562",     # 纯装饰（驼色），暖底 3.27:1，绝不承文字
            "INK": "241A10",       # 标题：暖底 15.99:1
            "BODY": "332A20",      # 正文：暖底 13.18:1
            "MUTED": "6E5F4C",     # 弱化文字：暖底 5.78:1
            "WHITE": "FFFFFF",
        },
    },
}

DEFAULT_THEME = "academic-blue"


def theme_dir() -> Path:
    return PKG / "assets" / "themes"


def theme_json(name: str) -> Path:
    return theme_dir() / f"{name}.json"

# ── 版面常量（会被 apply_theme 覆写）──
M = 0.66
CW = 12.01
FTY = 6.96
FOOT_TOP = 6.98
PAGE_W, PAGE_H = 13.33, 7.5
CT_TOP, CT_BOT = 1.66, 6.94

# ── 配色（会被 apply_theme 覆写；这些初值只是 import 期能读，之后一律以 theme 为准）──
BAND = "006DB8"
BORDER = "4472C4"
ACC = "6096E6"
SKY = "E8F2FB"
TABBG = "D3D9E9"
NUM_FILL = "7B8DBE"   # 色名 NUM「编号圆」与字体 NUMF「数字字体」刻意区分
TINT = "F2F6FB"
BROWN = "B39B77"
INK = "0F1423"
BODY = "262626"
MUTED = "56646E"
WHITE = "FFFFFF"

# 页面底色。**导出给闸门三（qa_visual.py）用**：它拿这个当「矢量填充查不到时的
# 回落底色」，深色主题不改这里就会把蓝底白字判成白底白字、全篇误报。
# 初值必须和 NEUTRAL_THEME["background"] 一致；import 后即可读，之后以 apply_theme 为准。
PAGE_BG = "FFFFFF"

# ── 字体（会被 apply_theme 覆写）──
LAT = "Times New Roman"
EAK = "微软雅黑"
NUMF = "Times New Roman"
FORM = "Times New Roman"

# ── 版式角色 ──
LCOVER = "标题幻灯片"
LCONTENT = "标题和内容"
LTWOCOL = "两栏内容"
LSECTION = "节标题"      # 章节分隔页：内置母版有这个版式但以前没人映射，见 §2c
LCLOSE = "末尾幻灯片"

TPL = ""            # 基座模板路径，apply_theme 解析
LAYOUTS = dict(NEUTRAL_THEME["layouts"])

PRS = None
PAGE = 0
MAN = []            # 版面清单，save_deck 落盘，供闸门一消费
OUT = None
TOTAL = 0
FOOT_BAR = ""       # 页脚左侧那一句

# 讲稿的唯一来源：{页码: (标题, 计时, 台词)}。
# notes(pn=…) 与 write_script_md() 都从它读，页面脚本只往这里填一处。
SCRIPT: dict[int, tuple] = {}

_ACTIVE_THEME: dict = {}     # 最近一次 apply_theme 归一化后的结果；kit_config 据此只换 template


def apply_theme(theme: dict) -> dict:
    """把一份 theme dict 写进模块级常量。返回归一化后的 theme（供调用方存档）。

    之所以「写模块全局」而不是到处传参：页面代码写起来要短，
    而改动主题的人（模型自己）最容易漏掉深层传参。全局是故意的，改 theme.json 即可。
    """
    global M, CW, FTY, FOOT_TOP, PAGE_W, PAGE_H, CT_TOP, CT_BOT, R
    global BAND, BORDER, ACC, SKY, TABBG, TINT, BROWN, INK, BODY, MUTED, WHITE
    global NUM_FILL, PAGE_BG, LAT, EAK, NUMF, FORM, TPL, LAYOUTS, R
    global LCOVER, LCONTENT, LTWOCOL, LSECTION, LCLOSE
    global _ACTIVE_THEME

    pal = dict(NEUTRAL_THEME["palette"])
    pal.update(theme.get("palette") or {})
    for k in ("BAND", "BORDER", "ACC", "SKY", "TABBG", "TINT", "BROWN",
              "INK", "BODY", "MUTED", "WHITE"):
        globals()[k] = pal[k]
    # 「编号圆」在设计语言里叫 NUM（色名），和字体的 NUMF（数字字体）区分开
    NUM_FILL = pal["NUM"]

    # 页面底色：旧 theme.json 没有 background 键时回落到中性主题的 FFFFFF，
    # 于是老主题文件的行为一字不变；新主题给了这个键才走深色底。
    background = (theme.get("background") or dict(NEUTRAL_THEME).get("background")
                  or "FFFFFF").upper()
    PAGE_BG = background

    f = dict(NEUTRAL_THEME["fonts"])
    f.update(theme.get("fonts") or {})
    LAT, EAK, NUMF, FORM = f["latin"], f["ea"], f["num"], f["formula"]

    g = dict(NEUTRAL_THEME["grid"])
    g.update(theme.get("grid") or {})
    M, CW = g["M"], g["CW"]
    FTY, FOOT_TOP = g["FTY"], g["FOOT_TOP"]
    PAGE_W, PAGE_H = g["W"], g["H"]
    CT_TOP, CT_BOT = g["CT_TOP"], g["CT_BOT"]
    R = M + CW

    LAYOUTS = dict(NEUTRAL_THEME["layouts"])
    LAYOUTS.update(theme.get("layouts") or {})
    LCOVER = LAYOUTS["cover"]
    LCONTENT = LAYOUTS["content"]
    LTWOCOL = LAYOUTS.get("two_col", LAYOUTS["content"])
    # section 是唯一可有可无的角色：老模板没有「节标题」版式时回退到 content，
    # 与 two_col 的降级路径一致（check_template 会打一行 WARN）。
    LSECTION = LAYOUTS.get("section", LAYOUTS["content"])
    LCLOSE = LAYOUTS["close"]

    tpl = theme.get("template") or NEUTRAL_THEME["template"]
    p = Path(tpl)
    TPL = str(p if p.is_absolute() else (PKG / p))
    _ACTIVE_THEME = {"name": theme.get("name", "theme"), "template": TPL,
                     "background": PAGE_BG, "palette": pal, "fonts": f,
                     "grid": g, "layouts": LAYOUTS}
    return dict(_ACTIVE_THEME)


def load_theme(path=None, name=None) -> dict:
    """套用主题。返回实际生效的 theme dict。

    三种写法，优先级从高到低：
      · `load_theme("theme.json")`        —— 用户/模板反推出来的文件（老接口，不变）
      · `load_theme("assets/themes/scholar-red.json")` —— 显式文件
      · `load_theme(name="scholar-red")` —— 直接切内置主题，模板自动指到
                                            `assets/themes/<name>.pptx`

    `path` 传进来的如果是内置主题名（BUILTIN_THEMES 里的键）也按 name 处理，
    所以 `--theme academic-blue` 和 `--theme assets/themes/academic-blue.json`
    两种写法都成立。文件不存在且不是内置名 → 回落到内置中性主题（不算错误）。
    """
    p = Path(path) if path else None
    if p is not None and str(p) in BUILTIN_THEMES:
        name, p = str(p), None
    if name:
        if name not in BUILTIN_THEMES:
            raise KeyError(f"没有内置主题「{name}」。可选：{sorted(BUILTIN_THEMES)}")
        t = json.loads(theme_json(name).read_text(encoding="utf-8")) \
            if theme_json(name).exists() else BUILTIN_THEMES[name]
        return apply_theme(t)
    if p is not None and p.exists():
        return apply_theme(json.loads(p.read_text(encoding="utf-8")))
    if p is not None:
        # 给了一个不存在的文件：按内置名再试一次，否则回落
        raise FileNotFoundError(
            f"主题文件不存在：{p}\n"
            f"  · 内置主题名：{sorted(BUILTIN_THEMES)}\n"
            f"  · 配套 json：{theme_dir() / (str(p) + '.json')}")
    return apply_theme({"name": "theme-neutral"})


def check_template() -> str:
    """确认基座模板在，且要求的版式都在。缺版式要明确报错，不静默降级。

    two_col 与 section 是仅有的两个允许降级的角色（页面可选），但**降级必须打一行告警**：
    模板没有对应版式却照跑，构图会整块塌掉，且闸门查不出来。
    """
    global LAYOUTS, LTWOCOL, LSECTION
    if not Path(TPL).exists():
        raise FileNotFoundError(
            f"基座模板不存在：{TPL}\n"
            f"  · 用内置中性母版：确认 {PKG / 'assets' / 'theme-neutral.pptx'} 存在，"
            f"或删掉 theme.json 里的 template 字段走默认值。\n"
            f"  · 用自己的模板：把路径写进 theme.json 的 template，"
            f"或调用 kit.kit_config(template=...)。")
    prs = Presentation(TPL)
    names = [l.name for l in prs.slide_masters[0].slide_layouts]
    need = {LCOVER, LCONTENT, LCLOSE}
    miss = need - set(names)
    if miss:
        raise KeyError(
            f"模板 {Path(TPL).name} 缺少版式：{sorted(miss)}。\n"
            f"  模板现有版式：{names}\n"
            f"  处理：① 打开模板把对应版式改名成上面这些名字；"
            f"或 ② 改 theme.json 的 layouts 字段指到现有版式名。\n"
            f"  注意：`--template` / `kit_config(template=...)` **只换基座文件**，"
            f"不碰 layouts——版式名只能靠 theme.json 指，否则这条建议在换基座的路径下无效。\n"
            f"  不做静默降级——版式错了，版式自带的装饰（标题带、封面装饰）会整片丢失。")
    if LTWOCOL not in names:
        print(f"  [WARN] 模板 {Path(TPL).name} 没有两栏版式「{LTWOCOL}」"
              f"（共 {len(names)} 个版式）——two_col 页面回退到「{LCONTENT}」。"
              f"要真两栏请在 theme.json 的 layouts.two_col 里指定。")
        LAYOUTS["two_col"] = LCONTENT
        LTWOCOL = LCONTENT
    if LSECTION not in names:
        print(f"  [WARN] 模板 {Path(TPL).name} 没有章节版式「{LSECTION}」"
              f"——section 页面回退到「{LCONTENT}」（会带着正文页的标题带）。"
              f"要真分隔页请在 theme.json 的 layouts.section 里指定。")
        LAYOUTS["section"] = LCONTENT
        LSECTION = LCONTENT
    return TPL


def kit_config(template=None, asset_dir=None, out_dir=None):
    """直接配置入口：**只覆盖 `template` 字段**（换个基座文件），其余全保留。

    保留已载入 theme 的 palette / fonts / grid / layouts——这是 F1 修掉的那个坑：
    旧实现重新调 `apply_theme({"name","template"})`，把 theme.json 的东西全冲回内置中性默认，
    于是 `check_template()` 报「缺版式」，而它给的建议（改 theme.json 的 layouts）在
    只换基座的路径下 provably 无效。
    """
    global ASSET, OUT_DIR
    if asset_dir:
        ASSET = Path(asset_dir)
    if out_dir:
        OUT_DIR = Path(out_dir)
    t = {k: v for k, v in (_ACTIVE_THEME or {}).items()
         if k in ("palette", "fonts", "grid", "layouts")}
    t["name"] = (_ACTIVE_THEME or {}).get("name", "theme")
    if template:
        t["template"] = template
        t["name"] = f"{t['name']}@template:{Path(template).stem}"
    return apply_theme(t)


# ── 元数据擦除 ──
# 发到 GitHub / 交付给别人的 pptx 里，docProps 是最容易被忽略的一处泄露面：
# 第三方库自带的模板会把作者写成它自己的作者、把 created/modified 写成
# Office 模板的 2013 年时间戳、还会在 app.xml 里留下一条 Mac 版 PowerPoint 的应用名
# 和一个写错的 4:3 标注（实际是 16:9 / 13.333in）。
# 这些都不是我们生成的，凭什么替文件署名。
#
# 所以落盘前统一擦一次：作者换成工具名、日期换成固定占位。
# 固定占位而不是「现在」——真实时间戳本身也是信息（谁在什么时候生成的这一稿）。
AUTHOR = "research-ppt-builder"
FIXED_TS = datetime(2000, 1, 1, 0, 0, 0)     # 占位时间戳，不是真实时间
DOC_TITLE = "research-ppt-builder deck"
_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


def scrub_metadata(prs, title=DOC_TITLE, author=AUTHOR):
    """把 docProps 擦成中性值。母版生成与产物保存**都走它**。

    core.xml：creator / lastModifiedBy = 工具名，title = 中性标题，revision = 1，
              created / modified / lastPrinted = 固定占位时间戳（不留真实时间），
              subject / keywords / description / category 等一律清空。
    app.xml ：Application = 工具名（不留 python-pptx / Microsoft 字样），
              PresentationFormat 改成真实的 16:9，删掉 Template / Company / Manager。

    为什么单独抽出来：母版文件和产物文件是两条落盘路径，早期只在产物侧擦过，
    母版里那份第三方署名一直跟着分发。只有一个入口才不会漏。
    """
    cp = prs.core_properties
    cp.title = title
    cp.author = author
    cp.last_modified_by = author
    cp.revision = 1
    cp.created = FIXED_TS
    cp.modified = FIXED_TS
    cp.last_printed = FIXED_TS
    for attr in ("subject", "keywords", "comments", "category",
                 "content_status", "identifier", "language", "version"):
        try:
            setattr(cp, attr, "")
        except Exception:          # 老版本 python-pptx 少几个属性，不该因此崩
            pass

    for part in prs.part.package.iter_parts():
        if str(part.partname) != "/docProps/app.xml":
            continue
        root = etree.fromstring(part.blob)
        def _set(tag, val):
            el = root.find(f"{{{_EP}}}{tag}")
            if el is None:
                el = etree.SubElement(root, f"{{{_EP}}}{tag}")
            el.text = val
        _set("Application", author)
        _set("PresentationFormat", "On-screen Show (16:9)")
        _set("TotalTime", "0")
        _set("AppVersion", AUTHOR)
        for tag in ("Template", "Company", "Manager"):
            el = root.find(f"{{{_EP}}}{tag}")
            if el is not None:
                root.remove(el)
        part._blob = etree.tostring(root, xml_declaration=True,
                                    encoding="UTF-8", standalone=True)
        break
    return dict(title=title, author=author, ts=FIXED_TS.isoformat() + "Z")


# ── 图片比例 ──
_AR_CACHE: dict[str, float] = {}


def ar(path) -> float | None:
    """图片原始宽高比（w/h）。用 pillow 实测；pillow 不可用返回 None。

    为什么要实测而不是让调用者手填：手填的常数一旦对不上，
    闸门一（qa_fit）与闸门二（qa_pdf）会同时把正常图判成「拉伸」。
    """
    p = str(path)
    if p in _AR_CACHE:
        return _AR_CACHE[p]
    try:
        from PIL import Image
        with Image.open(p) as im:
            v = im.width / im.height
    except Exception:
        v = None
    _AR_CACHE[p] = v
    return v


# ── deck 生命周期 ──
def open_base():
    """以基座模板为底座打开，清空模板自带的示例页。

    「以模板文件为基座打开」而不是新建演示文稿：母版和版式要原封不动地留在包里。
    删幻灯片必须做 XML 手术——从 `prs.slides._sldIdLst` 移除节点之后再
    `prs.part.drop_rel(rId)` 断关系，否则包里的 part 会变成孤儿。
    """
    global PRS
    check_template()
    PRS = Presentation(TPL)
    lst = PRS.slides._sldIdLst
    for sld in list(lst):
        PRS.part.drop_rel(sld.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
        lst.remove(sld)
    return PRS


def start_deck(deck: dict):
    """开启一份新产物：重置版面清单、总页数、页脚口号。

    总页数不是常量而是每场设一次——一份 deck 单独放映时页码必须自洽，
    页脚分母按本场 total 走。deck 需要：out / total / 可选 manifest / foot / title。
    """
    global OUT, TOTAL, MAN, FOOT_BAR
    OUT = OUT_DIR / deck["out"]
    TOTAL = deck["total"]
    MAN = []
    FOOT_BAR = deck.get("foot", "")


def save_deck(prs, deck: dict):
    """落盘 pptx + 版面清单 JSON；页数不符直接断言失败。

    版面清单（text_layout_*.json）是闸门一的唯一输入：它记着每个文本框的真实坐标
    与逐 run 字号。渲染后的 PDF 查不出溢出（LibreOffice 静默裁切），
    只能在源码几何层面查，所以这个文件是必须的，不是可选的调试产物。
    """
    OUT.parent.mkdir(parents=True, exist_ok=True)
    scrub_metadata(prs, title=deck.get("title") or DOC_TITLE)   # 落盘前擦 docProps
    prs.save(OUT)
    manifest = OUT.parent / deck.get("manifest", OUT.stem + "_layout.json")
    manifest.write_text(json.dumps(MAN, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len(prs.slides._sldIdLst)
    assert n == deck["total"], f"{deck['out']}：页数 {n} 与预期 {deck['total']} 不符"
    txt = sum(1 for d in MAN if d["kind"] == "text")
    img = sum(1 for d in MAN if d["kind"] == "img")
    print(f"{n} 页 → {OUT.name}")
    print(f"    版面清单：{txt} 个文字块 / {img} 张图 → {manifest.name}")
    return manifest


# ── 页面原语 ──
_ROLE_ALIAS = {"cover": "cover", "content": "content", "two_col": "two_col",
               "section": "section", "close": "close"}


def resolve_layout(name_or_role: str) -> str:
    """把 'content' / '标题和内容' 之类的写法解析成模板里真实存在的版式名。"""
    if name_or_role in _ROLE_ALIAS:
        return LAYOUTS[_ROLE_ALIAS[name_or_role]]
    return name_or_role


def new_slide(layout, page: int):
    """新增一页；抹掉版式克隆来的占位符，只留下版式自带的装饰。

    装饰（标题带、右上色块、封面几何块）不是占位符，会被 LibreOffice / PowerPoint
    自动画在幻灯片下层，不需要（也不应该）在页面函数里重画一遍。
    """
    global PAGE
    PAGE = page
    want = resolve_layout(layout)
    lays = {l.name: l for l in PRS.slide_masters[0].slide_layouts}
    if want not in lays:
        raise KeyError(f"模板里没有版式「{want}」。现有版式：{sorted(lays)}")
    s = PRS.slides.add_slide(lays[want])
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    return s


def notes(s, text: str | None = None, pn: int | None = None):
    """写入演讲者备注。放映不显示，排练与现场问答用；不占版面，不参与闸门。

    第一次访问 `notes_slide` 会自动创建备注页，不需要手工 add。

    讲稿的唯一来源是 `SCRIPT`：传 `pn` 就按页码从 `SCRIPT[pn]` 取台词，
    不必在页面脚本里再抄一份——两处同源才是「改页面就要改讲稿」能成立的前提。
    """
    if text is None and pn is not None:
        entry = SCRIPT.get(pn)
        text = entry[2] if entry else ""
    if text:
        s.notes_slide.notes_text_frame.text = text
    return s


# ── 底层样式 ──
def _style(r, size, bold, color, italic, latin, ea, spc):
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = RGBColor.from_string(color)
    f.name = latin                       # 只写 <a:latin>
    rPr = r._r.get_or_add_rPr()
    el = rPr.find(qn("a:ea"))
    if el is None:
        el = etree.Element(qn("a:ea"))
        lat = rPr.find(qn("a:latin"))
        if lat is not None:
            lat.addnext(el)              # 关键：<a:ea> 必须排在 <a:latin> 之后
        else:
            rPr.append(el)
    el.set("typeface", ea)
    if spc is not None:
        rPr.set("spc", str(int(round(spc * 100))))


def txt(s, x, y, w, h, runs, size=14, bold=False, color=None, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, ls=1.30, italic=False, latin=None, ea=None, spc=None,
        role=None):
    """文本框。runs 为 str 或 [(text, size|None, bold|None, color|None), ...]；text 内 \\n 断段。

    `latin` / `ea` 一律 None 默认、调用时才取当前值——见模块 docstring 约束 2。
    `role` 只进版面清单（目前 footer / page_no），让闸门能识别页脚。
    版面清单里每个 run 记 [文本, 字号, 是否粗体]，闸门一据此按真实字宽逐字累加。
    """
    color = color or BODY
    latin = latin or LAT
    ea = ea or EAK
    box = s.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    rec = []
    if isinstance(runs, str):
        runs = [(runs, None, None, None)]
    par = tf.paragraphs[0]
    par.alignment = align
    par.line_spacing = ls
    for item in runs:
        text = str(item[0])
        sz = item[1] if len(item) > 1 else None
        bd = item[2] if len(item) > 2 else None
        cl = item[3] if len(item) > 3 else None
        # 同一段内的多个 run 必须共用一个段落，只有 \n 才换段。
        # 写成「一个 run 一个新段落」会把 "图 4　＋说明" 这种双 run 标签拆成两行。
        for i, seg in enumerate(text.split("\n")):
            if i > 0:
                par = tf.add_paragraph()
                par.alignment = align
                par.line_spacing = ls
            if not seg:
                continue
            r = par.add_run()
            r.text = seg
            _style(r, sz or size, bd if bd is not None else bold,
                   cl or color, italic, latin, ea, spc)
            rec.append([seg, sz or size, bool(bd if bd is not None else bold)])
    d = dict(p=PAGE, kind="text", x=x, y=y, w=w, h=h, runs=rec, ls=ls, latin=latin)
    if role:
        d["role"] = role
    MAN.append(d)
    return box


def rect(s, x, y, w, h, fill=None, line=None, lw=2.25, alpha=None,
         shape=MSO_SHAPE.RECTANGLE):
    """填充矩形 / 任意自选图形。`alpha` 0–1，写成 <a:alpha val="千分比"/>。"""
    fill = WHITE if fill is None and line is None else fill
    sh = s.shapes.add_shape(shape, In(x), In(y), In(w), In(h))
    sh.shadow.inherit = False
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
        if alpha is not None:
            clr = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            etree.SubElement(clr, qn("a:alpha")).set("val", str(int(round(alpha * 1000))))
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = RGBColor.from_string(line)
        sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    MAN.append(dict(p=PAGE, kind="rect", x=x, y=y, w=w, h=h))
    return sh


def panel(s, x, y, w, h, fill=None, line=None, lw=2.25):
    """内容框：白底 + 描边。line=None 时取主题的 BORDER。

    `fill` 的默认值必须是 None 而不是 WHITE：**默认参数在 def 时就绑定**，
    绑成 "FFFFFF" 的话 apply_theme() 换了主题也不生效——midnight 这类深色主题
    下所有卡片会整片变白，浅色正文直接糊在上面。调用时再取 WHITE 才跟着主题走。
    """
    return rect(s, x, y, w, h, fill=fill or WHITE, line=line or BORDER, lw=lw)


def rule(s, x, y, w, color=None, h=0.016):
    return rect(s, x, y, w, h, fill=color or BORDER, line=None)


def _resolve_asset(name) -> str:
    p = Path(name)
    return str(p if p.exists() else (ASSET / p.name))


def _img_path(name) -> str:
    """解析图片路径（原路径 → ASSET/），**不存在就报错**。

    为什么要单独一步：`ar()` 读不出比例时只有两种原因（文件不存在 / pillow 缺失或文件损坏），
    都报同一句「读不出宽高比」的话，用户会去装 pillow——而真实原因常常是文件名打错了。
    这里先把「文件不存在」摘出去单独报。
    """
    p = _resolve_asset(name)
    if not Path(p).exists():
        raise FileNotFoundError(
            f"文件不存在：{p}\n"
            f"  · 传入的名字：{name}\n"
            f"  · kit.ASSET = {ASSET}（改它指向图片目录，或直接传绝对路径）\n"
            f"  · 这不是比例问题：比例问题由 ar() 读得出/读不出区分，不会走到这条报错。")
    return p


def pic(s, x, y, w, h, name, sized=False, role=None):
    """插图。`w`/`h` 必须等比（用 fig() / fig_fit()，别直接算），
    `ar` 与原图比例一并写进版面清单供闸门核对。"""
    p = _img_path(name)
    shp = s.shapes.add_picture(p, In(x), In(y), In(w), In(h))
    d = dict(p=PAGE, kind="img", x=x, y=y, w=w, h=h, path=p, sized=sized,
             ar=ar(p))
    if role:
        d["role"] = role
    MAN.append(d)
    return shp


def fig(s, x, y, w, name, no, text, gap=0.07, cap_h=0.30):
    """插图 + 图注（图注与图之间统一 gap，图注高统一 cap_h）。返回图注底部 y。

    图与正文的间距是常量不是随手填的：同一份 deck 里图注位置忽高忽低，
    翻页时会明显看出不齐。
    """
    p = _img_path(name)                  # 先定位：路径不通就别谈比例
    a = ar(p)
    if not a:
        raise ValueError(f"读不出 {p} 的宽高比（pillow 缺失或文件损坏）")
    h = w / a
    pic(s, x, y, w, h, p)
    cap(s, x, y + h + gap, w, no, text, h=cap_h)
    return y + h + gap + cap_h


def fig_fit(s, x, y, w, slot_h, name):
    """在 w × slot_h 槽内按原图等比缩放并双向居中；返回槽底 y。

    等比 + 居中是硬要求：拉伸变形在投影上一眼就能看出来，而且闸门二会直接报。
    """
    p = _img_path(name)
    a = ar(p)
    if not a:
        raise ValueError(f"读不出 {p} 的宽高比（pillow 缺失或文件损坏）")
    fh = w / a
    if fh > slot_h:
        fh = slot_h
    fw = fh * a
    pic(s, x + (w - fw) / 2, y + (slot_h - fh) / 2, fw, fh, p)
    return y + slot_h


def cap(s, x, y, w, no, text, h=0.30):
    """图注：序号用 BAND 粗体，说明用 MUTED，两个 run 共用一个段落。"""
    if no:
        txt(s, x, y, w, h, [(f"图 {no}　", 11.5, True, BAND), (text, 11.5, None, MUTED)],
            ls=1.26, anchor=MSO_ANCHOR.MIDDLE)
    else:
        txt(s, x, y, w, h, [(text, 11.5, None, MUTED)], ls=1.26, anchor=MSO_ANCHOR.MIDDLE)


def note(s, x, y, w, text, size=11, h=0.30, bold=False, color=None, ls=1.28):
    txt(s, x, y, w, h, [(text, size, bold)], color=color or MUTED, ls=ls,
        anchor=MSO_ANCHOR.MIDDLE)


def head(s, label, title, size=30):
    """蓝色标题带内的白色主标题 + 带下方的章节标签。

    标题带本身是版式自带的装饰（y 0.41–1.24），这里只写字、不画带。
    """
    txt(s, M, 0.45, 10.30, 0.74, [(title, size, True)], color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if label:
        txt(s, M, 1.31, 9.2, 0.30, [(label, 12, True)], color=BAND, spc=3,
            anchor=MSO_ANCHOR.MIDDLE)


def foot(s, n):
    """页脚：左侧一句口号，右侧 n / TOTAL。两条都在页脚带内，不参与压页脚判定。"""
    if FOOT_BAR:
        txt(s, M, FTY, 7.0, 0.26, [(FOOT_BAR, 10.5)], color=MUTED,
            anchor=MSO_ANCHOR.MIDDLE, role="footer")
    txt(s, R - 1.10, FTY, 1.10, 0.26, [(f"{n} / {TOTAL}", 10.5)], color=MUTED,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, role="page_no")


def headbar(s, x, y, w, text, h=0.46, size=15.5, fill=None, color=None, xpad=0.26):
    """卡片顶部的深色标题条。`color` 默认取 WHITE（深色块上的文字），调用时才取。"""
    fill = fill or BORDER
    rect(s, x, y, w, h, fill=fill, line=None)
    txt(s, x + xpad, y, w - 2 * xpad, h, [(text, size, True)], color=color or WHITE,
        anchor=MSO_ANCHOR.MIDDLE)


def chip(s, x, y, w, h, no, label, size=16, num_size=20):
    """编号标签：TABBG 底 + NUM 色圆 + 圆内白色大序号 + 右侧深色标签。"""
    rect(s, x, y, w, h, fill=TABBG, line=None)
    rect(s, x, y, h, h, fill=NUM_FILL, line=None, shape=MSO_SHAPE.OVAL)
    txt(s, x, y, h, h, [(no, num_size, True)], color=WHITE, latin=NUMF,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + h + 0.16, y, w - h - 0.30, h, [(label, size, True)], color=INK,
        anchor=MSO_ANCHOR.MIDDLE)


# ── 表格 ──
def _cell_ln(c, color=None, w=9525):
    """写单元格四边框。

    两条约束：① <a:lnL/lnR/lnT/lnB> 必须排在 <a:tcPr> 最前且互按此序；
    ② 单元格描边要在**设填充之前**写——OOXML 要求边框元素排在 solidFill 前面，
    顺序非法的话 PowerPoint 会忽略这段边框。所以 table() 里先 _cell_ln 再 fill。

    `color` 默认 None → 取 WHITE（浅色主题里就是 FFFFFF，行为不变）。
    """
    color = color or WHITE
    tcPr = c._tc.get_or_add_tcPr()
    for tag in ("a:lnBlToTr", "a:lnTlToBr", "a:lnB", "a:lnT", "a:lnR", "a:lnL"):
        for o in tcPr.findall(qn(tag)):
            tcPr.remove(o)
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):
        ln = etree.Element(qn(tag))
        ln.set("w", str(w))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        f = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        etree.SubElement(f, qn("a:srgbClr")).set("val", color)
        tcPr.insert(0, ln)


def table(s, x, y, colW, rows, rowH=0.50, hdr_fill=None, hdr_color=None,
          hdr_size=12.5, body_size=13, zebra=None):
    """rows[i][j] = 文本 或 (文本, size|None, bold|None, color, align)。

    `hdr_color` 默认 None → 取 WHITE（调用时才取，理由同 `panel()` 的 `fill`）。
    """
    hdr_fill = hdr_fill or BORDER
    hdr_color = hdr_color or WHITE
    zebra = zebra or TINT
    nrow, ncol = len(rows), len(colW)
    g = s.shapes.add_table(nrow, ncol, In(x), In(y), In(sum(colW)), In(rowH * nrow))
    tb = g.table
    tb.first_row = False
    tb.horz_banding = False
    tb.vert_banding = False
    for j, cw_ in enumerate(colW):
        tb.columns[j].width = In(cw_)
    for i in range(nrow):
        tb.rows[i].height = In(rowH)
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            text = cell if isinstance(cell, str) else cell[0]
            size = hdr_size if i == 0 else body_size
            bold = i == 0
            color = hdr_color if i == 0 else BODY
            align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
            if not isinstance(cell, str):
                if len(cell) > 1 and cell[1]:
                    size = cell[1]
                if len(cell) > 2 and cell[2] is not None:
                    bold = cell[2]
                if len(cell) > 3:
                    color = cell[3]
                if len(cell) > 4:
                    align = cell[4]
            c = tb.cell(i, j)
            _cell_ln(c)                      # 必须在 fill 之前
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(
                hdr_fill if i == 0 else (zebra if i % 2 == 0 else WHITE))
            c.margin_left = In(0.09)
            c.margin_right = In(0.09)
            c.margin_top = In(0.03)
            c.margin_bottom = In(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            par = tf.paragraphs[0]
            par.alignment = align
            par.line_spacing = 1.2
            r = par.add_run()
            r.text = text
            _style(r, size, bold, color, False, LAT, EAK, None)
    MAN.append(dict(p=PAGE, kind="table", x=x, y=y, rows=nrow, rowH=rowH))
    return tb


# ── 图表 ──
# 原生图表（add_chart）优先：可编辑、体积只有几十 KB、数值跟着数据走，
# 也不会触发图片拉伸那两道闸门。只有样式控制不住时才退回 matplotlib 出图，
# 那条路的完整技法见 references/figures-and-charts.md 与 scripts/figstyle.py。
#
# 两条内容纪律，代码层面各有一个参数对应：
#   · `values` 里可以放 None = **未参评**。绝不能把它当 0 画成柱——
#     0 柱会被读成「测出来是 0」，未参评和 0 是两件完全不同的事。
#   · `log_axis=True` 处理量级差。默认线性轴在差两个数量级时会把小的那几根
#     压成一条线，读者会以为没数据。
MISSING_TEXT = "未参评"      # 缺失值标注文字（灰色斜体，由 missing_marks 画）
# 标注颜色**不设模块级常量**：它跟随主题的 MUTED，而模块级常量在 import 期求值，
# 那时的 MUTED 还没被 theme.json 覆写，写死会拿到旧色（和 EAK 那条坑同源）。
# 要改色就调 missing_marks(color=…)。

# 新参数速查（详见各函数 docstring 与 references/figures-and-charts.md）：
#   column_chart(..., log_axis=False, missing=MISSING_TEXT, headroom=0.08,
#                dl_size=16, lbl_size=14, num_fmt="0.000", dl_color=None)
#   hbar_chart  (..., log_axis=False, missing=MISSING_TEXT, headroom=0.08,
#                reverse=True, ...)，labels 传长类别名，竖排不挤
#   value_labels(plot_or_series, ...)  统一数据标签样式，缺数值的位置画成灰斜体
GAP_WIDTH = 62              # 类间距（%）。62 ≈ 柱宽 1.6 倍，柱子不粘连、也不散
HEADROOM = 0.08             # 顶部留白比例：OUTSIDE_END 标签必须落在轴上限之外
LOG_HEADROOM = 1.6          # 对数轴的顶部留白：对数是乘性空间，只能放大不能加常数


def _hide_axis(axis):
    """彻底隐藏坐标轴：写 <c:delete val="1"/>，且必须排在 <c:scaling> 之后。"""
    el = getattr(axis, "_element", None)
    if el is None:
        return
    for o in el.findall(qn("c:delete")):
        el.remove(o)
    d = etree.Element(qn("c:delete"))
    d.set("val", "1")
    sc = el.find(qn("c:scaling"))
    if sc is not None:
        sc.addnext(d)
    else:
        el.insert(1 if el.find(qn("c:axId")) is not None else 0, d)


def _set_log_axis(axis, on=True):
    """把数值轴切成对数刻度：写 <c:logBase val="10"/>。

    `<c:logBase>` 必须是 `<c:scaling>` 的**第一个子元素**（CT_Scaling 的
    顺序是 logBase → orientation → max → min），插错位置 PowerPoint 会整段忽略。
    对数轴要求所有值 > 0，0 和负数都画不出来——调用方要自己先筛。
    """
    sc = axis._element.find(qn("c:scaling"))
    if sc is None:
        return False
    for o in sc.findall(qn("c:logBase")):
        sc.remove(o)
    if not on:
        return True
    lb = etree.Element(qn("c:logBase"))
    lb.set("val", "10")
    sc.insert(0, lb)
    return True


def _reverse_categories(axis):
    """横向柱状图默认把第一个类别画在最下面，读图顺序是反的；翻转成从上往下。"""
    sc = axis._element.find(qn("c:scaling"))
    if sc is None:
        return
    o = sc.find(qn("c:orientation"))
    if o is None:
        o = etree.SubElement(sc, qn("c:orientation"))
    o.set("val", "maxMin")


# 缺失值标注的槽位几何常数（单位：英寸 / 比例）。
# 这两个数不是拍脑袋来的，是从 LibreOffice 渲染出来的 PDF 里量出来的：
# 画布左右各内缩 2%，类目标签下方固定 0.24in + 一个行高。改这里之前先量一遍渲染结果。
PLOT_INSET_X = 0.02          # 每侧，占图表宽的比例
CAT_LABEL_PAD = 0.24         # 类目标签带下方的固定留白
MARK_SIZE = 11.5             # 「未参评」标注字号（比真值标签小一号：它是弱信息）


def _is_cjk(ch: str) -> bool:
    o = ord(ch)
    return 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF


def _text_w_in(text: str, size_pt: float) -> float:
    """估算一行文字的宽（英寸）。CJK 按 1.0em、西文按 0.5em。

    只用于给缺失值标注找位置，不需要 qa_fit 那套实测字宽——差 0.05in 在
    「未参评」这三个字上根本看不出来，而实测字宽要读字体文件，太重。
    """
    em = sum(1.0 if _is_cjk(c) else 0.5 for c in text) * size_pt / 72.0
    return em


def missing_marks(s, x, y, w, h, n_cats, missing_idx, labels=None, vertical=True,
                  lbl_size=14, text=MISSING_TEXT, color=None):
    """在缺失值的槽位上放灰色斜体「未参评」，返回这些文本框。

    **为什么是页面上的文本框而不是图表里的自定义数据标签**：写在 `<c:dLbl>` 里的
    自定义标签文本在 PowerPoint 里能显示，但 LibreOffice 渲染时，只要系列里有一个
    空值就会**把整组数据标签丢掉**——连有值的那几格的数字一起消失。也就是说
    「用图表自己的标签机制标缺失值」这条路的渲染结果是：PowerPoint 里对、
    闸门渲染出来的那份里错，而错的那份恰恰是验收依据。所以这里老老实实放文本框，
    两边渲染一致，闸门也能量到它（走 `kit.txt`，进版面清单）。

    代价是它属于页面而不属于图表：移动图表时要连它一起移动。
    """
    out = []
    if not missing_idx or not n_cats:
        return out
    color = color or MUTED
    for i in missing_idx:
        if vertical:      # 柱状图：槽在 x 方向，标注贴在轴线上方
            plot_x0 = x + PLOT_INSET_X * w
            plot_w = w * (1 - 2 * PLOT_INSET_X)
            cx = plot_x0 + (i + 0.5) * plot_w / n_cats
            base = y + h - (CAT_LABEL_PAD + lbl_size / 72.0)
            out.append(txt(s, cx - 0.65, base - 0.26, 1.30, 0.26, [(text, MARK_SIZE, True)],
                           color=color, italic=True, align=PP_ALIGN.CENTER,
                           anchor=MSO_ANCHOR.BOTTOM))
        else:             # 条形图：槽在 y 方向，标注贴在类目右侧
            lab_w = max((_text_w_in(str(t), lbl_size) for t in (labels or [])), default=0.0)
            cy = y + h * (i + 0.5) / n_cats
            out.append(txt(s, x + lab_w + 0.16, cy - 0.15, 1.40, 0.30, [(text, MARK_SIZE, True)],
                           color=color, italic=True, anchor=MSO_ANCHOR.MIDDLE))
    return out


def value_labels(target, dl_size=16, lbl_color=None, num_fmt="0.000", pos=None):
    """统一数据标签样式（字号 / 字色 / 数字格式 / 外端位置）。

    `target` 可以是 `plot`（整个系列一张标签表）也可以是 `series`。
    缺失值不在这里处理——它由 `missing_marks()` 单独画，理由见那个函数的 docstring。
    """
    plots = [target] if hasattr(target, "series") else [target]
    for plot in plots:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_value = True
        dl.show_category_name = False
        dl.show_series_name = False
        dl.show_legend_key = False
        if pos is not None:
            dl.position = pos
        dl.number_format = num_fmt
        dl.number_format_is_linked = False
        dl.font.size = Pt(dl_size)
        dl.font.bold = True
        dl.font.name = LAT
        dl.font.color.rgb = RGBColor.from_string(lbl_color or BORDER)
    return target


def _headroom(vals, log_axis=False, headroom=None):
    """算出轴上限：必须给柱端标签留出空间，否则 OUTSIDE_END 的数字会被裁掉。"""
    real = [v for v in vals if v is not None]
    if not real:
        return None
    top = max(real)
    if log_axis:
        return top * (LOG_HEADROOM if headroom is None else 1.0 + headroom)
    return top * (1.0 + (HEADROOM if headroom is None else headroom))


def style_chart(ch, colors, max_val=None, dl_size=16, lbl_size=14, num_fmt="0.000",
                 dl_color=None, cat_color=None, log_axis=False):
    """统一图表风格：藏轴、去网格、逐柱上色、外端数据标签。

    可配置的部分都做成了参数（母版/主题变了不用改这段代码）：
    逐点填色 + 去边框、gap_width=62、外端标签、隐藏数值轴与网格、顶部留白。
    """
    ch.has_title = False
    ch.has_legend = False
    plot = ch.plots[0]
    plot.gap_width = GAP_WIDTH
    ser = plot.series[0]
    try:
        for i, pt in enumerate(ser.points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)])
            pt.format.line.fill.background()
    except Exception:
        pass
    if log_axis:
        _set_log_axis(ch.value_axis, True)
    value_labels(plot, dl_size=dl_size, lbl_color=dl_color, num_fmt=num_fmt,
                 pos=XL_LABEL_POSITION.OUTSIDE_END)
    va = ch.value_axis
    va.has_major_gridlines = False
    if max_val:
        va.maximum_scale = max_val
    _hide_axis(va)
    ca = ch.category_axis
    ca.has_major_gridlines = False
    ca.tick_labels.font.size = Pt(lbl_size)
    ca.tick_labels.font.name = LAT
    ca.tick_labels.font.color.rgb = RGBColor.from_string(cat_color or BODY)
    return ch


def column_chart(s, x, y, w, h, categories, values, colors,
                 series_name="指标", max_val=None, num_fmt="0.000",
                 log_axis=False, missing=MISSING_TEXT, headroom=None,
                 dl_size=16, lbl_size=14, dl_color=None):
    """原生柱状图。比画图片好：可编辑、数值跟着数据走、不触发图片拉伸闸门。

    向后兼容：原有调用 `column_chart(s, x, y, w, h, categories, values, colors)`
    与 `series_name` / `max_val` / `num_fmt` 全部照旧有效。

    新增（都有默认值，不传就是原来的行为）：
      · `values` 里可以放 `None` = 未参评 → 不画柱，灰斜体「未参评」标在原位；
      · `log_axis=True` → 数值轴对数刻度，自动按 LOG_HEADROOM 留标签空间
        （对数是乘性空间，加常数没有意义）；要求所有非 None 值 > 0；
      · `headroom` 覆盖默认顶部留白比例（线性 8%）；
      · `dl_size` / `lbl_size` / `dl_color` 直接改数据标签与类目标签样式。
    """
    cd = CategoryChartData()
    cd.categories = categories
    cd.add_series(series_name, values)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, In(x), In(y), In(w), In(h), cd)
    style_chart(gf.chart, colors, max_val=max_val or _headroom(values, log_axis, headroom),
                dl_size=dl_size, lbl_size=lbl_size, num_fmt=num_fmt, dl_color=dl_color,
                log_axis=log_axis)
    MAN.append(dict(p=PAGE, kind="rect", x=x, y=y, w=w, h=h))
    if missing:
        missing_marks(s, x, y, w, h, len(categories),
                      [i for i, v in enumerate(values) if v is None],
                      vertical=True, lbl_size=lbl_size, text=missing)
    return gf.chart


def hbar_chart(s, x, y, w, h, labels, values, colors,
               series_name="指标", max_val=None, num_fmt="0.0",
               log_axis=False, missing=MISSING_TEXT, headroom=None,
               reverse=True, dl_size=15, lbl_size=13, dl_color=None):
    """原生横向条形图。**类别名长时优先用它**——竖排的类目标签会被折行或自动旋转，
    「跨模态预训练」这类标签在竖排下要么竖着排要么被截断，横向就没这个问题。

    参数与 `column_chart` 同义（`categories`→`labels`、默认数字格式更适合整数），
    另加 `reverse=True`：默认把第一个类别翻到最上面，读图顺序和输入顺序一致。
    """
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series(series_name, values)
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, In(x), In(y), In(w), In(h), cd)
    ch = gf.chart
    if reverse:
        _reverse_categories(ch.category_axis)
    style_chart(ch, colors, max_val=max_val or _headroom(values, log_axis, headroom),
                dl_size=dl_size, lbl_size=lbl_size, num_fmt=num_fmt, dl_color=dl_color,
                log_axis=log_axis)
    MAN.append(dict(p=PAGE, kind="rect", x=x, y=y, w=w, h=h))
    if missing:
        missing_marks(s, x, y, w, h, len(labels),
                      [i for i, v in enumerate(values) if v is None],
                      labels=labels, vertical=False, lbl_size=lbl_size, text=missing)
    return ch


def write_script_md(deck: dict, pages: list | None = None, path=None, title="", meta="", qna=None):
    """把讲稿渲染成 markdown，供提前背稿与现场翻阅。

    讲稿的唯一来源是 `SCRIPT`（{页码: (标题, 计时, 台词)}）：不传 `pages` 时直接由它生成，
    它同时产出 PPT 内嵌备注与本文件，两份同源，不允许各写各的。
    """
    if pages is None:
        pages = [dict(pn=pn, title=v[0], sec=v[1], say=v[2])
                 for pn, v in sorted(SCRIPT.items())]
    out = Path(path or (OUT_DIR / (Path(deck["out"]).stem + "_讲稿.md")))
    lines = [f"# 讲稿 · {title}", "", f"> {meta}", "",
             "照读即可；括号里的短句是现场提示（停顿、指屏、看提问人），读的时候跳过。",
             "", "---", ""]
    for p in pages:
        lines += [f"## {p['pn']}. {p['title']}", "",
                  f"⏱ {p.get('sec', '')}", "", p["say"], ""]
    if qna:
        lines += ["---", "", "## 预设问答", "",
                  "提问方大概率会问的几组，按这个口径答。", ""]
        for q, a in qna:
            lines += [f"**问：{q}**", "", a, ""]
    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"    讲稿 → {out.name}")
    return out


load_theme()          # 导入即套用：有 theme.json 就用，没有就用内置中性主题
