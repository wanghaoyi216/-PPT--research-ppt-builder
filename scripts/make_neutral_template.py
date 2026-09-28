# -*- coding: utf-8 -*-
"""
生成内置中性母版 assets/theme-neutral.pptx
================================================

为什么要有这个文件
------------------
做「像某个组会模板」的研究汇报 PPT，最省事、最稳的路子不是自己画一套版式，
而是**把模板文件当母版用**：以模板为基座打开 → add_slide(版式) → 删占位符。
版式里的装饰（非占位符普通形状）会自动渲染在页面下层，一页装饰都不用手画。

但直接分发一个有机构标识的模板文件是版权红线（校名、校徽、校园照片）。
所以这里做一份**零机构标识**的中性母版：保留那套好用的结构
（封面几何块 / 顶部标题带 / 右上角抽象色块 / 结尾页细线 + 色块），
把机构标识换成抽象几何装饰。

三条硬约束
----------
1. 装饰必须是**非占位符普通形状**。只有这样 add_slide() 才只克隆占位符、
   装饰自动落到下层，页面代码一行装饰都不用加（见 kit.new_slide 的 docstring）。
2. 装饰不得遮挡文字区。封面装饰压在左下，文字列从 x≈6.95 起排；
   正文装饰压在 y 0.41–1.24（标题带）与右上角，文字从 y 1.66 起排。
3. 可复现。本文件与 scripts/kit.py 共同构成 assets/themes/*.pptx 的唯一生成方式。

主题
----
六套：`academic-blue`（默认，商务蓝）、`scholar-red`（学术红）、`minimal-mono`（灰阶 + 强调色）、
`midnight`（深色底）、`forest-green`（深绿）、`warm-sand`（暖砂）。
色板与**页面底色**都不在本文件里写死，统一读 `kit.BUILTIN_THEMES`——母版和页面用同一份配置，
两处各写一份迟早会漂，漂了就变成「母版是蓝的、页面是红的」。
改配色改 kit.py 的 BUILTIN_THEMES，本文件自动跟上。

页面底色走主题 schema 的顶层 `background` 字段（不进 palette）：`lt1` 与母版背景都跟着它走，
三套浅色主题的值都是 "FFFFFF"，所以老主题的行为一个字没变。深色主题（midnight）靠这个字段
把整页底色压成深蓝，再配合整套反转的色板（INK 变浅字、WHITE 变面板底）。

用法
----
    python scripts/make_neutral_template.py                    # 默认 academic-blue
    python scripts/make_neutral_template.py --theme scholar-red
    python scripts/make_neutral_template.py -o out.pptx
生成后应当用 scripts/analyze_template.py 跑一遍复核反推结果。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import kit
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches

PKG = Path(__file__).resolve().parent.parent
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

PAGE_W, PAGE_H = 13.333, 7.5          # 16:9
BACKGROUND = "FFFFFF"      # 页面底色默认值：主题没给 background 字段时的回落
WHITE = "FFFFFF"           # 仅作为「浅色默认值」；装饰里的菱形改用 pal["WHITE"]
LATIN, EA = "Times New Roman", "微软雅黑"


def _rel_lum(hexrgb):
    """WCAG 2.1 相对亮度。色板注释里写的每个对比度都由这个函数算出来。"""
    def f(v):
        v = v / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (f(int(hexrgb[i:i + 2], 16)) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(fg, bg) -> float:
    la, lb = _rel_lum(fg), _rel_lum(bg)
    hi, lo = (la, lb) if la >= lb else (lb, la)
    return (hi + 0.05) / (lo + 0.05)


def theme_colors(pal: dict, bg: str = BACKGROUND) -> list:
    """色板 → theme1.xml 的 12 个配色角色。

    为什么还要单独映射一遍：页面上用的是语义名（BAND / TINT），
    主题配色用的是角色名（accent1 / lt2），用户在 PowerPoint 里
    「改主题配色」时改的是后者。两套名字要对上，设计语言才不会散成两套。

    `lt1` 绑**页面底色**而不是 palette 的 WHITE：WHITE 在深色主题里是「面板底」，
    和页面底是两回事。写死 WHITE 的话 PowerPoint 里「文本框默认填充」会跟着变错。
    """
    return [
        ("dk1", pal["INK"]), ("lt1", bg), ("dk2", pal["INK"]), ("lt2", pal["TINT"]),
        ("accent1", pal["ACC"]), ("accent2", pal["BROWN"]), ("accent3", pal["NUM"]),
        ("accent4", pal["BAND"]), ("accent5", pal["BORDER"]), ("accent6", pal["TABBG"]),
        ("hlink", pal["BAND"]), ("folHlink", pal["NUM"]),
    ]

# python-pptx 自带模板的 11 个版式 → 语义化中文名。
# 前四个是 kit.py 需要的三类 + 可选两栏；其余留着但改名，避免 analyze 输出里
# 出现 "Title and Content" 这种和脚本对不上的英文名。
LAYOUT_NAMES = {
    0: "标题幻灯片", 1: "标题和内容", 2: "节标题", 3: "两栏内容", 4: "比较",
    5: "仅标题", 6: "空白", 7: "内容与说明", 8: "图片与说明", 9: "竖排标题与文本",
    10: "末尾幻灯片",
}


def _shape(tree, x, y, w, h, fill, shape="rect", line=None, lw=1.0, alpha=None):
    """往 spTree 里塞一个普通形状（非占位符）。"""
    xml = (
        '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="%s">'
        '<p:nvSpPr><p:cNvPr id="900%d" name="deco%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>%s%s</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (
        A, len(tree), len(tree),
        int(x * 914400), int(y * 914400), int(w * 914400), int(h * 914400),
        {"rect": "rect", "ellipse": "ellipse", "diamond": "diamond"}[shape],
        f'<a:solidFill><a:srgbClr val="{fill}">{alpha}</a:srgbClr></a:solidFill>'
        if fill else "<a:noFill/>",
        f'<a:ln w="{int(lw * 12700)}"><a:solidFill><a:srgbClr val="{line}"/>'
        f'</a:solidFill></a:ln>' if line else "<a:ln><a:noFill/></a:ln>",
    )
    tree.append(etree.fromstring(xml))


def patch_theme(prs, pal, bg: str = BACKGROUND):
    """改写 theme1.xml 的配色与字体族。

    为什么不直接画：主题色一旦对不上，用户在 PowerPoint 里改主题配色时，
    我们硬编码的 srgbClr 不会跟着变——设计语言就散成两套了。
    """
    theme = prs.slide_masters[0].part.part_related_by(RT.THEME)
    root = etree.fromstring(theme.blob)
    clr = root.find(f".//{{{A}}}clrScheme")
    for role, val in theme_colors(pal, bg):
        el = clr.find(f"{{{A}}}{role}")
        for ch in list(el):
            el.remove(ch)
        etree.SubElement(el, f"{{{A}}}srgbClr").set("val", val)
    for kind in ("majorFont", "minorFont"):
        f = root.find(f".//{{{A}}}fontScheme/{{{A}}}{kind}")
        f.find(f"{{{A}}}latin").set("typeface", LATIN)
        f.find(f"{{{A}}}ea").set("typeface", EA)
    theme._blob = etree.tostring(root, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)


def set_page_background(prs, bg: str = BACKGROUND):
    """母版显式设页面底色。函数名从 `white_background` 改过来：底色不再是白。

    不设的话背景色来自主题 lt1，用户换主题时正文页会跟着变色——
    而我们的对比度是按**主题底色**算的（SKY/ACC 在底色上根本不能承字）。
    所以底色必须在母版里写死成显式值，不能靠 lt1 隐式继承。
    """
    m = prs.slide_masters[0]
    bg_el = m._element.find(qn("p:cSld")).find(qn("p:bg"))
    if bg_el is not None:
        m._element.find(qn("p:cSld")).remove(bg_el)
    xml = (
        '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="%s"><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>' % (A, bg)
    )
    cs = m._element.find(qn("p:cSld"))
    cs.insert(0, etree.fromstring(xml))


def decorate(prs, pal):
    """给四个版式加装饰形状（两栏内容沿用正文页装饰）。

    装饰只吃 6 个颜色角色：BAND（标题带/分隔线）、ACC（细线）、SKY（大面底色）、
    TABBG（抽象圆）、BORDER（抽象菱形底）、BROWN（右上角色块）。
    给任何新主题配色时，这 6 个角色都要保证「相互之间不靠承小字」——
    它们全是装饰，闸门三只会因为「文字压在装饰上」才去量对比度。
    """
    BAND, ACC, SKY = pal["BAND"], pal["ACC"], pal["SKY"]
    TABBG, BORDER, BROWN = pal["TABBG"], pal["BORDER"], pal["BROWN"]
    WHITE = pal["WHITE"]          # 菱形标记：浅色主题是白、深色主题是面板色
    scratch = prs.slides.add_slide(prs.slide_masters[0].slide_layouts[6])
    trees = {i: prs.slide_masters[0].slide_layouts[i].shapes._spTree
             for i in (0, 1, 2, 3, 10)}

    # ── 封面：左下低饱和几何块，右侧 x≥6.6 留给文字列，不做任何遮挡 ──
    t = trees[0]
    _shape(t, 0.00, 3.10, 6.35, 4.40, SKY)                    # 大面
    _shape(t, 0.00, 3.10, 6.35, 0.06, ACC)                    # 顶边细线
    _shape(t, 0.95, 4.05, 2.40, 2.40, TABBG, shape="ellipse")    # 抽象圆
    _shape(t, 4.60, 4.85, 1.10, 1.10, BROWN, shape="ellipse", alpha='<a:alpha val="42000"/>')
    _shape(t, 6.20, 3.10, 0.055, 4.40, BAND)                  # 竖向分隔

    # ── 正文：顶部蓝色标题带 + 右上角抽象色块（没有 logo 位）──
    for ti in (trees[1], trees[3]):
        _shape(ti, 0.00, 0.41, PAGE_W, 0.83, BAND)            # 标题带 y 0.41–1.24
        _shape(ti, 0.00, 1.24, PAGE_W, 0.025, ACC)            # 带下装饰线
        _shape(ti, 11.12, 0.30, 2.08, 0.90, BROWN)            # 右上色块
        _shape(ti, 12.02, 0.53, 0.44, 0.44, WHITE, shape="diamond")  # 抽象标记
        _shape(ti, 11.60, 0.53, 0.24, 0.44, WHITE, shape="diamond", alpha='<a:alpha val="70000"/>')

    # ── 节标题：左侧整条竖带 + 右侧大面积留白 ──
    # 这个版式默认没有任何装饰形状，是套设计语言里唯一「裸」的版式。
    # 章节分隔页的内容区按 1.66 起算，但整页只有「章节号 + 一句标题」两行字，
    # 不压装饰的话闸门四必判「顶部大空隙 + 覆盖率 10%」。左侧整条竖带把页面
    # 分成「章节标记 | 留白」两栏，右侧文字列从 x 6.95 起排（与封面同一条竖线）。
    t = trees[2]
    _shape(t, 0.00, 0.00, 4.30, PAGE_H, BAND)                 # 左整条竖带
    _shape(t, 0.00, 0.00, 0.055, PAGE_H, ACC)                 # 左缘细线
    _shape(t, 1.05, 3.10, 2.20, 2.20, WHITE, shape="diamond", alpha='<a:alpha val="70000"/>')
    _shape(t, 1.85, 3.60, 1.00, 1.00, BROWN, shape="ellipse", alpha='<a:alpha val="42000"/>')
    _shape(t, 5.10, 2.30, 7.23, 0.025, ACC)                   # 右侧顶部细线
    _shape(t, 5.10, 6.30, 7.23, 0.025, ACC)                   # 右侧底部细线
    _shape(t, 11.12, 0.60, 2.08, 0.90, BROWN)                 # 右上色块（家族特征）
    _shape(t, 12.02, 0.83, 0.44, 0.44, WHITE, shape="diamond")
    _shape(t, 11.60, 0.83, 0.24, 0.44, WHITE, shape="diamond", alpha='<a:alpha val="70000"/>')

    # ── 结尾：一整块低饱和底 + 上下细线 + 一条小色线 ──
    # 结尾页也要占满内容区。它天然没有标题带（内容区顶仍按 1.66 算），
    # 如果只放几行字，闸门四会判「顶部 1.2in 大空隙 + 覆盖率 10%」。
    # 压一整块底色既好看又把版面撑满，是这类版式的常规解法。
    t = trees[10]
    _shape(t, 0.00, 1.86, PAGE_W, 4.64, SKY)                 # 底：1.86–6.50
    _shape(t, 6.29, 2.14, 0.74, 0.09, BROWN)                 # 顶部小色线（细长条，不构成卡片）
    _shape(t, 4.20, 6.20, 4.93, 0.035, BAND)                 # 底部两条细线
    _shape(t, 5.40, 6.29, 2.53, 0.014, ACC)

    # 装饰加完，scratch 这页不要了
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        prs.part.drop_rel(sld.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
        lst.remove(sld)


def strip_thumbnail(path: Path):
    """删掉 python-pptx 自带模板留下的 docProps/thumbnail.jpeg。

    母版是要随 skill 分发的资产，包里不该留任何来历不明的图片——
    即使它只是 Office 默认模板的缩略图，也没法用 grep 证明它没有机构标识。
    删干净之后，全包 ppt/media 与 docProps 里都没有任何位图，结论是硬的。
    """
    import shutil
    import zipfile
    src = path.with_suffix(".pptx.tmp")
    shutil.move(str(path), str(src))
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "docProps/thumbnail.jpeg":
                continue
            data = zin.read(item.filename)
            if item.filename == "_rels/.rels":
                t = data.decode("utf-8")
                import re as _re
                t = _re.sub(r"<Relationship[^>]*thumbnail[^>]*/>", "", t)
                data = t.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                t = data.decode("utf-8")
                import re as _re
                t = _re.sub(r'<Default[^>]*Extension="jpeg"[^>]*/>', "", t)
                data = t.encode("utf-8")
            zout.writestr(item, data)
    src.unlink()


def build(out: Path, theme_name: str = "academic-blue"):
    th = kit.BUILTIN_THEMES[theme_name]
    pal = dict(kit.NEUTRAL_THEME["palette"])
    pal.update(th["palette"])
    # 页面底色：旧主题没有这个字段时回落到 FFFFFF，行为与旧版完全一致
    bg = th.get("background") or kit.NEUTRAL_THEME.get("background") or BACKGROUND
    prs = Presentation()
    prs.slide_width = Inches(PAGE_W)
    prs.slide_height = Inches(PAGE_H)
    patch_theme(prs, pal, bg)
    set_page_background(prs, bg)
    for i, lay in enumerate(prs.slide_masters[0].slide_layouts):
        if i in LAYOUT_NAMES:
            lay.name = LAYOUT_NAMES[i]
    def set_ph(lay, want, x, y, w, h, n=0):
        """把第 n 个匹配 want 的占位符挪到指定几何（单位英寸）。

        占位符位置不是装饰，但它是**栅格的书面声明**：
        analyze_template.py 正是靠占位符反推 M / CW。
        python-pptx 自带模板的占位符是 0.5in 边距、9in 宽的 Office 默认值，
        和本设计语言的 M=0.66 / CW=12.01 不一致；不校正的话反推出来的 theme
        会带上一套错的栅格。用不了占位符（kit.new_slide 会删掉它们）也照样要对齐，
        否则"模板自带的信息"和"设计规格"就对不上号。
        """
        hit = 0
        for ph in lay.placeholders:
            if want in str(ph.placeholder_format.type):
                if hit == n:
                    ph.left, ph.top = Inches(x), Inches(y)
                    ph.width, ph.height = Inches(w), Inches(h)
                    return True
                hit += 1
        return False

    lays = list(prs.slide_masters[0].slide_layouts)
    M, CW, CT_TOP, CT_BOT = 0.66, 12.01, 1.66, 6.94
    set_ph(lays[1], "TITLE", M, 0.41, 10.30, 0.83)
    set_ph(lays[1], "OBJECT", M, CT_TOP, CW, CT_BOT - CT_TOP)
    set_ph(lays[3], "TITLE", M, 0.41, 10.30, 0.83)
    set_ph(lays[3], "OBJECT", M, CT_TOP, 5.85, CT_BOT - CT_TOP, n=0)
    set_ph(lays[3], "OBJECT", 6.82, CT_TOP, 5.85, CT_BOT - CT_TOP, n=1)
    # 封面文字列起排在右侧：左下压着几何装饰块，遮住就没地方排字了
    set_ph(lays[0], "CENTER_TITLE", 6.95, 1.80, 5.72, 1.10)
    set_ph(lays[0], "SUBTITLE", 6.95, 3.10, 5.72, 2.60)
    # 节标题的文字列同样排在右侧：左侧 x<4.30 压着整条 BAND 竖带
    set_ph(lays[2], "TITLE", 6.95, 2.90, 5.72, 1.30)
    set_ph(lays[2], "BODY", 6.95, 4.30, 5.72, 1.20)
    set_ph(lays[10], "BODY", M, CT_TOP, CW, CT_BOT - CT_TOP)
    set_ph(lays[10], "TITLE", M, CT_TOP, CW, 0.90)

    decorate(prs, pal)
    out.parent.mkdir(parents=True, exist_ok=True)
    kit.scrub_metadata(prs, title=f"{theme_name} 母版")   # 母版也是一份 docProps，一样要擦
    prs.save(out)
    strip_thumbnail(out)
    return out


def write_theme_json(theme_name: str) -> Path:
    """把色板落成 assets/themes/<name>.json，供 kit.load_theme(name=…) 读。

    模板路径写的是**包内相对路径**：kit.apply_theme 会按 PKG 解析，写绝对路径
    会让这份 json 换台机器就失效。
    """
    tpl = f"assets/themes/{theme_name}.pptx"
    th = kit.BUILTIN_THEMES[theme_name]
    data = {"name": theme_name, "template": tpl,
            "background": th.get("background") or BACKGROUND,
            "palette": dict(th["palette"])}
    p = kit.theme_json(theme_name)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return p


def report_contrast(pal: dict, bg: str = BACKGROUND):
    """打印本主题每个颜色的实测对比度。

    注释里写的对比度是「算出来的」，不是「估出来的」：改色时先看这张表，
    低于正文 4.5 / 大字 3.0 的组合不要用。

    `bg` 是**页面底色**（主题 schema 顶层的 background 字段），不是 palette 的 WHITE。
    原来这 6 条检查全拿 WHITE 当底，是因为浅色主题里两个值恰好都是 FFFFFF；
    深色主题下它们分家，检查仍要落在页面底上——那才是正文真正坐着的颜色。
    """
    w = pal["WHITE"]
    checks = [
        ("白字 on BAND", w, pal["BAND"]), ("SKY on BAND", pal["SKY"], pal["BAND"]),
        ("BAND on TINT", pal["BAND"], pal["TINT"]), ("BORDER on 底色", pal["BORDER"], bg),
        ("白字 on BORDER", w, pal["BORDER"]), ("白字 on NUM", w, pal["NUM"]),
        ("INK on TABBG", pal["INK"], pal["TABBG"]), ("INK on 底色", pal["INK"], bg),
        ("BODY on 底色", pal["BODY"], bg), ("MUTED on 底色", pal["MUTED"], bg),
        ("ACC on 底色（装饰线）", pal["ACC"], bg), ("BROWN on 底色（禁用承字）", pal["BROWN"], bg),
    ]
    lows = []
    for name, fg, _bg in checks:
        r = contrast(fg, _bg)
        flag = "ok " if r >= 4.5 else ("大字 " if r >= 3.0 else "低  ")
        if flag.startswith("低"):
            lows.append(name)
        print(f"    {flag}{r:5.2f}:1  {name}")
    print(f"  页面底色 {bg}｜" + ("无「低」" if not lows else f"有「低」：{lows}"))
    return lows


def main():
    ap = argparse.ArgumentParser(description="生成内置母版 assets/themes/<theme>.pptx")
    ap.add_argument("-o", "--out", default="", help="输出路径（默认按主题名写到 assets/themes/）")
    ap.add_argument("--theme", default="academic-blue", choices=sorted(kit.BUILTIN_THEMES),
                    help="主题名（默认 academic-blue，向后兼容）")
    a = ap.parse_args()
    out = Path(a.out) if a.out else (kit.theme_dir() / f"{a.theme}.pptx")
    out = build(out, a.theme)
    prs = Presentation(out)
    names = [l.name for l in prs.slide_masters[0].slide_layouts]
    print(f"母版 → {out}   主题 {a.theme}")
    print(f"  页面 {prs.slide_width / 914400:.3f} × {prs.slide_height / 914400:.2f} in")
    print(f"  版式 {names}")
    for i in (0, 1, 10):
        lay = prs.slide_masters[0].slide_layouts[i]
        deco = [s for s in lay.shapes if not s.is_placeholder]
        print(f"  {lay.name}：装饰 {len(deco)} 个非占位符形状 → "
              + ", ".join(f"{s.name.split()[0]}" for s in deco[:4]) + " …")
    pal = dict(kit.NEUTRAL_THEME["palette"])
    pal.update(kit.BUILTIN_THEMES[a.theme]["palette"])
    bg = kit.BUILTIN_THEMES[a.theme].get("background") or BACKGROUND
    print(f"  页面底色 {bg}")
    report_contrast(pal, bg)
    js = write_theme_json(a.theme)
    print(f"  主题 json → {js.relative_to(PKG)}")
    # academic-blue 另存一份老名字：既有文档、脚本与 SKILL.md 都还指着它
    if a.theme == kit.DEFAULT_THEME and a.out == "":
        alias = PKG / "assets" / "theme-neutral.pptx"
        alias.write_bytes(out.read_bytes())
        print(f"  默认别名 → {alias.relative_to(PKG)}")
    import zipfile
    media = [n for n in zipfile.ZipFile(out).namelist()
             if n.startswith("ppt/media/")
             or (n.startswith("docProps/") and n.endswith((".jpeg", ".jpg", ".png")))]
    print(f"  位图素材：{media or '无（母版零机构标识、零图片）'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

