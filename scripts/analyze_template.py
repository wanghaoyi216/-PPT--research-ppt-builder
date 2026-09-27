# -*- coding: utf-8 -*-
"""
从任意 .pptx / .potx 基座反推视觉语言，输出 theme JSON
========================================================

为什么需要它
------------
换一份用户自带的模板，页面代码一个字都不用改——只要让它产出同形状的 theme.json，
kit.load_theme() 套上就行。反推要回答四件事：

  1. 主题色（ppt/theme/*.xml 的 clrScheme）与字体族（fontScheme + 实际 run 里的字体）
  2. 母版 / 版式清单，以及每个版式里**非占位符装饰形状**的真实几何与填充
     （装饰才是模板的灵魂：标题带在哪、右上角挂的什么、封面压了什么）
  3. 栅格：左右边距 M、内容区宽 CW、标题带上下沿、页脚带起点
  4. 版式名能不能对上 kit 的四个角色（cover / content / two_col / close）

用法
----
    python scripts/analyze_template.py <模板.pptx>                 # 打印报告
    python scripts/analyze_template.py <模板.pptx> -o theme.json   # 同时写 theme.json
    python scripts/analyze_template.py <模板.pptx> --strict         # 只认名称命中，否则非零退出

关于「对不上」
-------------
模板的版式名五花八门（"Title Slide"、自定义中文、甚至 "首页"）。所以匹配走两级：
  ① 版式名精确/别名匹配；② 兜底按形状几何 + 占位符类型判定（见 `resolve_layouts`）。
两级都判不出来就**明确报错并列出该模板的全部版式名**，绝不静默降级——
版式配错的后果是版式自带的装饰整片丢失，页面会看起来"没有设计"。

`--strict` 只认**名称命中**：只要有角色是几何兜底猜出来的就非零退出。几何兜底靠的是
「装饰形状个数 + 有标题占位符」这类弱特征，4:3 异种模板很容易碰巧猜对版式序——
猜对不等于对，所以那种情况必须人工在 theme.json 里指定 layouts。
"""
from __future__ import annotations

import argparse
import collections
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

EMU_IN = 914400.0
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# 角色 → 版式名别名。匹配时先按别名表，再按几何兜底。
ALIASES = {
    "cover": ["标题幻灯片", "封面", "Title Slide", "title slide", "首页"],
    "content": ["标题和内容", "正文", "Title and Content", "title and content",
                "Content with Title", "Standard"],
    "two_col": ["两栏内容", "Two Content", "two content", "Comparison"],
    "close": ["末尾幻灯片", "结尾", "Thank you", "End Slide", "结束"],
}

# 各角色的几何特征（兜底用）：(最小宽度占页宽比, 垂直位置区间, 是否需要标题占位符)
GEOM_HINTS = {
    "cover": dict(need_title=True, band=None),
    "content": dict(need_title=True, band=True),
    "two_col": dict(need_title=True, band=True, ncol=2),
    "close": dict(need_title=True, band=False),
}


def inch(v):
    return None if v is None else round(v / EMU_IN, 2)


# ── 主题色 / 字体 ──
def read_theme(path: Path):
    z = zipfile.ZipFile(path)
    names = [n for n in z.namelist() if n.startswith("ppt/theme/") and n.endswith(".xml")]
    if not names:
        return {}, {}
    root = etree.fromstring(z.read(names[0]))
    ns = {"a": A}
    colors = {}
    clr = root.find("a:themeElements/a:clrScheme", ns)
    if clr is not None:
        for el in clr:
            role = etree.QName(el).localname
            srgb = el.find("a:srgbClr", ns)
            sysc = el.find("a:sysClr", ns)
            if srgb is not None:
                colors[role] = srgb.get("val")
            elif sysc is not None:
                colors[role] = sysc.get("lastClr")
    fonts = {}
    for kind in ("majorFont", "minorFont"):
        f = root.find(f"a:themeElements/a:fontScheme/a:{kind}", ns)
        if f is None:
            continue
        lat = f.find("a:latin", ns)
        ea = f.find("a:ea", ns)
        fonts[kind] = {"latin": lat.get("typeface") if lat is not None else None,
                       "ea": ea.get("typeface") if ea is not None else None}
    return colors, fonts


# ── 版式与装饰几何 ──
def shape_fill(shp):
    try:
        f = shp.fill
        if f.type == 1:
            return str(f.fore_color.rgb)
        if f.type is not None and f.type != 5:
            return f"theme:{f.fore_color.theme_color}"
    except Exception:
        pass
    return None


def shape_line(shp):
    try:
        ln = shp.line
        if ln.fill.type == 1:
            return {"color": str(ln.fill.fore_color.rgb),
                    "width_pt": round(ln.width.pt, 2) if ln.width else None}
    except Exception:
        pass
    return None


def dump_layout(lay, page_w, page_h):
    phs, decos = [], []
    for shp in lay.shapes:
        geom = dict(x=inch(shp.left), y=inch(shp.top),
                    w=inch(shp.width), h=inch(shp.height))
        if shp.is_placeholder:
            phs.append(dict(type=str(shp.placeholder_format.type), name=shp.name, **geom))
        else:
            decos.append(dict(name=shp.name, prst=str(shp.shape_type),
                              fill=shape_fill(shp), line=shape_line(shp), **geom))
    decos.sort(key=lambda d: (d["y"] or 0, d["x"] or 0))
    return dict(name=lay.name, placeholders=phs, decorations=decos)


def find_band(lay_json, page_w, page_h, tol=0.06):
    """在版式装饰里找「顶部标题带」：横向贯穿、宽度接近满版、上沿在页面上 1/4 以内。

    这是栅格推断里最关键的一项——内容区顶（CT_TOP）要落在它的下沿之下。
    """
    best = None
    for d in lay_json["decorations"]:
        if not d["w"] or d["w"] < page_w * 0.85:
            continue
        if d["y"] is None or d["y"] > page_h * 0.25:
            continue
        if d["h"] is None or d["h"] < 0.25:
            continue
        if best is None or d["y"] < best["y"]:
            best = d
    if best is None:
        return None
    return dict(y0=round(best["y"], 2), y1=round(best["y"] + best["h"], 2),
                color=best["fill"], tol=tol)


# ── 栅格推断 ──
def infer_grid(master, layouts, page_w, page_h, roles=None):
    """从版式几何反推栅格。

    两个必须排除的干扰项，否则 M 会被算成 0：
      · 满宽装饰（标题带、满幅色块）——它本来就该顶到页边；
      · 日期 / 页脚 / 页码占位符——它们贴页边，跟版心无关。
    所以只拿「宽度落在 25%–85% 页宽之间」的形状算左右边距，
    并且优先用 content 版式（封面和结尾页的构图不代表版心）。
    """
    roles = roles or {}
    order = [roles["content"]] if "content" in roles else []
    order += [l["name"] for l in layouts if l["name"] not in order]
    band = None
    for nm in order:
        lj = next((l for l in layouts if l["name"] == nm), None)
        if lj:
            band = find_band(lj, page_w, page_h)
            if band:
                break

    def usable(g):
        # 上界放到 0.97：真正的满幅装饰（标题带、色带）宽度就是页宽，
        # 卡在 0.85 会把「整宽内容框」一起误排除，版心宽度会被算窄一整块。
        return (g["x"] is not None and g["w"] and 0.25 * page_w <= g["w"] <= 0.97 * page_w)

    def edges(nms):
        ls, rs = [], []
        for nm in nms:
            lj = next((l for l in layouts if l["name"] == nm), None)
            if not lj:
                continue
            for g in lj["placeholders"] + lj["decorations"]:
                if usable(g):
                    ls.append(g["x"])
                    rs.append(g["x"] + g["w"])
        return ls, rs

    # 先只看 content 版式：封面与结尾页的构图（满幅色块、居中大标题）不代表版心，
    # 混进来会把 M 算成 0。content 拿不到再退回全部版式。
    lefts, rights = edges([roles["content"]]) if "content" in roles else ([], [])
    if not lefts:
        lefts, rights = edges(order)
    m = min(lefts) if lefts else round(page_w * 0.05, 2)
    right = max(rights) if rights else page_w - m
    cw = round(right - m, 2)

    band_bottom = band["y1"] if band else round(page_h * 0.165, 2)
    return dict(
        W=round(page_w, 2), H=round(page_h, 2),
        M=round(m, 2), CW=cw,
        FOOT_TOP=round(page_h - 0.52, 2),
        FTY=round(page_h - 0.54, 2),
        CT_TOP=round(band_bottom + 0.42, 2),
        CT_BOT=round(page_h - 0.56, 2),
        band=band,
    )


# ── 版式名解析 ──
def resolve_layouts(layouts, warn=print):
    """把模板的版式名解析成 kit 的四个角色。两级：别名 → 几何兜底。"""
    by_name = {l["name"]: l for l in layouts}
    found, how = {}, {}
    used = set()

    for role, alist in ALIASES.items():
        for a in alist:
            for nm, lj in by_name.items():
                if nm == a or nm.lower() == a.lower():
                    found[role] = nm
                    how[role] = "别名"
                    used.add(nm)
                    break
            if role in found:
                break

    # 几何兜底：还没定下来的角色，找带标题占位符、装饰形状最像模板主体的那个版式。
    # 标记写成「几何兜底（非名称命中）」——它是猜测，不是识别。
    for role in ("content", "cover", "close"):
        if role in found:
            continue
        cand = []
        for nm, lj in by_name.items():
            if nm in used:
                continue
            has_title = any("TITLE" in p["type"] or "CENTER_TITLE" in p["type"]
                            for p in lj["placeholders"])
            ndec = len(lj["decorations"])
            if not has_title:
                continue
            cand.append((ndec, nm))
        if not cand:
            continue
        cand.sort(reverse=True)
        found[role] = cand[0][1]
        how[role] = f"几何兜底（{cand[0][0]} 个装饰形状 + 标题占位符）｜非名称命中"
        used.add(cand[0][1])

    missing = [r for r in ("cover", "content", "close") if r not in found]
    if missing:
        warn("  [需要人工确认] 未能判定角色："
             + "、".join(missing)
             + f"\n      模板全部版式名：{sorted(by_name)}\n"
             f"      处理：把对应版式在 PowerPoint 里改名成 {ALIASES[missing[0]][0]!r} 之类，"
             f"或直接手改 theme.json 的 layouts 字段。")
    return found, how


# ── 页面里的实际用色 / 字体 / 字号 ──
def scan_content(prs):
    col = collections.Counter()
    fnt = collections.Counter()
    siz = collections.Counter()
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_text_frame:
                continue
            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        fnt[r.font.name] += 1
                    if r.font.size:
                        siz[round(r.font.size.pt, 1)] += 1
                    try:
                        col[str(r.font.color.rgb)] += 1
                    except Exception:
                        pass
    return col, fnt, siz


def main():
    ap = argparse.ArgumentParser(description="反推模板视觉语言，输出 theme JSON")
    ap.add_argument("template")
    ap.add_argument("-o", "--out", help="theme.json 输出路径")
    ap.add_argument("--strict", action="store_true", help="版式角色判不全就非零退出")
    ap.add_argument("--name", default=None, help="theme 名，默认取文件名")
    a = ap.parse_args()

    path = Path(a.template)
    if not path.exists():
        print(f"[FAIL] 模板不存在：{path}", file=sys.stderr)
        return 2
    prs = Presentation(str(path))
    pw, ph = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN

    print("=" * 78)
    print(f"[模板] {path}")
    print(f"  页面 {pw:.3f} × {ph:.2f} in　母版 {len(prs.slide_masters)} 个　"
          f"版式 {len(prs.slide_masters[0].slide_layouts)} 个　"
          f"自带示例页 {len(prs.slides)} 页")
    print("=" * 78)

    colors, fonts = read_theme(path)
    print("\n主题配色：")
    for k, v in colors.items():
        print(f"  {k:<9s} {v}")
    print("主题字体：")
    for k, v in fonts.items():
        print(f"  {k:<9s} latin={v.get('latin')}  ea={v.get('ea')}")

    layouts = [dump_layout(l, pw, ph) for l in prs.slide_masters[0].slide_layouts]
    print("\n母版与版式（★ 是 kit 需要的角色）：")
    roles, how = resolve_layouts(layouts)
    for lj in layouts:
        star = next((r for r, n in roles.items() if n == lj["name"]), None)
        tag = f"★ {star}（{how.get(star, '')}）" if star else ""
        print(f"\n  ── {lj['name']} {tag}")
        for p in lj["placeholders"]:
            print(f"      占位 {p['type']:<22s} {p['name'][:18]:<18s} "
                  f"({p['x']}, {p['y']}) {p['w']}×{p['h']}")
        for d in lj["decorations"]:
            print(f"      装饰 {str(d['prst']):<22s} {d['name'][:18]:<18s} "
                  f"({d['x']}, {d['y']}) {d['w']}×{d['h']}  fill={d['fill']} "
                  f"line={d['line']}")

    grid = infer_grid(None, layouts, pw, ph, roles)
    print(f"\n栅格推断：M={grid['M']}  CW={grid['CW']}  "
          f"CT_TOP={grid['CT_TOP']}  CT_BOT={grid['CT_BOT']}  "
          f"FOOT_TOP={grid['FOOT_TOP']}")
    if grid["band"]:
        print(f"  标题带：y {grid['band']['y0']}–{grid['band']['y1']}  "
              f"色 {grid['band']['color']}")
    else:
        print("  标题带：未在版式装饰里找到满宽色带 —— 该模板可能是纯文本模板，"
              "标题带要自己在页面里画（此时把 head() 改成 txt + rect 组合）")

    col, fnt, siz = scan_content(prs)
    if col or fnt:
        print(f"\n模板示例页里实际用到的（若有）：")
        print("  颜色 Top8：", ", ".join(f"{c}×{n}" for c, n in col.most_common(8)))
        print("  字体 Top6：", ", ".join(f"{f}×{n}" for f, n in fnt.most_common(6)))
        print("  字号 Top8：", ", ".join(f"{s}pt×{n}" for s, n in siz.most_common(8)))

    theme = dict(
        name=a.name or path.stem,
        template=str(path),
        palette={},          # 色板语义映射需要人/模型判断，不自动填
        fonts={
            "latin": fonts.get("majorFont", {}).get("latin") or "Times New Roman",
            "ea": fonts.get("minorFont", {}).get("ea") or "微软雅黑",
            "num": fonts.get("majorFont", {}).get("latin") or "Times New Roman",
            "formula": fonts.get("majorFont", {}).get("latin") or "Times New Roman",
        },
        grid={k: v for k, v in grid.items() if k != "band"},
        layouts={k: v for k, v in roles.items() if v},
    )
    print("\n产出的 theme（palette 为空：色板语义映射要人/模型看图决定，别自动猜）")
    print(json.dumps(theme, ensure_ascii=False, indent=1))

    if a.out:
        out = Path(a.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(theme, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\n→ {out}")

    if a.strict:
        missing = [r for r in ("cover", "content", "close") if r not in roles]
        guessed = [r for r in ("cover", "content", "close")
                  if str(how.get(r, "")).startswith("几何兜底")]
        if missing:
            print("\n[FAIL] --strict：版式角色判不全（缺 " + "、".join(missing) + "）",
                  file=sys.stderr)
            return 1
        if guessed:
            # 名称一个都没命中、只靠几何猜出来的角色必须人工确认：
            # 版式序碰巧和默认值一致时会「猜对」，但那是巧合，不是识别。
            print("\n[FAIL] --strict：版式名无法识别，"
                  + "、".join(f"{r}={roles[r]}" for r in guessed)
                  + " 全部是几何兜底猜出来的（非名称命中）。\n"
                  "       模板全部版式名：" + str(sorted(l["name"] for l in layouts)) + "\n"
                  "       处理：把对应版式在 PowerPoint 里改名成 "
                  + "、".join(repr(ALIASES[r][0]) for r in guessed)
                  + " 之类，或直接手工在 theme.json 的 layouts 字段里指定。",
                  file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
