# -*- coding: utf-8 -*-
"""
闸门二：渲染后几何
==================

读 LibreOffice 导出的 PDF，复核真实落位。查五件事：
  1. 文本越出页面
  2. 正文越进页脚带（页脚文字与页码除外，它们本来就画在那一带）
  3. 文本行两两重叠
  4. 图片被拉伸（放置比例 vs 原图像素比例，直接从 PDF 读，不用外部表）
  5. 空页 / 稀页 / 字体内嵌

为什么重叠要按**行框**算
------------------------
PyMuPDF 的 block 是它自己做的分组，块框是组内所有行的并集。于是并排的左右两栏
会被合成一个大框、互相「重叠」——三栏卡片页、图组页、双列清单全是这种假阳性。
行框（line bbox）才是眼睛实际看到的墨迹范围，重叠判断必须以它为准。

用法：python scripts/qa_pdf.py <render/*.pdf> [--foot-text "页脚口号"] [--manifest 版面清单.json]
      --manifest 用来自动认出页脚那句口号（清单里 role=footer 的文本块），
      不给就用 --foot-text 手写。不给两者时，页脚口号会被误报成「压页脚」——
      因为它本来就画在页脚带里。
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import fitz

# ── 阈值 ──
PT = 72.0
PW, PH = 13.33 * PT, 7.5 * PT      # 页面尺寸（pt）。换页面尺寸改这里
EDGE = 2.0                        # 越界容差（pt）
MIN_OVERLAP = 0.22                # 重叠面积 / 较小块面积 的判定阈值
FOOTER_BAND_TOP = 6.98 * PT       # 页脚带起点，正文不许越进来
STRETCH = 0.02                    # 图片比例偏差阈值
MIN_CHARS = 40                    # 低于此字数判为稀页

FOOT_NUM = re.compile(r"^\d{1,3} / \d{1,3}$")


def area(r):
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def inter(a, b):
    return [max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])]


def blocks_of(page):
    """返回 [(行 bbox, 文本)]，只保留含文字的行框。理由见模块 docstring。"""
    out = []
    for bl in page.get_text("dict")["blocks"]:
        if bl.get("type") != 0:
            continue
        for ln in bl.get("lines", []):
            txt = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
            if txt:
                out.append((tuple(ln["bbox"]), txt))
    return out


def foot_text_from_manifest(path):
    """从版面清单里取页脚口号（role=footer 的文本块）。取不到返回空串。"""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return ""
    for d in data:
        if d.get("role") == "footer":
            t = "".join(r[0] for r in d.get("runs", [])).strip()
            if t:
                return t
    return ""


def main(pdf_path, foot_text=""):
    doc = fitz.open(pdf_path)
    print(f"页数 {len(doc)}　页面 {doc[0].rect.width:.0f}×{doc[0].rect.height:.0f} pt"
          f"（{doc[0].rect.width / PT:.2f}×{doc[0].rect.height / PT:.2f} in）\n")
    problems = 0

    for idx, page in enumerate(doc, 1):
        blk = blocks_of(page)
        total_chars = sum(len(t) for _, t in blk)
        msgs = []

        # 1. 越界
        for bb, t in blk:
            if bb[0] < -EDGE or bb[1] < -EDGE or bb[2] > PW + EDGE or bb[3] > PH + EDGE:
                msgs.append(f"    [越界] bbox=({bb[0]:.0f},{bb[1]:.0f},{bb[2]:.0f},{bb[3]:.0f}) 「{t[:26]}」")
                problems += 1

        # 2. 正文越进页脚带
        for bb, t in blk:
            if bb[1] > FOOTER_BAND_TOP and not (
                (foot_text and t.startswith(foot_text)) or FOOT_NUM.match(t)
            ):
                msgs.append(f"    [压页脚] y0={bb[1]:.0f}pt 「{t[:26]}」")
                problems += 1

        # 3. 文本行互相重叠
        for i in range(len(blk)):
            for j in range(i + 1, len(blk)):
                (a, ta), (b, tb) = blk[i], blk[j]
                ov = area(inter(a, b))
                if ov <= 0:
                    continue
                ratio = ov / min(area(a), area(b))
                if ratio >= MIN_OVERLAP:
                    msgs.append(f"    [重叠 {ratio:.0%}] 「{ta[:20]}」 × 「{tb[:20]}」")
                    problems += 1

        # 4. 图片拉伸核对：放置比例 vs 原图像素比例，都从 PDF 里读，无需外部对照表
        for im in page.get_image_info(xrefs=True):
            bb, wh = im["bbox"], (im["width"], im["height"])
            if wh[1] == 0 or (bb[3] - bb[1]) == 0:
                continue
            placed, src = (bb[2] - bb[0]) / (bb[3] - bb[1]), wh[0] / wh[1]
            if abs(placed - src) / src > STRETCH:
                msgs.append(f"    [拉伸] 放置 {placed:.3f} vs 原图 {src:.3f}"
                            f"（{placed / src - 1:+.1%}）")
                problems += 1

        # 5. 稀页
        if total_chars < MIN_CHARS:
            msgs.append(f"    [稀页] 仅 {total_chars} 字 / {len(blk)} 个文本行框")
            problems += 1

        flag = "  ⚠" if msgs else "  ✓"
        print(f"第 {idx:>2} 页{flag}　文本行 {len(blk):>2}　字符 {total_chars:>4}　"
              f"图片 {len(page.get_image_info())}")
        for m in msgs:
            print(m)

    fonts = set()
    for page in doc:
        for f in page.get_fonts():
            fonts.add(f[3])
    print(f"\nPDF 内嵌字体：{sorted(fonts)}")
    print(f"\n{'全部通过，无几何问题' if problems == 0 else f'共 {problems} 处需要修复'}")
    doc.close()
    return 0 if problems == 0 else 1


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(2)
    pdf, ft = args[0], ""
    for k in ("--foot-text", "--manifest"):
        if k in args:
            v = args[args.index(k) + 1]
            ft = v if k == "--foot-text" else (foot_text_from_manifest(v) or ft)
    sys.exit(main(pdf, ft))
