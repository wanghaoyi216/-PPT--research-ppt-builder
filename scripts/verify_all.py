# -*- coding: utf-8 -*-
"""
一键跑完四道闸门（跨平台，替代 bash 版 verify_all.sh）
======================================================

顺序：build → render → gate1 → gate2 → gate3 → gate4

退出码（CI 只看这个，所以必须能区分「真失败」和「没跑成」）：
    0 = 四道闸门全过
    1 = 有闸门失败，或渲染失败
    2 = 有闸门因为缺环境被 SKIP —— **没有全绿过，不要当通过**

    python scripts/verify_all.py                       # 全跑
    python scripts/verify_all.py --pptx a.pptx --manifest a_layout.json
    python scripts/verify_all.py --skip-build --skip-render   # 只重跑 PDF 三闸
    python scripts/verify_all.py --soffice "C:/Program Files/LibreOffice/program/soffice.exe"
    SOFFICE=/usr/bin/soffice python scripts/verify_all.py

找 soffice 的顺序：--soffice → 环境变量 SOFFICE → PATH → 常见安装位置。
**找不到不会崩**，会打一段清楚的提示并只跑闸门一（源码几何那层不需要渲染），
其余三道标记为 SKIP 并把退出码抬到 2——不静默假装通过。

四个必须知道的坑（都在代码里标了）
---------------------------------
1. `-env:UserInstallation` 的值必须是 **file:/// 三个斜杠**。写成 file://C:/... 时
   soffice 会静默产不出任何文件而 exit 仍为 0，于是 PDF 一直停在旧 mtime 上，
   闸门拿旧 PDF 判新代码——最典型的假绿灯。
2. 转换前先清残留 soffice 进程。残留进程会占住 UserInstallation 目录的锁，
   新转换立刻静默退出、不产出文件。
3. soffice 的写盘比进程退出慢，拿到 0 退出码之后还要等文件落盘。
4. 每次渲染前删掉旧 PDF。否则即使上面三条都出问题，闸门也会拿旧文件判新代码。
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
PKG = BASE.parent
GATES_PDF = ["qa_pdf.py", "qa_visual.py", "qa_layout.py"]

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice", "/usr/bin/libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice(explicit=""):
    for c in (explicit, os.environ.get("SOFFICE", "")):
        if c and (shutil.which(c) or Path(c).exists()):
            return shutil.which(c) or str(Path(c))
    for c in ("soffice", "soffice.exe", "libreoffice"):
        w = shutil.which(c)
        if w:
            return w
    for c in SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    return None


def profile_url(profile: Path) -> str:
    """无头配置目录的 URL 形式。三个斜杠，不能少。"""
    p = profile.resolve().as_posix()
    if not p.startswith("/"):
        p = "/" + p
    return "file://" + p


def kill_soffice():
    """清残留进程：Windows taskkill，POSIX pkill。失败无所谓。"""
    for cmd in (["taskkill", "/F", "/IM", "soffice.bin"],
                ["taskkill", "/F", "/IM", "soffice.exe"],
                ["pkill", "-f", "soffice"]):
        try:
            subprocess.run(cmd, capture_output=True, timeout=20)
        except Exception:
            pass
    time.sleep(2)


def render(pptx: Path, outdir: Path, soffice: str, profile: Path) -> Path | None:
    outdir.mkdir(parents=True, exist_ok=True)
    profile.mkdir(parents=True, exist_ok=True)
    pdf = outdir / (pptx.stem + ".pdf")
    if pdf.exists():
        pdf.unlink()                      # 坑 4：先删旧的，否则会拿旧 PDF 判新代码
    kill_soffice()                        # 坑 2
    cmd = [soffice, "--headless",
           f"-env:UserInstallation={profile_url(profile)}",   # 坑 1：三个斜杠
           "--convert-to", "pdf", "--outdir", str(outdir), str(pptx)]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    for _ in range(15):                   # 坑 3：写盘比进程退出慢
        if pdf.exists() and pdf.stat().st_size > 0:
            break
        time.sleep(1)
    if not (pdf.exists() and pdf.stat().st_size > 0):
        print(f"    [FAIL] 渲染失败：{pptx.name}")
        print(f"           soffice 退出码 {r.returncode}")
        for s in (r.stdout, r.stderr):
            if s and s.strip():
                print("           " + s.strip().replace("\n", "\n           "))
        print(f"           确认没在跑别的 soffice（它会占住 profile 锁），"
              f"或用 --profile 指定独立配置目录。")
        return None
    print(f"    渲染 → {pdf.name}（{pdf.stat().st_size // 1024} KB）")
    return pdf


def run_gate(script: str, args: list[str], tail=25) -> tuple[bool, str]:
    p = subprocess.run([sys.executable, str(BASE / script)] + [str(x) for x in args],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = "\n".join(l for l in (p.stdout or "").splitlines()
                    if "MuPDF error" not in l)
    if p.stderr and p.stderr.strip():
        out += "\n" + p.stderr.strip()
    return p.returncode == 0, out


def main():
    ap = argparse.ArgumentParser(description="构建 + 渲染 + 四道闸门，一次跑完")
    ap.add_argument("--pptx", help="要验的 pptx（不给则跑 scripts/build_deck.py）")
    ap.add_argument("--manifest", help="对应的 text_layout_*.json（不给则按 pptx 名猜）")
    ap.add_argument("--build", default=str(BASE / "build_deck.py"))
    ap.add_argument("--render-dir", default=str(PKG / "out" / "render"))
    # LibreOffice 的无头 profile 是缓存，不是交付物。放在 out/ 会让每次跑验证都往
    # 产物目录里倒几百个文件（实测 out/_lo_profile 一个就有 130+ 项），把交付物搞脏。
    # 放 .cache/ 下跨次复用（复用能避开扩展注册表冷启动），删掉这个目录即可完全清干净。
    ap.add_argument("--profile", default=str(PKG / ".cache" / "lo_profile"))
    ap.add_argument("--soffice", default="")
    ap.add_argument("--skip-build", action="store_true")
    ap.add_argument("--skip-render", action="store_true")
    # 透传给 build 脚本。--theme 是 build_deck.py 的参数，但「换主题后必须重跑四闸门」
    # 和「不换主题也要重跑」是同一条纪律，所以这里必须能一键跑完主题变体；
    # 否则 verify_all 只能用默认主题验，改了主题的稿子就处在没验收过的状态。
    ap.add_argument("--theme", default="",
                    help="透传给 build 脚本的主题名（academic-blue/scholar-red/minimal-mono）")
    ap.add_argument("--tail", type=int, default=25, help="闸门失败时打印的尾部行数")
    a = ap.parse_args()

    rc = 0
    skipped = 0
    build_dir = PKG / "out"
    build_dir.mkdir(parents=True, exist_ok=True)

    # ── 1. 构建 ──
    if not a.skip_build:
        print("== 构建 ==")
        if a.pptx:
            print(f"    [跳过] 用了现成的 {a.pptx}")
        else:
            env = dict(os.environ, PYTHONPATH=str(BASE), OUT_DIR=str(build_dir))
            cmd = [sys.executable, a.build]
            if a.theme:
                cmd += ["--theme", a.theme]
                print(f"    [主题] {a.theme}")
            r = subprocess.run(cmd, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", env=env, cwd=str(BASE))
            print((r.stdout or "").rstrip())
            if r.returncode != 0:
                print((r.stderr or "").rstrip())
                print("构建失败")
                return 1

    pptx = Path(a.pptx) if a.pptx else (build_dir / "示例汇报_中文科研组会.pptx")
    if not pptx.exists():
        found = sorted(build_dir.glob("*.pptx"))
        if a.pptx:
            print(f"    [WARN] 你给的 --pptx 不存在：{pptx}")
        if not found:
            print(f"[FAIL] 找不到要验的 pptx：{pptx}（{build_dir} 下也没有任何 .pptx）")
            return 2
        if len(found) > 1:
            print(f"[FAIL] {build_dir} 下有多份 pptx，无法判断该验哪一份，"
                  f"请用 --pptx 显式指定：{[f.name for f in found]}")
            return 2
        pptx = found[0]
        print(f"    [WARN] 改用 {pptx}"
              f"{'（路径不存在，落在 ' + str(build_dir) + ' 下的唯一一份）' if a.pptx else ''}")
    manifest = Path(a.manifest) if a.manifest else (build_dir / (pptx.stem + "_layout.json"))
    print(f"\n验：{pptx.name}\n    清单：{manifest.name if manifest.exists() else '（缺）'}")

    # ── 2. 渲染 ──
    pdf = None
    if a.skip_render:
        pdf = Path(a.render_dir) / (pptx.stem + ".pdf")
        if not pdf.exists():
            print(f"[FAIL] --skip-render 但 {pdf} 不存在")
            return 2
    else:
        print("\n== 渲染 ==")
        soffice = find_soffice(a.soffice)
        if not soffice:
            print("  [提示] 没找到 LibreOffice（soffice）。三道 PDF 闸门需要它。\n"
                  "         · 安装 LibreOffice，或用 --soffice 指定全路径\n"
                  "         · 例：--soffice \"C:/Program Files/LibreOffice/program/soffice.exe\"\n"
                  "         · 没有它时只跑闸门一（源码几何那层不需要渲染），其余标记 SKIP，\n"
                  "           退出码 2（区别于真失败的 1）——这一稿不算验收通过。\n"
                  "         · 注意 soffice 常常不在 PATH 上，必须给全路径。")
        else:
            print(f"    soffice = {soffice}")
            pdf = render(pptx, Path(a.render_dir), soffice, Path(a.profile))
            if pdf is None:
                rc = 1

    # ── 3. 闸门 ──
    print("\n== 闸门 ==")
    ok, out = run_gate("qa_fit.py", [manifest])
    print(("    [ok] qa_fit\n" if ok else "── [FAIL] qa_fit ──\n") + out)
    rc |= 0 if ok else 1

    if pdf:
        # qa_pdf 需要认出页脚那句口号，否则会把页脚误报成「压页脚」。
        # 清单里 role=footer 的文本块就是它。
        for g, extra in ((GATES_PDF[0], ["--manifest", manifest]), (GATES_PDF[1], []),
                         (GATES_PDF[2], [])):
            ok, out = run_gate(g, [pdf] + extra)
            print((f"    [ok] {g}\n" if ok else f"── [FAIL] {g} ──\n")
                  + ("\n".join(out.splitlines()[-a.tail:]) if not ok else ""))
            rc |= 0 if ok else 1
    else:
        for g in GATES_PDF:
            print(f"    [SKIP] {g}（没有渲染 PDF）")
            skipped += 1
        print("    注意：闸门二/三/四没有跑过，这一稿不能算验收通过。")

    if rc:
        print("\n有闸门失败/未执行，见上")
        return rc
    if skipped:
        print(f"\n未完成验收：{skipped} 道闸门因缺少 LibreOffice 被跳过")
        return 2
    print("\n全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
