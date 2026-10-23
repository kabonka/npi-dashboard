# -*- coding: utf-8 -*-
"""
sync_web.py — 部署複製器：把 NPI_vx 產出的乾淨檔案複製進 Tencentyyy（git 倉庫根）。
由 一键生成.bat / 一键生成2.bat 在建置完成後呼叫。

2026-10-23 架構改版（做法 A：Tencentyyy = EdgeOne 唯一部署倉庫）：
  倉庫佈局（一功能一資料夾，入口首頁在根目錄）：
    index.html                 入口首頁（源：portal.html，卡片連結各功能）
    NPI_vx_web/                查詢 + 檢查工具（npi_search / TTM_check，皆自足）
    webnpi_web/index.html      產品矩陣（由 push_webnpi.py 單獨管理，本腳本不清除）
    functions/api/expect-mp.js Pages Function（必须在倉庫根目錄，路由才會是 /api/expect-mp）
    TTM_expect_mp.json         期望MP 靜態基準（頁面 fallback 用，留在根目錄）
    npi_dashboard*.html / *.xlsx / MP变动记录*  根目錄其餘上線檔（互相不連結，留根目錄）

  不再注入 gate.js：線上靠 EdgeOne 訪問密碼全站保護（2026-10-23 與使用者確認）。
  安全網：若來源檔殘留 <script src="gate.js"> 引用會自動剝離；剝不掉就中止（避免頁面被鎖死）。

  prune：刪除 Tencentyyy 中不在部署清單的檔案（自動清除舊 gate.js 版副本），
  但保留 .git 與 webnpi_web/。要新增部署檔記得加進 MAPPING，否則會被 prune 掉。

  路徑可攜：WEB_ROOT 預設 = 本腳本上層資料夾的 Tencentyyy（與 NPI_vx 同級），
  可用環境變數 WEB_ROOT 覆蓋。
"""
import os
import re
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))  # NPI_vx（build 工作區）
WEB_ROOT = os.environ.get(
    "WEB_ROOT",
    os.path.abspath(os.path.join(BASE, "..", "Tencentyyy")),
)

GATE_RE = re.compile(r'[ \t]*<script src="gate\.js"></script>\r?\n')

# (NPI_vx 來源檔) -> (Tencentyyy 倉庫內的相對路徑，用 / 分隔)
MAPPING = [
    ("portal.html",                "index.html"),
    ("npi_search.html",            "NPI_vx_web/npi_search.html"),
    ("TTM_check.html",             "NPI_vx_web/TTM_check.html"),
    ("npi_dashboard.html",         "npi_dashboard.html"),
    ("npi_dashboard2.html",        "npi_dashboard2.html"),
    ("npi_data.json",              "npi_data.json"),
    ("npi_dashboard.xlsx",         "npi_dashboard.xlsx"),
    ("MP变动记录.xlsx",             "MP变动记录.xlsx"),
    ("MP变动记录_dashboard.html",   "MP变动记录_dashboard.html"),
    ("TTM_expect_mp.json",         "TTM_expect_mp.json"),
    ("functions/api/expect-mp.js", "functions/api/expect-mp.js"),
    # 根目錄靜態資源（npi_dashboard 系列的截圖功能用；TTM/npi_search 已全內聯不需）
    ("crop-screenshot.css",        "crop-screenshot.css"),
    ("crop-screenshot.js",         "crop-screenshot.js"),
    ("html2canvas.min.js",         "html2canvas.min.js"),
]

# webnpi_web 由 push_webnpi.py 單獨推送：本腳本不複製它，但 prune 時要保留
EXTERNAL_KEEP = {"webnpi_web/index.html"}

ALLOWED = {dst for _, dst in MAPPING} | EXTERNAL_KEEP
TEXT_EXT = (".html", ".js", ".css")


def prune(root):
    """刪除 root 中不在 ALLOWED 的檔案（保留 .git），並清理空的非 .git 資料夾。"""
    removed = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d != ".git"]  # 不走進 .git
        for name in filenames:
            p = os.path.join(dirpath, name)
            rel = os.path.relpath(p, root).replace(os.sep, "/")
            if rel not in ALLOWED:
                try:
                    os.remove(p)
                    removed.append(rel)
                except Exception as e:
                    print("  WARN cannot prune:", rel, e)
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        if os.path.abspath(dirpath) == os.path.abspath(root):
            continue
        if ".git" in dirpath.replace(os.sep, "/").split("/"):
            continue
        if not os.listdir(dirpath):
            try:
                os.rmdir(dirpath)
            except Exception:
                pass
    return removed


def main():
    print("=== Deploy copy -> Tencentyyy ===")
    print("SRC :", BASE)
    print("DST :", WEB_ROOT)
    if not os.path.isdir(WEB_ROOT):
        print("[FAIL] WEB_ROOT not found:", WEB_ROOT)
        return 1

    copied = stripped = skipped = 0
    for src_name, dst_rel in MAPPING:
        src_path = os.path.join(BASE, src_name)
        if not os.path.isfile(src_path):
            print("  SKIP (missing source):", src_name)
            skipped += 1
            continue
        dst_path = os.path.join(WEB_ROOT, *dst_rel.split("/"))
        os.makedirs(os.path.dirname(dst_path), exist_ok=True)
        if dst_rel.endswith(TEXT_EXT):
            with open(src_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            new_content = GATE_RE.sub("", content)
            if new_content != content:
                stripped += 1
                content = new_content
            if "gate.js" in content:
                print("[FAIL] gate.js reference remains in", src_name, "- abort")
                return 1
            with open(dst_path, "w", encoding="utf-8", newline="") as f:
                f.write(content)
        else:
            shutil.copy2(src_path, dst_path)
        print("  ->", dst_rel)
        copied += 1

    removed = prune(WEB_ROOT)
    if removed:
        print("  pruned:", ", ".join(removed))

    print("DONE. %d file(s) deployed, %d gate-stripped, %d source(s) skipped."
          % (copied, stripped, skipped))
    return 0


if __name__ == "__main__":
    sys.exit(main())
