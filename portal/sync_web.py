# -*- coding: utf-8 -*-
"""
sync_web.py - deploy copier: copy NPI_vx clean files into Tencentyyy (git repo root)
and inject front-end gate gate.js (SHA-256, unlock once per origin).
Called by bat; output is GitHub kabonka/npi-portal auto-deploy source.

2026-10-26 fix: re-enable gate.js injection (EdgeOne free plan has no built-in access
password, so the client-side gate is the real protection).
"""
import os
import re
import shutil
import sys
import datetime
import glob as globmod
import subprocess
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))   # NPI_vx\portal
ROOT = os.path.dirname(BASE)                        # NPI_vx
WEB_ROOT = os.environ.get("WEB_ROOT", os.path.abspath(os.path.join(ROOT, "..", "Tencentyyy")))

GATE_SRC = os.path.join(BASE, "gate.js")
HEAD_RE = re.compile(r'(<head[^>]*>)', re.IGNORECASE)
GATE_REF_RE = re.compile(r'\r?\n?[ \t]*<script src="[^"]*gate\.js"></script>\r?\n?', re.IGNORECASE)

CDN_RE = re.compile(r'https://cdnjs\.cloudflare\.com/ajax/libs/xlsx/[0-9.]+/xlsx\.full\.min\.js')

# refresh the homepage "UPDATED <em>YYYY.MM.DD</em>" stamp at deploy time
UPDATED_RE = re.compile(r'(UPDATED\s*<em>)[\d.]+(</em>)')

# (来源, 落点) — 来源相对 BASE(=portal/)；落点相对 Tencentyyy 仓库根目录
MAPPING = [
    ("portal.html",                "index.html"),
    # 入口上的「說明檔」：部署對照表（靜態頁，無生成器）
    ("deploy_guide.html",          "deploy_guide.html"),
    ("version.json",               "version.json"),
    ("DEPLOY_README.md",           "README.md"),
    ("functions/api/expect-mp.js", "functions/api/expect-mp.js"),
    ("../npi_search.html",         "NPI_vx_web/npi_search.html"),
    ("../TTM_check.html",          "NPI_vx_web/TTM_check.html"),
    ("../TTM_expect_mp.json",      "TTM_expect_mp.json"),
    ("../../spec2/index.standalone.html", "spec2_web/index.html"),
    ("../../epeats/epeat_dashboard.html",  "epeats_web/index.html"),
    ("../epeats_web/xlsx.full.min.js",     "epeats_web/xlsx.full.min.js"),
    ("../../kpc/html material/material.html", "kpc_web/material.html"),
    ("../../WB/WB 訂單管控表 - *.html",   "weibu_web/weiBu_shipment.html"),
    ("../../../certificate/dashboard.html", "cert_map_web/dashboard.html"),
    # materail prep 3 deliverables + landing page (self-contained, data inlined)
    ("../../materail prep/web/index.html",                  "material_prep_web/index.html"),
    ("../../materail prep/web/KP公版比案表单_v3.html",       "material_prep_web/KP公版比案表单_v3.html"),
    ("../../materail prep/web/ME_part_Dashboard.html",      "material_prep_web/ME_part_Dashboard.html"),
    ("../../materail prep/web/material_prep_viewer.html",   "material_prep_web/material_prep_viewer.html"),
]

EXTERNAL_KEEP = {"webnpi_web/index.html"}
ALLOWED = {dst for _, dst in MAPPING} | EXTERNAL_KEEP | {"gate.js"}
TEXT_EXT = (".html", ".js", ".css")

# ===========================================================================
# 「NPI Material」(G4) 三個頁面：部署前自動確認是最新產出
#   - 來源資料（Excel）比 HTML 新 => 自動重跑 materail prep 的建置步驟
#   - 再把最新的 HTML 複製到 <materail prep>\ 與 <materail prep>\web\
#   - 之後 MAPPING 照舊把 web\ 的檔複製進 Tencentyyy
# 目的：使用者只要跑「一键生成2.bat」一個動作，就會重建 + 同步 + 上傳。
# 想跳過重建（例如只想重推）=> 加 --no-build；只看判斷不寫檔 => --dry-run
# ===========================================================================
MP_ROOT = os.path.abspath(os.path.join(ROOT, "..", "materail prep"))
MP_FORM = os.path.join(MP_ROOT, "form")
MP_ME = os.path.join(MP_ROOT, "ME_part")
MP_NEW = os.path.join(MP_ROOT, "1new")
MP_WEB = os.path.join(MP_ROOT, "web")

# 產出物（相對 MP_ROOT）：任一存在即可
MP_OUTPUTS = [
    "form/KP*.html",
    "ME_part/ME_part_Dashboard.html",
    "1new/material_prep_viewer.html",
]
# 逐頁判斷「來源比產出物新」才重建（避免不相干的檔案時間戳造成每次重跑）
# (產出物 pattern, [該頁真正的來源 pattern])
MP_GROUPS = [
    ("form/KP*.html", ["form/KP_material/*.xlsx", "form/Spec總表.xlsx"]),
    ("ME_part/ME_part_Dashboard.html", ["ME_part/ME material/*.xlsx"]),
    ("1new/material_prep_viewer.html",
     ["1new/Model_MKT_list.xlsx", "form/KP_material/*.xlsx", "ME_part/ME material/*.xlsx"]),
]
# 重建步驟：(工作目錄, 腳本, 額外參數, 失敗是否致命)
# 註：sync_spec_master.py 已移到 PRE_STEPS，執行前只跑一次。
MP_STEPS = [
    (MP_FORM, "check_inputs.py", [], True),
    (MP_FORM, "build_kp_dropdown2.py", [], True),
    (MP_FORM, "build_kp_form_v3.py", [], True),
    (MP_FORM, "build_export_forms.py", [], True),
    (MP_FORM, "apply_kp_format.py", [], True),
    (MP_ME, "build_dashboard.py", [], True),
    (MP_NEW, "merge_material.py", [], True),           # 預覽
    (MP_NEW, "merge_material.py", ["--apply"], True),  # 正式覆蓋（撞鎖會自行中止）
    (MP_NEW, "build_viewer.py", [], True),
]

# Spec 總表 single-source 同步：先跑一次，之後各頁的「來源比成品新」判斷才有意義
PRE_STEPS = [
    (ROOT, "sync_spec_master.py", [], False),
]

AA_ROOT = os.path.abspath(os.path.join(ROOT, ".."))        # ...\AI assistant
DESKTOP_DIR = os.path.dirname(AA_ROOT)                      # ...\Desktop

# 其餘 6 組頁面：輸出 / 來源 / 重建步驟
OTHER_BUILDS = [
    {
        "name": "NPI Search",
        "pairs": [((ROOT, "npi_search.html"), [(ROOT, "Spec總表.xlsx")])],
        "steps": [(ROOT, "build_npi.py", ["--no-upload"], True)],
        "env": {"NPI_NO_PAUSE": "1"},
    },
    {
        "name": "TTM Check",
        # TTM_expect_mp.json 是人工匯入的輸入（generator 只讀不寫），
        # 所以它算「來源」，不是產出，否則會永遠被判過期。
        "pairs": [((ROOT, "TTM_check.html"),
                   [(ROOT, "Spec總表.xlsx"), (ROOT, "TTM_expect_mp.json")])],
        "steps": [(ROOT, "build_ttm_npi.py", [], True)],
    },
    {
        "name": "Spec2 standalone",
        "pairs": [((os.path.join(AA_ROOT, "spec2"), "index.standalone.html"),
                   [(os.path.join(AA_ROOT, "spec2"), "Spec總表.xlsx")])],
        "steps": [(os.path.join(AA_ROOT, "spec2"), "update_all.py", [], True)],
    },
    {
        "name": "KeyPart material",
        "pairs": [((os.path.join(AA_ROOT, "kpc"), os.path.join("html material", "material.html")),
                   [(os.path.join(AA_ROOT, "kpc"), os.path.join("KP_material", "*.xlsx"))])],
        "steps": [(os.path.join(AA_ROOT, "kpc"), "build_keypart_material.py", [], True)],
    },
    {
        "name": "WB shipment",
        "pairs": [((os.path.join(AA_ROOT, "WB"), "WB 訂單管控表 - *.html"),
                   [(os.path.join(AA_ROOT, "WB"), "WB 訂單管控表 - *.xlsx")])],
        "steps": [(os.path.join(AA_ROOT, "WB"), "build_dashboard.py", [], True)],
    },
    {
        "name": "Certificate map",
        "pairs": [((os.path.join(DESKTOP_DIR, "certificate"), "dashboard.html"),
                   [(os.path.join(DESKTOP_DIR, "certificate"), "certificate0626.xlsx")])],
        "steps": [(os.path.join(DESKTOP_DIR, "certificate"), "build_dashboard.py", [], True)],
    },
]



def _mp_files(rel_patterns):
    out = []
    for pat in rel_patterns:
        out += [p for p in globmod.glob(os.path.join(MP_ROOT, pat)) if os.path.isfile(p)]
    return out


def _stamp(path):
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%m-%d %H:%M")


def mp_stale_reason():
    """回傳 HTML 過期的原因字串；空字串代表不用重建。逐頁比對，避免誤判。"""
    reasons = []
    for out_pat, src_pats in MP_GROUPS:
        outs = _mp_files([out_pat])
        srcs = _mp_files(src_pats)
        if not outs:
            reasons.append("%s：找不到產出" % out_pat)
            continue
        if not srcs:
            continue
        out = max(outs, key=os.path.getmtime)
        newest = max(srcs, key=os.path.getmtime)
        if os.path.getmtime(newest) > os.path.getmtime(out):
            reasons.append("%s 比 %s（%s）新" % (
                os.path.relpath(newest, MP_ROOT), os.path.basename(out), _stamp(out)))
    return "；".join(reasons)


def run_mp_pipeline(dry_run=False):
    """依序重跑 materail prep 建置；回傳致命失敗的步驟清單（空=成功）。"""
    print("\n=== [M] materail prep rebuild ===")
    fails = []
    for cwd, script, extra, fatal in MP_STEPS:
        path = os.path.join(cwd, script)
        label = script + (" " + " ".join(extra) if extra else "")
        if not os.path.isfile(path):
            print("  SKIP (missing script):", label)
            if fatal:
                fails.append(label)
            continue
        print("  ->", label, " (cwd=%s)" % (os.path.relpath(cwd, MP_ROOT) or "."))
        if dry_run:
            continue
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, path] + extra, cwd=cwd, env=env,
                           stdin=subprocess.DEVNULL)
        if r.returncode != 0:
            if fatal:
                print("  [FAIL] %s exit=%s" % (label, r.returncode))
                fails.append(label)
                break
            print("  [WARN] 非致命失敗，繼續：", label)
    return fails


def refresh_mp_web(dry_run=False):
    r"""把最新的產出 HTML 複製到 <materail prep>\ 與 <materail prep>\web\。"""
    print("\n=== [M] refresh materail prep root + web\\ ===")
    targets = []
    for pat in MP_OUTPUTS:
        files = [p for p in globmod.glob(os.path.join(MP_ROOT, pat)) if os.path.isfile(p)]
        if not files:
            print("  SKIP (no build output):", pat)
            continue
        targets.append(max(files, key=os.path.getmtime))
    copied = 0
    for src in targets:
        name = os.path.basename(src)
        for dest_dir in (MP_ROOT, MP_WEB):
            if not os.path.isdir(dest_dir):
                print("  mkdir", dest_dir)
                if not dry_run:
                    os.makedirs(dest_dir, exist_ok=True)
            dst = os.path.join(dest_dir, name)
            if (os.path.isfile(dst)
                    and os.path.getmtime(dst) >= os.path.getmtime(src)
                    and os.path.getsize(dst) == os.path.getsize(src)):
                continue
            print("  ->", name, "=>", os.path.relpath(dst, MP_ROOT), "(src %s)" % _stamp(src))
            if not dry_run:
                shutil.copy2(src, dst)
            copied += 1
    print("  updated", copied, "copy(ies)")
    return copied


def _spec_path(spec):
    """(目錄, 相對檔名或 glob) -> 絕對路徑。"""
    return os.path.join(spec[0], spec[1])


def _files_of(specs):
    out = []
    for spec in specs:
        out += [p for p in globmod.glob(_spec_path(spec)) if os.path.isfile(p)]
    return out


def _pair_reasons(pairs):
    """逐組比對「來源是否比產出新」；回傳原因清單（空=最新）。"""
    reasons = []
    for out_spec, src_specs in pairs:
        outs = _files_of([out_spec])
        srcs = _files_of(src_specs)
        if not outs:
            reasons.append("缺少產出 %s" % out_spec[1])
            continue
        if not srcs:
            continue
        o = max(outs, key=os.path.getmtime)
        s = max(srcs, key=os.path.getmtime)
        if os.path.getmtime(s) > os.path.getmtime(o):
            reasons.append("%s（%s）比 %s（%s）新" % (
                os.path.basename(s), _stamp(s), os.path.basename(o), _stamp(o)))
    return reasons


def _run_steps(steps, dry_run=False, extra_env=None, base=None):
    """依序執行 (cwd, script, args, fatal)；回傳致命失敗的步驟標籤。"""
    fails = []
    for cwd, script, extra, fatal in steps:
        path = os.path.join(cwd, script)
        label = script + (" " + " ".join(extra) if extra else "")
        if not os.path.isfile(path):
            print("    SKIP (missing script):", label)
            if fatal:
                fails.append(label)
            continue
        shown = os.path.relpath(cwd, base) if base else cwd
        print("    -> %s  (cwd=%s)" % (label, shown))
        if dry_run:
            continue
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        if extra_env:
            env.update(extra_env)
        r = subprocess.run([sys.executable, path] + extra, cwd=cwd, env=env,
                           stdin=subprocess.DEVNULL)
        if r.returncode != 0:
            if fatal:
                print("    [FAIL] %s exit=%s" % (label, r.returncode))
                fails.append(label)
                break
            print("    [WARN] 非致命失敗，繼續：", label)
    return fails


def refresh_material_prep(auto_build=True, dry_run=False):
    """部署前把所有頁面更新到最新；回傳 0=可繼續部署，1=中止。"""
    print("\n=== [auto] 檢查 / 重建過期頁面 ===")

    # 1) Spec 總表 single-source 同步（內容真的有變才覆蓋；失敗只警告）
    _run_steps(PRE_STEPS, dry_run=dry_run)

    # 2) 其餘 6 組頁面（逐組判斷，只有過期的才重建）
    for b in OTHER_BUILDS:
        reasons = _pair_reasons(b["pairs"])
        if not reasons:
            print("  [ok]    %s：已是最新" % b["name"])
            continue
        print("  [stale] %s：%s" % (b["name"], "；".join(reasons)))
        if not auto_build:
            print("          --no-build -> 沿用現有 HTML")
            continue
        fails = _run_steps(b["steps"], dry_run=dry_run, extra_env=b.get("env"))
        if fails:
            print("\n[FAIL] %s 重建失敗：%s" % (b["name"], ", ".join(fails)))
            print("       常見原因：來源 xlsx 還開在 Excel / LibreOffice，或腳本本身報錯。")
            return 1

    # 3) NPI Material（3 頁 + 複製到 root / web\）
    if not os.path.isdir(MP_ROOT):
        print("[warn] 找不到 materail prep 資料夾，跳過該組：", MP_ROOT)
        return 0
    reason = mp_stale_reason()
    if reason and auto_build:
        print("  [stale] NPI Material：%s" % reason)
        fails = run_mp_pipeline(dry_run=dry_run)
        if fails:
            print("\n[FAIL] materail prep 重建失敗：", ", ".join(fails))
            print("       最常見原因：materail prep\\1new\\material preparation 內的 xlsx 還開在 Excel / LibreOffice。")
            print("       請全部關閉後重跑；或先手動跑 materail prep\\build_all_master.bat，再跑 --no-build。")
            return 1
    elif reason:
        print("  [stale] NPI Material：%s（--no-build -> 沿用現有 HTML）" % reason)
    else:
        print("  [ok]    NPI Material：已是最新")
    refresh_mp_web(dry_run=dry_run)
    return 0




def resolve_src(src_name):
    """Resolve a MAPPING source; if it uses a wildcard, pick the NEWEST match.

    WB 訂單管控表 emits one dated HTML per xlsx version
    (e.g. "WB 訂單管控表 - 1027.html" / "... - 1029.html"), so the deploy copy
    must always take the latest version instead of a hardcoded old filename.

    Sorting is by FILENAME (the date is embedded and zero-padded), because
    mtime is useless here: rebuilding produces several versions with the
    same timestamp.
    """
    if not any(ch in src_name for ch in "*?["):
        return src_name
    matches = globmod.glob(os.path.join(BASE, src_name))
    if not matches:
        return src_name
    rels = [os.path.relpath(p, BASE).replace(os.sep, "/") for p in matches]
    return sorted(rels)[-1]


def gate_rel_path(dst_rel):
    parts = dst_rel.split("/")
    depth = max(0, len(parts) - 1)
    return ("../" * depth) + "gate.js"


def inject_gate(content, dst_rel):
    content = GATE_REF_RE.sub("", content)
    rel = gate_rel_path(dst_rel)
    tag = '<script src="' + rel + '"></script>\n'
    m = HEAD_RE.search(content)
    if m:
        content = content[:m.end()] + "\n" + tag + content[m.end():]
    else:
        content = tag + content
    return content


def prune(root):
    removed = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d != ".git"]
        for name in filenames:
            p = os.path.join(dirpath, name)
            rel = os.path.relpath(p, root).replace(os.sep, "/")
            if rel not in ALLOWED:
                try:
                    os.remove(p); removed.append(rel)
                except Exception as e:
                    print("  WARN cannot prune:", rel, e)
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        if os.path.abspath(dirpath) == os.path.abspath(root):
            continue
        if ".git" in dirpath.replace(os.sep, "/").split("/"):
            continue
        if not os.listdir(dirpath):
            try: os.rmdir(dirpath)
            except Exception: pass
    return removed


def inject_all_html(root):
    n = 0
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d != ".git"]
        for name in filenames:
            if not name.lower().endswith(".html"):
                continue
            p = os.path.join(dirpath, name)
            rel = os.path.relpath(p, root).replace(os.sep, "/")
            try:
                with open(p, "r", encoding="utf-8", errors="replace") as f:
                    c = f.read()
                c2 = inject_gate(c, rel)
                if c2 != c:
                    with open(p, "w", encoding="utf-8", newline="") as f:
                        f.write(c2)
                    print("  gate injected ->", rel)
                    n += 1
            except Exception as e:
                print("  WARN cannot gate:", rel, e)
    return n


def _file_changed(dst_path, new_bytes):
    """回傳 True 表示 dst 不存在或內容與 new_bytes 不同（需要寫入）。"""
    if not os.path.isfile(dst_path):
        return True
    try:
        with open(dst_path, "rb") as f:
            return f.read() != new_bytes
    except Exception:
        return True


def main(auto_build=True, dry_run=False):
    print("=== Deploy copy -> Tencentyyy (with gate.js) ===")
    print("SRC :", BASE); print("DST :", WEB_ROOT)
    if not os.path.isdir(WEB_ROOT):
        print("[FAIL] WEB_ROOT not found:", WEB_ROOT); return 1
    if not os.path.isfile(GATE_SRC):
        print("[FAIL] gate.js not found in source:", GATE_SRC); return 1

    # [M] 先把 NPI Material 三頁更新到最新（必要時重建），再往下複製到 Tencentyyy
    rc = refresh_material_prep(auto_build=auto_build, dry_run=dry_run)
    if rc:
        print("\n[ABORT] 未寫入 Tencentyyy，避免上傳舊版 HTML。")
        return rc
    if dry_run:
        print("\n[dry-run] 以上僅為判斷與模擬，未寫入任何檔案。")
        return 0

    real_changed = False
    # gate.js：只在內容真的變動時才覆寫（避免每次都產生無意義 diff）
    with open(GATE_SRC, "rb") as f:
        gate_bytes = f.read()
    if _file_changed(os.path.join(WEB_ROOT, "gate.js"), gate_bytes):
        with open(os.path.join(WEB_ROOT, "gate.js"), "wb") as f:
            f.write(gate_bytes)
        real_changed = True
        print("  -> gate.js (root) [changed]")
    else:
        print("  -> gate.js (root) [same]")
    copied = skipped = 0
    today = datetime.datetime.now().strftime("%Y.%m.%d")
    for src_entry, dst_rel in MAPPING:
        src_name = resolve_src(src_entry)
        src_path = os.path.join(BASE, src_name)
        if not os.path.isfile(src_path):
            print("  SKIP (missing source):", src_name); skipped += 1; continue
        dst_path = os.path.join(WEB_ROOT, *dst_rel.split("/"))
        os.makedirs(os.path.dirname(dst_path), exist_ok=True)
        if dst_rel == "version.json":
            # 版本戳記延後到迴圈結束才寫，且只有「本批次真有檔案變動」才更新，
            # 避免沒有 .html 變動時也觸發 git push / EdgeOne 重建。
            continue
        if dst_rel.endswith(TEXT_EXT):
            with open(src_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            content = CDN_RE.sub("xlsx.full.min.js", content)
            # refresh homepage UPDATED stamp to deploy date（同日重跑不會產生 diff）
            content = UPDATED_RE.sub(lambda m: m.group(1) + today + m.group(2), content)
            # 複製當下就把 gate.js 注入 html 內容，使寫入結果與線上最終狀態一致，
            # 否則下一輪複製會把「已注入 gate 的線上檔」蓋成「未注入的源檔」造成誤判變動。
            if dst_rel.endswith(".html"):
                content = inject_gate(content, dst_rel)
            new_bytes = content.encode("utf-8", "replace")
        else:
            with open(src_path, "rb") as f:
                new_bytes = f.read()
        if _file_changed(dst_path, new_bytes):
            with open(dst_path, "wb") as f:
                f.write(new_bytes)
            real_changed = True
            print("  ->", dst_rel, ("(latest: " + src_name + ")") if src_name != src_entry else ""); copied += 1
        else:
            print("  same:", dst_rel)
    removed = prune(WEB_ROOT)
    if removed:
        real_changed = True
        print("  pruned:", ", ".join(removed))
    n = inject_all_html(WEB_ROOT)
    if n:
        real_changed = True
    print("  gate injected into", n, "file(s).")
    # version.json：只有本批次確實有變動（或檔案不存在）才寫入新戳記
    vpath = os.path.join(WEB_ROOT, "version.json")
    if real_changed or not os.path.isfile(vpath):
        now = datetime.datetime.now()
        stamp = "{:04d}{:02d}{:02d}-{:02d}{:02d}{:02d}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
        with open(vpath, "w", encoding="utf-8", newline="") as f:
            f.write(stamp + "\n")
        print("  -> version.json (stamp", stamp + ", 因有變動才更新)"); copied += 1
    else:
        print("  SKIP version.json (no change -> 不觸發 push)")
    print("DONE.", copied, "file(s) deployed,", skipped, "source(s) skipped.",
          "(real_changed=" + str(real_changed) + ")")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-build", action="store_true",
                    help="不重建 materail prep，只把現有最新 HTML 同步/部署")
    ap.add_argument("--dry-run", action="store_true",
                    help="只印出判斷結果與將複製的檔，不寫入任何檔案")
    args = ap.parse_args()
    sys.exit(main(auto_build=not args.no_build, dry_run=args.dry_run))
