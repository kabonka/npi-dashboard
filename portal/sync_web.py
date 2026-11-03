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
]

EXTERNAL_KEEP = {"webnpi_web/index.html"}
ALLOWED = {dst for _, dst in MAPPING} | EXTERNAL_KEEP | {"gate.js"}
TEXT_EXT = (".html", ".js", ".css")


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


def main():
    print("=== Deploy copy -> Tencentyyy (with gate.js) ===")
    print("SRC :", BASE); print("DST :", WEB_ROOT)
    if not os.path.isdir(WEB_ROOT):
        print("[FAIL] WEB_ROOT not found:", WEB_ROOT); return 1
    if not os.path.isfile(GATE_SRC):
        print("[FAIL] gate.js not found in source:", GATE_SRC); return 1
    shutil.copy2(GATE_SRC, os.path.join(WEB_ROOT, "gate.js"))
    print("  -> gate.js (root)")
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
            now = datetime.datetime.now()
            stamp = "{:04d}{:02d}{:02d}-{:02d}{:02d}{:02d}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
            with open(dst_path, "w", encoding="utf-8", newline="") as f:
                f.write(stamp + "\n")
            print("  ->", dst_rel, "(stamp", stamp + ")"); copied += 1; continue
        if dst_rel.endswith(TEXT_EXT):
            with open(src_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            content = CDN_RE.sub("xlsx.full.min.js", content)
            # refresh homepage UPDATED stamp to deploy date
            content = UPDATED_RE.sub(lambda m: m.group(1) + today + m.group(2), content)
            with open(dst_path, "w", encoding="utf-8", newline="") as f:
                f.write(content)
        else:
            shutil.copy2(src_path, dst_path)
        print("  ->", dst_rel, ("(latest: " + src_name + ")") if src_name != src_entry else ""); copied += 1
    removed = prune(WEB_ROOT)
    if removed:
        print("  pruned:", ", ".join(removed))
    n = inject_all_html(WEB_ROOT)
    print("  gate injected into", n, "file(s).")
    print("DONE.", copied, "file(s) deployed,", skipped, "source(s) skipped.")
    return 0


if __name__ == "__main__":
    sys.exit(main())