# -*- coding: utf-8 -*-
"""
push_webnpi.py — 单独推送 WebNPI Product Matrix 页面上线（配合 push_webnpi.bat 双击使用）

源头：webnpi\\product_matrix.html（webnpi 项目 build.py 的 canonical 产出；
      本身不含 gate.js 引用 — 那行是 sync_web.py 同步到 Tencentyyy 时才注入的，
      所以本脚本默认源头完全不需要 gate.js）
落点：NPI_vx\\webnpi_web\\index.html（git 仓库内的部署路径 → 线上 /webnpi_web/index.html）

特点：
  - 安全网：若源头万一含 <script src="gate.js"> 引用会自动移除
    （线上靠 EdgeOne 访问密码全站保护，不需要客户端密码门；
      且仓库/线上都没有 gate.js 文件，留着引用会让页面锁死打不开）
  - 只 add / commit / push 这一个文件，不碰一键生成2.bat 上传清单里的其他内容
  - 源头路径可用环境变量 WEBNPI_SRC 覆盖

用法：双击 push_webnpi.bat，或命令行 python push_webnpi.py
"""
import os
import re
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))  # NPI_vx 仓库目录
SRC = os.environ.get(
    "WEBNPI_SRC",
    r"C:\Users\msipm\Desktop\AI assistant\webnpi\product_matrix.html",
)
DST = os.path.join(BASE, "webnpi_web", "index.html")
GATE_RE = re.compile(r'[ \t]*<script src="gate\.js"></script>\r?\n')


def git(*args):
    return subprocess.run(
        ["git"] + list(args),
        cwd=BASE,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
        stdin=subprocess.DEVNULL,
    )


def main():
    if not os.path.isfile(SRC):
        print("[FAIL] source not found: " + SRC)
        print("       (set WEBNPI_SRC to override)")
        return 1

    with open(SRC, "r", encoding="utf-8", errors="replace") as f:
        html = f.read()
    stripped = bool(GATE_RE.search(html))
    html = GATE_RE.sub("", html)
    if "gate.js" in html:
        print("[FAIL] gate.js reference still present after strip; abort")
        return 1

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    with open(DST, "w", encoding="utf-8", newline="") as f:
        f.write(html)
    print("[OK] copied product_matrix.html -> webnpi_web/index.html (%d bytes, gate.js ref %s)"
          % (len(html), "stripped" if stripped else "not-present"))

    r = git("add", "webnpi_web/index.html")
    if r.returncode != 0:
        print("[FAIL] git add: " + (r.stderr or r.stdout))
        return 1

    r = git("status", "--porcelain", "--", "webnpi_web/index.html")
    if r.returncode != 0:
        print("[FAIL] git status: " + (r.stderr or r.stdout))
        return 1
    if not r.stdout.strip():
        print("[OK] no change - already up to date")
        return 0

    r = git("commit", "-m", "Update WebNPI Product Matrix page")
    if r.returncode != 0:
        print("[FAIL] git commit: " + (r.stderr or r.stdout))
        return 1
    print("[OK] committed")

    r = git("push", "origin", "main")
    if r.returncode != 0:
        print("[FAIL] git push (exit %d): %s"
              % (r.returncode, (r.stderr or r.stdout).strip()[:500]))
        return 1
    print("[OK] pushed to origin/main - EdgeOne will auto-deploy")
    return 0


if __name__ == "__main__":
    sys.exit(main())
