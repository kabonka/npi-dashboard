# -*- coding: utf-8 -*-
"""
deploy_push.py — 推送 Tencentyyy 部署倉庫到 GitHub（EdgeOne 收到 push 自動部署）。

由 一键生成2.bat 在 sync_web.py（部署複製）之後呼叫。
安全模式（沿用 build_npi.py git_upload 的實戰經驗）：
  1. git fetch origin
  2. 以 FETCH_HEAD 為準 git reset --soft（避免分支分叉，保留工作區檔案）
     ※ 不用 origin/main tracking ref：本機 Tencentyyy 的 refs/remotes 寫入會被吃
       （fetch 報 success 但 ref 消失，原因未明），FETCH_HEAD 一律可靠。
  3. git add -A（部署資料夾只放白名單檔，全部納入管理所當然）
  4. 無變更 → 直接結束；有變更 → commit + push
  5. push 失敗不中斷 bat（印出錯誤供查看），本機 commit 保留、日後可重推

路徑可攜：WEB_ROOT 預設 = NPI_vx 上層資料夾的 Tencentyyy（與 NPI_vx 同級），
可用環境變數 WEB_ROOT 覆蓋。
"""
import os
import subprocess
import sys
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))  # NPI_vx\portal
ROOT = os.path.dirname(BASE)                       # NPI_vx
WEB_ROOT = os.environ.get(
    "WEB_ROOT",
    os.path.abspath(os.path.join(ROOT, "..", "Tencentyyy")),
)


def git(*args):
    return subprocess.run(
        ["git"] + list(args),
        cwd=WEB_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
        stdin=subprocess.DEVNULL,
    )


def main():
    print("=== Deploy push (Tencentyyy) ===")
    print("REPO:", WEB_ROOT)
    if not os.path.isdir(os.path.join(WEB_ROOT, ".git")):
        print("[FAIL] not a git repo:", WEB_ROOT)
        print("       first time: git init + create GitHub repo (see handoff notes)")
        return 1

    r = git("fetch", "origin")
    if r.returncode != 0:
        print("[WARN] fetch failed:", (r.stderr or r.stdout).strip()[:300])
        print("[NOTE] skip reset - push will still detect divergence if any")

    r = git("rev-parse", "--verify", "FETCH_HEAD")
    if r.returncode == 0:
        r = git("reset", "--soft", "FETCH_HEAD")
        if r.returncode != 0:
            print("[WARN] reset --soft failed:", (r.stderr or r.stdout).strip()[:300])
    else:
        print("[WARN] FETCH_HEAD unavailable - skip reset")

    r = git("add", "-A")
    if r.returncode != 0:
        print("[FAIL] git add:", (r.stderr or r.stdout).strip()[:300])
        return 1

    r = git("status", "--porcelain")
    if r.returncode != 0:
        print("[FAIL] git status:", (r.stderr or r.stdout).strip()[:300])
        return 1
    if not r.stdout.strip():
        print("[OK] no change - deploy repo already up to date")
        return 0

    msg = "Deploy NPI portal " + datetime.now().strftime("%Y-%m-%d %H:%M")
    r = git("commit", "-m", msg)
    if r.returncode != 0:
        print("[FAIL] git commit:", (r.stderr or r.stdout).strip()[:300])
        return 1
    print("[OK] committed:", msg)

    r = git("push", "origin", "main")
    if r.returncode != 0:
        # git push 的進度訊息走 stderr；失敗時一併印出供判斷
        print("[FAIL] git push (exit %d)" % r.returncode)
        print((r.stderr or r.stdout).strip()[:500])
        print("[NOTE] local commit kept - retry push later, nothing lost")
        return 1
    print("[OK] pushed to origin/main - EdgeOne will auto-deploy")
    return 0


if __name__ == "__main__":
    sys.exit(main())
