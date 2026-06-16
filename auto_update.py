"""
NPI Dashboard 自动更新脚本
检查 Excel 是否有更新，有则自动构建 Dashboard 并推送到 GitHub Pages。
用于 Windows 任务计划程序定时执行。
更新时通过 Server酱 推送微信通知。
"""

import os
import sys
import json
import subprocess
import datetime
import urllib.request
import urllib.parse
import urllib.error

# ── 路径配置 ──
PROJECT_DIR = r"C:\npi-dashboard"
EXCEL_PATH = r"C:\Users\msipm\Desktop\work\Spec總表.xlsx"
BUILD_SCRIPT = os.path.join(PROJECT_DIR, "build_npi.py")
DASHBOARD_HTML = os.path.join(PROJECT_DIR, "npi_dashboard.html")
DATA_JSON = os.path.join(PROJECT_DIR, "npi_data.json")
STATE_FILE = os.path.join(PROJECT_DIR, ".last_update_state.json")
GIT_EXE = r"C:\Program Files\Git\cmd\git.exe"
LOG_FILE = os.path.join(PROJECT_DIR, "auto_update.log")

# 代理配置
PROXY = "http://127.0.0.1:7890"

# ── 微信通知配置（Server酱）──
# 获取 SendKey: https://sct.ftqq.com 登录 → Key & API → 复制 SendKey
# 留空则不发送通知
SERVERCHAN_SENDKEY = ""  # 已禁用


def log(msg):
    """写日志到文件和控制台"""
    ts = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def send_wechat_notify(title, desp=""):
    """通过 Server酱 发送微信通知"""
    if not SERVERCHAN_SENDKEY:
        log("未配置 Server酱 SendKey，跳过微信通知。")
        return False
    try:
        url = f"https://sctapi.ftqq.com/{SERVERCHAN_SENDKEY}.send"
        data = urllib.parse.urlencode({
            "title": title,
            "desp": desp,
        }).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        with urllib.request.urlopen(req, timeout=10) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            if result.get("code") == 0:
                log("微信通知发送成功！")
                return True
            else:
                log(f"微信通知发送失败: {result.get('message', '未知错误')}")
                return False
    except Exception as e:
        log(f"微信通知发送异常: {e}")
        return False


def get_excel_mtime():
    """获取 Excel 文件的最后修改时间（时间戳）"""
    if not os.path.exists(EXCEL_PATH):
        return None
    return os.path.getmtime(EXCEL_PATH)


def get_last_build_mtime():
    """读取上次构建时记录的 Excel 修改时间"""
    if not os.path.exists(STATE_FILE):
        return None
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
        return state.get("excel_mtime")
    except Exception:
        return None


def save_state(excel_mtime):
    """保存当前 Excel 修改时间到状态文件"""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "excel_mtime": excel_mtime,
            "build_time": datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        }, f, ensure_ascii=False, indent=2)


def run_build():
    """执行 build_npi.py 构建 Dashboard"""
    log("正在构建 Dashboard...")
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    try:
        result = subprocess.run(
            [sys.executable, BUILD_SCRIPT],
            cwd=PROJECT_DIR,
            capture_output=True,
            text=True,
            env=env,
            encoding="utf-8",
            errors="replace",
            timeout=120,  # 2分钟超时，避免无限等待
        )
    except subprocess.TimeoutExpired as e:
        log(f"构建超时（超过120秒），可能被网络请求卡住。stdout: {e.stdout[:500] if e.stdout else 'None'}")
        return False
    if result.returncode != 0:
        log(f"构建失败: {result.stderr.strip()}")
        return False
    log(f"构建成功: {result.stdout.strip().splitlines()[-1] if result.stdout.strip() else 'OK'}")
    return True


def run_git_push():
    """执行 git add/commit/push"""
    log("正在提交并推送到 GitHub Pages...")
    env = os.environ.copy()
    env["http_proxy"] = PROXY
    env["https_proxy"] = PROXY

    # git add
    r = subprocess.run(
        [GIT_EXE, "add", "npi_dashboard.html", "npi_search.html", "npi_dashboard.xlsx", "npi_data.json"],
        cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
    )
    if r.returncode != 0:
        log(f"git add 失败: {r.stderr.strip()}")
        return False

    # git commit
    now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
    r = subprocess.run(
        [GIT_EXE, "commit", "-m", f"auto update {now_str}", "--quiet"],
        cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
    )
    if r.returncode != 0:
        # 没有变更也算成功
        if "nothing to commit" in r.stdout or "nothing to commit" in r.stderr:
            log("没有变更需要提交")
            return True
        log(f"git commit 失败: {r.stderr.strip()}")
        return False

    # git push
    r = subprocess.run(
        [GIT_EXE, "push", "origin", "main"],
        cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
    )
    if r.returncode != 0:
        # push 被拒绝，尝试 stash 后 pull --rebase 再重试
        log("push 被拒绝，尝试 git stash + pull --rebase 后重试...")
        subprocess.run(
            [GIT_EXE, "stash"],
            cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
        )
        pr = subprocess.run(
            [GIT_EXE, "pull", "--rebase", "origin", "main"],
            cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
        )
        if pr.returncode != 0:
            log(f"git pull --rebase 失败: {pr.stderr.strip()}")
            subprocess.run(
                [GIT_EXE, "stash", "pop"],
                cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
            )
            return False
        # rebase 成功，恢复 stash 后重新 push
        subprocess.run(
            [GIT_EXE, "stash", "pop"],
            cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
        )
        r2 = subprocess.run(
            [GIT_EXE, "push", "origin", "main"],
            cwd=PROJECT_DIR, capture_output=True, text=True, env=env, encoding="utf-8", errors="replace"
        )
        if r2.returncode != 0:
            log(f"重试 push 仍然失败: {r2.stderr.strip()}")
            return False

    log("推送成功！GitHub Pages 将自动更新。")
    return True


def main():
    log("=" * 40)
    log("NPI Dashboard 自动更新检查")

    # 1. 获取 Excel 修改时间
    excel_mtime = get_excel_mtime()
    if excel_mtime is None:
        log(f"Excel 文件不存在: {EXCEL_PATH}")
        sys.exit(1)

    excel_time_str = datetime.datetime.fromtimestamp(excel_mtime).strftime("%Y/%m/%d %H:%M:%S")
    log(f"Excel 修改时间: {excel_time_str}")

    # 2. 与上次构建记录比较
    last_mtime = get_last_build_mtime()
    if last_mtime is not None:
        last_time_str = datetime.datetime.fromtimestamp(last_mtime).strftime("%Y/%m/%d %H:%M:%S")
        log(f"上次构建时的 Excel 时间: {last_time_str}")

        if excel_mtime <= last_mtime:
            log("Excel 未更新，无需重新构建。")
            return
    else:
        log("首次运行，将执行构建。")

    # 3. Excel 有更新，执行构建
    log(f"检测到 Excel 有更新！开始构建...")
    if not run_build():
        log("构建失败，终止流程。")
        send_wechat_notify(
            "❌ NPI Dashboard 构建失败",
            f"Excel 检测到更新，但构建失败。\n\n**Excel 修改时间**: {excel_time_str}"
        )
        sys.exit(1)

    # 4. 推送到 GitHub Pages
    push_ok = run_git_push()
    if not push_ok:
        log("推送失败，Dashboard 已本地更新但未同步到 GitHub。")
        # 仍然保存状态，避免重复构建

    # 5. 发送微信通知
    if push_ok:
        send_wechat_notify(
            "✅ NPI Dashboard 已更新",
            f"检测到 Excel 有更新，Dashboard 已自动构建并推送到 GitHub Pages。\n\n"
            f"**Excel 修改时间**: {excel_time_str}\n\n"
            f"[查看 Dashboard](https://kabonka.github.io/npi-dashboard/npi_search.html)"
        )
    else:
        send_wechat_notify(
            "⚠️ NPI Dashboard 本地已更新，GitHub 推送失败",
            f"Excel 检测到更新，Dashboard 已本地构建，但推送到 GitHub 失败。\n\n"
            f"**Excel 修改时间**: {excel_time_str}"
        )

    # 6. 保存状态
    save_state(excel_mtime)
    log(f"已记录 Excel 修改时间: {excel_time_str}")


if __name__ == "__main__":
    main()
