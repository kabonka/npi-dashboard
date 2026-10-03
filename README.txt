NPI Dashboard 便携版 (NPI_vx)
===============================

使用说明：
1. 将 Spec總表.xlsx 放在本文件夹内（已包含示例）
2. 双击「一键生成.bat」或「一键生成2.bat」，会依序重新生成全部仪表板
3. 生成文件：
   - npi_dashboard.html      主仪表盘（双击用浏览器打开）
   - npi_dashboard2.html     NPI 仪表盘（build_npi_1.py 产出）
   - npi_search.html         搜索仪表盘
   - npi_data.json           数据快照
   - npi_dashboard.xlsx      Excel 报表
   - TTM_dashboard.html      TTM 仪表盘（含 📧 週報格式，可长截图 / 另存 Excel）
   - MP变动记录_dashboard.html  MP 变动记录

两个 bat 的差异（只差在 STEP 4 跑哪一支）：
   一键生成.bat   → STEP 4 跑 build_npi_1.py   【本地更新版】
   一键生成2.bat  → STEP 4 跑 build_npi.py     【上传 GitHub 版】
   两支都会依序再跑 build_mp_dashboard.py（STEP 5）与 build_ttm_npi.py（STEP 6）

⚠️ 两支生成器刻意分开、都要保留，用途不同：
   build_npi_1.py  本地更新 —— 跳过 GitHub 比对，不碰 git，只重算本地文件
   build_npi.py    上传 GitHub —— 多出 MP 变动比对（detect_mp_changes /
                   update_mp_change_log）与 git_ensure_repo / git_upload
   两支都会写同一个 npi_search.html（后跑的那支赢），所以两边的前端逻辑
   要同步修改，不要只改一支。

生成器与产物的对应（改的时候要改对文件）：
   build_npi_1.py         → npi_dashboard2.html
   build_npi.py           → npi_dashboard.html / npi_search.html / npi_data.json / npi_dashboard.xlsx
   build_mp_dashboard.py  → MP变动记录_dashboard.html
   build_ttm_npi.py       → TTM_dashboard.html

⚠️ 要修改 TTM_dashboard.html 时，请改 build_ttm_npi.py 再重新生成，
   不要直接编辑 TTM_dashboard.html —— 下次重新生成会被覆盖。

依赖：
- Python 3.7+
- openpyxl（首次运行会自动提示安装：pip install openpyxl）
- lib\ 文件夹（jszip.js / xlsx.js / html2canvas.min.js / crop-screenshot.*），
  缺少时 build_ttm_npi.py 会直接报错列出缺哪几个文件

注意事项：
- 请确保 Spec總表.xlsx 中包含「Schedule」工作表
- GitHub 对比功能在无网络环境下会自动跳过
- 生成的 HTML 都是单一自帶档（JS/CSS 全部 inline），用 file:// 直接打开即可，不需要联网
