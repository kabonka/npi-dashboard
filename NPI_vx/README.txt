NPI Dashboard 便携版 (NPI_vx)
===============================

使用说明：
1. 将 Spec總表.xlsx 放在本文件夹内（已包含示例）
2. 双击「一键生成.bat」或执行：python build_npi.py
3. 生成文件：
   - npi_dashboard.html   主仪表盘（双击用浏览器打开）
   - npi_search.html      搜索仪表盘
   - npi_data.json        数据快照
   - npi_dashboard.xlsx   Excel 报表

依赖：
- Python 3.7+
- openpyxl（首次运行会自动提示安装：pip install openpyxl）

注意事项：
- 请确保 Spec總表.xlsx 中包含「Schedule」工作表
- GitHub 对比功能在无网络环境下会自动跳过
