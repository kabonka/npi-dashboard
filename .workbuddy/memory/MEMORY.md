# 长期记忆

## NPI Check 项目

### 文件位置
- 目标文件：`c:\Users\msipm\Desktop\work\Spec總表.xlsx`（已从 NPI check_整理版.xlsx 切换）
- Sheet：Schedule（Row 1=表头, Row 2~874=数据, 共873条）
- 甘特图已集成到 HTML Dashboard，不再需要独立脚本或 PNG 图片

### Dashboard 生成
- 脚本：`c:\Users\msipm\WorkBuddy\20260422080636\build_npi.py`（长期保留，用完不删）
- 功能：读取 Excel → 生成 `npi_data.json` + 把数据嵌入 `npi_dashboard.html` 的 `const DATA = {...}`
- Excel 列映射：A(1)=Model Name, B(2)=MKT Name, C(3)=Series, D(4)=Segment, E(5)=CPU, F(6)=GPU, G(7)=NPM, H(8)=SPM, I(9)=Stage, J(10)=ID Frozen, K(11)=Kickoff, L(12)=DVT-start, M(13)=EVT-start, N(14)=MVT-start, O(15)=BTO ready, P(16)=ATS-start, Q(17)=MP, R(18)=Current Status, S(19)=Highlight
- 运行命令（PowerShell）：`$env:PYTHONIOENCODING="utf-8"; python build_npi.py`
- 桌面 bat：`C:\Users\msipm\Desktop\生成Dashboard.bat`（固定绝对路径）

### 数据结构（Spec總表.xlsx - Schedule sheet）
- Row 1 = 表头
- Row 2~874 = 数据行（873 条记录）
- Stage 可能值：Design, DVT, EVT, MVT, ATS, Study, Pending, BTO-, ATS-, MP
- 统计栏和 Stage 筛选器都是动态从数据中读取，非硬编码

### 统计栏
- 动态：从数据中收集所有出现过的 Stage，按 STAGE_ORDER 优先级排序显示
- 每种 Stage 有独立颜色（STAGE_COLORS 映射）
- 点击统计卡片 = 筛选该 Stage

### 甘特图规则（v3 已实现）
1. **布局**：
   - Row 1 = 统计栏（A1~G1，公式驱动），P1起=年份合并(2024/2025/2026)
   - Row 2 = 表头(A~O) + 月份编号(P起)
   - Row 3~47 = 数据+甘特图
   - Col A(1)~O(15) = 原始数据列
   - Col P(16) 起 = 甘特图（月份列 24/11~26/11）
2. **Stage 时间段逻辑**：
   - Kickoff日期 ~ DVT-start = Kickoff阶段
   - DVT-start ~ EVT-start = DVT阶段
   - EVT-start ~ MVT-start = EVT阶段
   - MVT-start ~ ATS-start = MVT阶段
   - ATS-start ~ MP = ATS阶段
   - MP = 仅标注MP所在月份
3. **na 处理**：某阶段为 na 时直接跳到下一个有效阶段
4. **MP 之后不画**：MP月份之后不标色
5. **字母标记+颜色**：K(蓝)/D(紫)/E(橙)/m(红)/A(绿)/M(粉)
6. **MVT 用小写 m，MP 用大写 M**
7. **空白格**：未标记的甘特图单元格填黑色+黑色框线
8. **月份底色**：2024=蓝浅色, 2025=绿浅色, 2026=橙浅色
9. **年份表头底色**：2024=深蓝, 2025=中蓝, 2026=浅蓝
10. **冻结窗格**：Q3（前16列+前2行固定）
11. **Auto-filter**：已启用（A2:O47）

### Python 环境
- 系统有 Python 3.13，已安装 matplotlib, numpy, openpyxl
- build_npi.py 长期保留，不要删

## 用户偏好
- 英语水平 B1，日常练习听力、跟读、阅读理解
- 交互风格：简洁中文反馈，需要手把手 UI 导航指引
- GitHub ID: kabonka
- 常通过截图求助，逐步给出具体约束迭代优化视觉输出
