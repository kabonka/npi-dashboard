@echo off
chcp 65001
echo ======================================
echo     GitHub 一键自动上传脚本
echo ======================================
echo.

:: 拉取远程最新代码，避免冲突
git pull

:: 添加所有修改、新增、删除的文件
git add .

:: 自动生成提交时间备注
set "commit_msg=自动更新 %date% %time%"
git commit -m "%commit_msg%"

:: 推送到 GitHub 主分支
git push origin main

echo.
echo ✅ 全部上传完成！
echo ======================================
pause