@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

:: 1. Build Dashboard
echo [1/3] Building Dashboard...
python build_npi.py
if errorlevel 1 (
    echo Build failed, check build_npi.py output
    pause
    exit /b 1
)
echo.

:: 2. Git commit
echo [2/3] Committing to local repo...
"C:\Program Files\Git\cmd\git.exe" add npi_dashboard.html npi_dashboard2.html npi_search.html npi_dashboard.xlsx npi_data.json
"C:\Program Files\Git\cmd\git.exe" commit -m "auto update %date% %time%" --quiet 2>nul
if errorlevel 1 (
    echo No changes to commit
) else (
    echo Committed
)
echo.

:: 3. Push to GitHub (using local proxy)
echo [3/3] Pushing to GitHub...
set http_proxy=http://127.0.0.1:7890
set https_proxy=http://127.0.0.1:7890
"C:\Program Files\Git\cmd\git.exe" push origin main 2>nul
if errorlevel 1 (
    echo Push failed, check network/proxy
    echo Run manually: git push origin main
) else (
    echo Push success! GitHub Pages will update.
)
echo.
echo All done!
pause
