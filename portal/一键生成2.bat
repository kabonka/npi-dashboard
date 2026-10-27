@echo off
rem ============================================================
rem  一键生成2.bat - DEPLOY step only (since 2026-10-26):
rem    sync sources into Tencentyyy -> push to GitHub ->
rem    EdgeOne auto-deploys (~15-20s).
rem
rem  Full flow (3 stages):
rem    [1] source data update  (edit Excel / HTML by hand)
rem    [2] update_xxx.bat      rebuild + sync into Tencentyyy
rem    [3] 一键生成2.bat        upload to EdgeOne  <- this file
rem
rem  This script NO LONGER rebuilds anything. If you changed
rem  source data, run the matching update_*.bat FIRST:
rem    update_npi_search.bat      (Spec總表.xlsx -> NPI Search)
rem    update_TTM_check.bat       (Spec總表.xlsx -> TTM Check)
rem    update_product_matrix.bat  (webnpi xlsx -> Product Matrix)
rem    update_standalone.bat      (spec2 xlsx -> Spec standalone)
rem  portal.html / expect-mp.js / TTM_expect_mp.json have no
rem  generator - just run this file, its sync covers them.
rem ============================================================
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo === STEP 1: Find Python ===

REM Check for Python Launcher first (real Python from python.org)
where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYCMD=py -3
        goto :found
    )
)

REM Check for common real Python install paths
if exist "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python313\python.exe" (
    set PYCMD=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python313\python.exe
    goto :found
)
if exist "C:\Python313\python.exe" (
    set PYCMD=C:\Python313\python.exe
    goto :found
)
if exist "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe" (
    set PYCMD=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe
    goto :found
)
if exist "C:\Python312\python.exe" (
    set PYCMD=C:\Python312\python.exe
    goto :found
)
if exist "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe" (
    set PYCMD=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe
    goto :found
)
if exist "C:\Python311\python.exe" (
    set PYCMD=C:\Python311\python.exe
    goto :found
)

REM Fallback to wherever 'where python' finds it
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python NOT FOUND
    pause
    exit /b 1
)
for /f "delims=" %%i in ('where python') do set PYCMD=%%i
echo WARNING: Using 'where python' result.
echo If you get ExitCode 9009, this is the Microsoft Store stub Python.
echo Install real Python from https://www.python.org/downloads/

:found
echo Using: %PYCMD%
%PYCMD% --version

echo.
echo === STEP 2: Install openpyxl ===
%PYCMD% -m pip install openpyxl 2>nul 1>nul
echo Done

echo.
echo === STEP 3: Sync sources into Tencentyyy (safety pass) ===
%PYCMD% "%~dp0sync_web.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 4: Push deploy repo (Tencentyyy = LIVE site) ===
%PYCMD% "%~dp0deploy_push.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 5: Backup push old repo (kabonka/npi-dashboard) ===
%PYCMD% "%~dp0..\build_npi.py" --upload-only
echo ExitCode: %errorlevel%

echo.
echo ==========================================
echo  DONE. Pushed to GitHub - EdgeOne deploys
echo  in ~15-20 seconds. Ctrl+F5 to verify.
echo ==========================================
pause
