@echo off
rem ============================================================
rem  deploy2.bat - DEPLOY + auto-rebuild (since 2026-11-04):
rem    [1] rebuild ONLY the pages whose sources are newer than the HTML
rem        (implemented inside sync_web.py; see the list below)
rem    [2] sync sources into Tencentyyy
rem    [3] push to GitHub -> EdgeOne auto-deploys (~15-20s)
rem  NEW STEP 3: build_whitelist.py regenerates gate.js ALLOWED_EMAILS from portal\whitelist.xlsx.
rem  whitelist.xlsx is a LOCAL source only - it is never deployed to the site.
rem
rem  The per-page update_*.bat files still work, but they are no longer
rem  required before running this file.
rem
rem  MOST pages are now AUTO-REBUILT before deploy (Cert. Map is a static copy). sync_web.py compares
rem  each page's sources against the built HTML and rebuilds only the
rem  stale ones, then copies into Tencentyyy and pushes:
rem    NPI Search       <- NPI_vx\Spec master workbook
rem    TTM Check        <- NPI_vx\Spec master workbook + TTM_expect_mp.json
rem    Spec2 standalone <- spec2\Spec master copy
rem    WebNPI Product Matrix <- webnpi\Spec master copy + mktt.xlsx + 不統計.xlsx + template.html
rem                       (same logic as webnpi\update.bat steps 1-2, but only rebuilds
rem                        when a source is newer; copy+gate done by sync_web.py MAPPING)
rem    KeyPart material <- kpc\KP_material\*.xlsx
rem    WB shipment      <- WB\WB order-control sheet (*.xlsx)
rem    Certificate map  <- AI assistant\certificate1\certificate\MSI_Weibu_认证档案.html (靜態拷貝，手動 rebuild_archive.bat 後部署)
rem    NPI Material     <- materail prep (KP_material, ME material, Spec,
rem                        Model_MKT_list) -> root + web\ -> Tencentyyy
rem  A shared Spec master sync runs first, so editing only the master
rem  workbook in NPI_vx is enough to refresh every Spec-fed page.
rem    Flags: --no-build (deploy existing HTML only)
rem           --dry-run  (print decisions, write nothing)
rem    NOTE: close every Excel / LibreOffice file first. A failed rebuild
rem    aborts the deploy instead of uploading stale pages.
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
echo === STEP 3: Rebuild gate.js whitelist from whitelist.xlsx ===
%PYCMD% "%~dp0build_whitelist.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 4: Refresh NPI Material + sync sources into Tencentyyy ===
%PYCMD% "%~dp0sync_web.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 5: Push deploy repo (Tencentyyy = LIVE site) ===
%PYCMD% "%~dp0deploy_push.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 6: Backup push old repo (kabonka/npi-dashboard) ===
%PYCMD% "%~dp0..\build_npi.py" --upload-only
echo ExitCode: %errorlevel%

echo.
echo ==========================================
echo  DONE. Pushed to GitHub - EdgeOne deploys
echo  in ~15-20 seconds. Ctrl+F5 to verify.
echo ==========================================
pause
