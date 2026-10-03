@echo off
rem 讓主控台用 UTF-8，Python 輸出的中文才不會亂碼 (2026-10-03)
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
echo If you get ExitCode 9009, you have the Microsoft Store stub Python.
echo Install real Python from https://www.python.org/downloads/

:found
echo Using: %PYCMD%
%PYCMD% --version

echo.
echo === STEP 2: Install openpyxl ===
%PYCMD% -m pip install openpyxl 2>nul 1>nul
echo Done

echo.
echo === STEP 3: Check files ===
if not exist "%~dp0npi_dashboard.html" (
    echo ERROR: npi_dashboard.html NOT FOUND
    dir "%~dp0"
    pause
    exit /b 1
)
echo npi_dashboard.html OK
if not exist "%~dp0npi_dashboard2.html" (
    echo ERROR: npi_dashboard2.html NOT FOUND
    pause
    exit /b 1
)
echo npi_dashboard2.html OK
if not exist "%~dp0build_npi.py" (
    echo ERROR: build_npi.py NOT FOUND
    pause
    exit /b 1
)
echo build_npi.py OK
if not exist "%~dp0build_ttm_npi.py" (
    echo ERROR: build_ttm_npi.py NOT FOUND
    pause
    exit /b 1
)
echo build_ttm_npi.py OK

echo.
echo === BEFORE (%date% %time%) ===
dir "%~dp0*.html" | findstr ".html"

rem === 順序調整 (2026-10-03) ===
rem build_mp_dashboard.py 會「讀取」build_npi.py 產生的 MP變動記錄.xlsx,
rem 所以 mp/ttm 一定要在 build_npi.py 之後才跑.
rem 但上傳只在最後做一次 -> 三個 dashboard 的最新版一起推上去, 不再慢一輪.
rem (原本 build_npi.py 內建上傳, 現在用 --no-upload 關掉, 改由 STEP 7 統一推送)

echo.
echo === STEP 4: Run build_npi.py (build only, no upload) ===
set NPI_NO_PAUSE=1
%PYCMD% "%~dp0build_npi.py" --no-upload
echo ExitCode: %errorlevel%

echo.
echo === STEP 5: Run build_mp_dashboard.py (MP變動記錄) ===
if exist "%~dp0build_mp_dashboard.py" (
    %PYCMD% "%~dp0build_mp_dashboard.py"
    echo ExitCode: %errorlevel%
) else (
    echo SKIP: build_mp_dashboard.py NOT FOUND
)

echo.
echo === STEP 6: Run build_ttm_npi.py (TTM Dashboard) ===
%PYCMD% "%~dp0build_ttm_npi.py"
echo ExitCode: %errorlevel%

echo.
echo === STEP 7: Upload everything to GitHub (unified push) ===
%PYCMD% "%~dp0build_npi.py" --upload-only
echo ExitCode: %errorlevel%
set NPI_NO_PAUSE=

echo.
echo === AFTER (%date% %time%) ===
dir "%~dp0*.html" | findstr ".html"

echo.
echo ==========================================
echo  DONE. 所有 dashboard 已生成並推送 GitHub
echo ==========================================
pause
