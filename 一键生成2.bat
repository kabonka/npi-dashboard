@echo off

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

echo.
echo === BEFORE (%date% %time%) ===
dir "%~dp0*.html" | findstr ".html"

echo.
echo === STEP 4: Run build_npi.py ===
%PYCMD% "%~dp0build_npi.py"
echo ExitCode: %errorlevel%

echo.
echo === AFTER (%date% %time%) ===
dir "%~dp0*.html" | findstr ".html"

pause
