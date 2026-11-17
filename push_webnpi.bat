@echo off
REM push_webnpi.bat - sync + push WebNPI Product Matrix page (ASCII only, no Chinese)
REM Source: %WEBNPI_SRC% or default ..\Tencentyyy\webnpi_web\index.html
REM Target: .\webnpi_web\index.html in this repo, then git commit + push
setlocal
set PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if not exist "%PY%" set PY=python
"%PY%" "%~dp0push_webnpi.py"
echo.
pause
