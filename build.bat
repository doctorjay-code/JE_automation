@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Build python is located automatically by build.py
REM   (1) _dev\build_python.txt   (2) env JE_BUILD_PY   (3) python in PATH

where python >nul 2>nul
if errorlevel 1 goto nopython

python "%~dp0build.py"
if errorlevel 1 goto failed

echo.
pause
exit /b 0

:nopython
echo [ERROR] python not found in PATH.
echo Install Python 3, or put your python path into _dev\build_python.txt
echo.
pause
exit /b 1

:failed
echo.
echo [ERROR] build failed. See the message above.
echo.
pause
exit /b 1
