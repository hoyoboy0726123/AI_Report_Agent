@echo off
chcp 65001 >nul
set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYEXE%" (
  echo [ERROR] Project venv not found: %PYEXE%
  echo Create it with:  py -3.13 -m venv .venv  then  .venv\Scripts\python -m pip install -r backend\requirements.txt
  pause
  exit /b 1
)
echo Starting AI Report Agent backend at http://127.0.0.1:8756 ...
start "" http://127.0.0.1:8756
cd /d "%~dp0backend"
"%PYEXE%" server.py
echo.
echo [Server stopped] exit code %errorlevel%
pause
