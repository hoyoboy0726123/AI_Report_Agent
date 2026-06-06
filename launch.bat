@echo off
REM AI Report Agent - 啟動本機 Web 工作台
chcp 65001 >nul
cd /d "%~dp0"

REM 沿用 AI- 專案的 venv(已含報告引擎與 fastapi/uvicorn)
set PYEXE=..\AI-\.venv\Scripts\python.exe
if not exist "%PYEXE%" (
    echo 找不到 Python venv: %PYEXE%
    echo 請先在 AI- 專案建立 venv 並安裝 backend\requirements.txt
    pause
    exit /b 1
)

echo 啟動後端服務 http://127.0.0.1:8756 ...
start "" http://127.0.0.1:8756
cd backend
"%PYEXE%" server.py
