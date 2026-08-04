@echo off
chcp 65001 >nul
cd /d "%~dp0backend"
echo ========================================
echo   工厂工作台 API 服务 v2.0
echo   前后端分离架构
echo ========================================
echo.
echo 正在启动...
echo.
python -m uvicorn main:app --host 0.0.0.0 --port 8933 --reload
pause
