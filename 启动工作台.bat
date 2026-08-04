@echo off
echo 正在启动工厂工作台...
start /B "" "C:\Users\pc\AppData\Local\Python\bin\python3.exe" -X utf8 "C:\Users\pc\ZCodeProject\server\工厂工作台.py"
start /B "" "C:\Users\pc\AppData\Local\Python\bin\python3.exe" -X utf8 "C:\Users\pc\ZCodeProject\test\工厂工作台.py"
timeout /t 3 >nul
start http://127.0.0.1:8932/
echo 稳定版: http://127.0.0.1:8932/
echo 测试版: http://127.0.0.1:8933/
echo.
pause
