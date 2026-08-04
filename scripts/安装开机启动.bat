@echo off
chcp 65001 >nul
echo 正在安装开机自启动...
schtasks /create /tn "FactoryWorkbench" /tr "C:\Users\pc\ZCodeProject\scripts\run_servers.cmd" /sc onlogon /rl limited /f
echo.
echo 完成！开机自动启动已设置。
echo 重启电脑后，两个版本都会自动运行。
pause
