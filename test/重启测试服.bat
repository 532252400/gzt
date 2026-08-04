@echo off
cd /d C:\Users\pc\ZCodeProject\test
taskkill /f /fi "IMAGENAME eq python.exe" /fi "WINDOWTITLE eq TestServer" 2>nul
start "" /B "C:\Program Files\AutoClaw\resources\python\python.exe" "?????.py"
