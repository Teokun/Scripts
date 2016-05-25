@echo off
REM reg.exe import "%SCRIPTROOT%\timezone.reg"
ping 127.0.0.1 -n 5 >nul
net time \\fmi-data /set /y