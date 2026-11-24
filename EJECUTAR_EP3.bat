@echo off
setlocal
call "%~dp0ATAJOS\EP3\EJECUTAR_DEMO_EP3.bat" -Principal -Visible %*
exit /b %errorlevel%
