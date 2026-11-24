@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\..\EP3\ejecutar_demo.ps1" %*
set "resultado=%errorlevel%"
echo.
pause
exit /b %resultado%
