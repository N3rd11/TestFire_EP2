@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\..\EP3\conectar_smtp.ps1" %*
set "resultado=%errorlevel%"
pause
exit /b %resultado%
