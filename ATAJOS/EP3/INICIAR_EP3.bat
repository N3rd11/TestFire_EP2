@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\..\EP3\abrir_studio.ps1" %*
if errorlevel 1 (pause & exit /b 1)
