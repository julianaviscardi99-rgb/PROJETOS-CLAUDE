@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0abrir_resultado_tlm.ps1" %*
if errorlevel 1 pause
