@echo off
rem Arrossega el fitxer .txt de la newsletter a sobre d'aquest fitxer.
rem Tambe pots fer-hi doble clic i triar el fitxer.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0generar.ps1" "%~1"
