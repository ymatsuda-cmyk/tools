@echo off
rem Register dashrun:// URL scheme. To remove: install_dashrun.bat -Uninstall
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_dashrun.ps1" %*
pause
