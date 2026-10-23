@echo off
rem Fallback launcher (no exe needed). Starts the same local app on port 4173.
cd /d "%~dp0web"
call launch-windows.bat
