@echo off
rem Parts Analysis - local launcher (Windows). Double-click to run. No install needed.
rem Uses Python if available, otherwise falls back to a built-in PowerShell server.
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
  start "Parts Analysis server (close to stop)" cmd /c python -m http.server 4173
  timeout /t 1 >nul
  start "" http://localhost:4173
  exit /b
)

where py >nul 2>nul
if %errorlevel%==0 (
  start "Parts Analysis server (close to stop)" cmd /c py -m http.server 4173
  timeout /t 1 >nul
  start "" http://localhost:4173
  exit /b
)

rem No Python found - use the built-in PowerShell server (opens the browser itself).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launch.ps1"
