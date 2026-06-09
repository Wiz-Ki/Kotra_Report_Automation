@echo off
setlocal
cd /d "%~dp0\.."

if "%GITHUB_TOKEN%"=="" if "%GH_TOKEN%"=="" (
  echo GITHUB_TOKEN or GH_TOKEN is required.
  echo Example:
  echo   set GITHUB_TOKEN=ghp_your_token
  echo   usage_monitor\run_monitor.bat
  exit /b 2
)

python usage_monitor\github_usage_monitor.py %*
if errorlevel 1 exit /b %errorlevel%

start "" "%cd%\usage_monitor\output\github_usage_report.xlsx"
start "" "%cd%\usage_monitor\output\dashboard.html"
