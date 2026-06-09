@echo off
setlocal
cd /d "%~dp0\.."

set TASK_NAME=KOTRA GitHub Usage Monitor
set TASK_TIME=09:00
set PYTHON_CMD=python
set SCRIPT_PATH=%cd%\usage_monitor\github_usage_monitor.py

echo Registering daily task: %TASK_NAME%
echo The task uses the current user's environment variables.
echo Make sure GITHUB_TOKEN is saved with:
echo   setx GITHUB_TOKEN ghp_your_token

schtasks /Create /F /SC DAILY /ST %TASK_TIME% /TN "%TASK_NAME%" /TR "\"%PYTHON_CMD%\" \"%SCRIPT_PATH%\""
if errorlevel 1 exit /b %errorlevel%

echo Done. The monitor will update daily at %TASK_TIME%.
