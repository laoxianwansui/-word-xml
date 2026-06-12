@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "LAUNCHER_DIR=%~dp0"
for %%I in ("%LAUNCHER_DIR%..") do set "REPO_ROOT=%%~fI"
set "TOOL=%REPO_ROOT%\docx_workspace_tool.py"
set "LOG=%REPO_ROOT%\web_gui_last_run.log"

>"%LOG%" echo DOCX XML Tool v2.5 Web GUI started at %date% %time%
>>"%LOG%" echo REPO_ROOT=%REPO_ROOT%
>>"%LOG%" echo TOOL=%TOOL%

if not exist "%TOOL%" goto TOOL_MISSING

where py >nul 2>nul
if errorlevel 1 goto RUN_WITH_PYTHON
>>"%LOG%" echo Running with py -3
py -3 "%TOOL%" --web-gui --host 127.0.0.1 --port 8765 >>"%LOG%" 2>&1
goto CHECK_RESULT

:RUN_WITH_PYTHON
>>"%LOG%" echo Running with python
python "%TOOL%" --web-gui --host 127.0.0.1 --port 8765 >>"%LOG%" 2>&1
goto CHECK_RESULT

:CHECK_RESULT
if errorlevel 1 goto RUN_FAILED
echo.
echo Web GUI stopped.
echo Log file: %LOG%
>>"%LOG%" echo STOPPED
pause
exit /b 0

:TOOL_MISSING
echo Tool script not found: %TOOL%
>>"%LOG%" echo ERROR: tool not found
pause
exit /b 1

:RUN_FAILED
echo.
echo Web GUI failed. Log file: %LOG%
type "%LOG%"
pause
exit /b 1
