@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "LAUNCHER_DIR=%~dp0"
for %%I in ("%LAUNCHER_DIR%..") do set "REPO_ROOT=%%~fI"
set "TOOL=%REPO_ROOT%\docx_workspace_tool.py"
set "LOG=%REPO_ROOT%\markdown_apply_last_run.log"

>"%LOG%" echo DOCX Markdown apply started at %date% %time%
>>"%LOG%" echo REPO_ROOT=%REPO_ROOT%
>>"%LOG%" echo TOOL=%TOOL%

if not exist "%TOOL%" goto TOOL_MISSING

if "%~1"=="" goto ASK_INPUT
set "WORKSPACE=%~1"
goto HAVE_INPUT

:ASK_INPUT
echo Drag a DOCX XML workspace folder onto this .bat file.
echo.
set /p "WORKSPACE=Or paste the full workspace folder path here and press Enter: "
goto HAVE_INPUT

:HAVE_INPUT
set "WORKSPACE=%WORKSPACE:"=%"
>>"%LOG%" echo WORKSPACE=%WORKSPACE%

if not exist "%WORKSPACE%" goto WORKSPACE_MISSING
if not exist "%WORKSPACE%\original.docx" goto NOT_WORKSPACE
if not exist "%WORKSPACE%\document.md" goto NOT_WORKSPACE
if not exist "%WORKSPACE%\content_map.json" goto NOT_WORKSPACE

where py >nul 2>nul
if errorlevel 1 goto RUN_WITH_PYTHON
>>"%LOG%" echo Running with py -3
py -3 "%TOOL%" "%WORKSPACE%" --apply-md >>"%LOG%" 2>&1
goto CHECK_RESULT

:RUN_WITH_PYTHON
>>"%LOG%" echo Running with python
python "%TOOL%" "%WORKSPACE%" --apply-md >>"%LOG%" 2>&1
goto CHECK_RESULT

:CHECK_RESULT
if errorlevel 1 goto RUN_FAILED
echo.
echo Done. Output DOCX is under: %WORKSPACE%\edits
echo Log file: %LOG%
>>"%LOG%" echo SUCCESS
pause
exit /b 0

:TOOL_MISSING
echo Tool script not found: %TOOL%
>>"%LOG%" echo ERROR: tool not found
pause
exit /b 1

:WORKSPACE_MISSING
echo Workspace folder not found: %WORKSPACE%
>>"%LOG%" echo ERROR: workspace folder not found
pause
exit /b 1

:NOT_WORKSPACE
echo This does not look like a DOCX XML workspace.
echo Required files: original.docx, document.md, content_map.json
>>"%LOG%" echo ERROR: required workspace files missing
pause
exit /b 1

:RUN_FAILED
echo.
echo Apply failed. Log file: %LOG%
type "%LOG%"
pause
exit /b 1
