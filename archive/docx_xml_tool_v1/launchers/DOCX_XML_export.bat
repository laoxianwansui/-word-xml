@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "LAUNCHER_DIR=%~dp0"
for %%I in ("%LAUNCHER_DIR%..") do set "REPO_ROOT=%%~fI"
set "TOOL=%REPO_ROOT%\docx_xml_exporter.py"
set "BASE_OUT=%REPO_ROOT%\docx_xml_outputs"
set "LOG=%BASE_OUT%\last_run.log"

if not exist "%BASE_OUT%" mkdir "%BASE_OUT%"

>"%LOG%" echo DOCX XML export started at %date% %time%
>>"%LOG%" echo REPO_ROOT=%REPO_ROOT%
>>"%LOG%" echo TOOL=%TOOL%
>>"%LOG%" echo BASE_OUT=%BASE_OUT%

if not exist "%TOOL%" goto TOOL_MISSING

if "%~1"=="" goto ASK_INPUT
set "DOCX=%~1"
goto HAVE_INPUT

:ASK_INPUT
echo Drag a .docx file onto this .bat file.
echo.
set /p "DOCX=Or paste the full .docx path here and press Enter: "
goto HAVE_INPUT

:HAVE_INPUT
set "DOCX=%DOCX:"=%"
>>"%LOG%" echo DOCX=%DOCX%

if not exist "%DOCX%" goto INPUT_MISSING
if /i not "%DOCX:~-5%"==".docx" goto NOT_DOCX

where py >nul 2>nul
if errorlevel 1 goto RUN_WITH_PYTHON
>>"%LOG%" echo Running with py -3
py -3 "%TOOL%" "%DOCX%" --workspace-root "%BASE_OUT%" >>"%LOG%" 2>&1
goto CHECK_RESULT

:RUN_WITH_PYTHON
>>"%LOG%" echo Running with python
python "%TOOL%" "%DOCX%" --workspace-root "%BASE_OUT%" >>"%LOG%" 2>&1
goto CHECK_RESULT

:CHECK_RESULT
if errorlevel 1 goto RUN_FAILED
echo.
echo Done. Output directory: %BASE_OUT%
echo Log file: %LOG%
>>"%LOG%" echo SUCCESS
pause
exit /b 0

:TOOL_MISSING
echo Tool script not found: %TOOL%
>>"%LOG%" echo ERROR: tool not found
pause
exit /b 1

:INPUT_MISSING
echo Input file not found: %DOCX%
>>"%LOG%" echo ERROR: input file not found
pause
exit /b 1

:NOT_DOCX
echo Only .docx files are supported: %DOCX%
>>"%LOG%" echo ERROR: input is not .docx
pause
exit /b 1

:RUN_FAILED
echo.
echo Export failed. Log file: %LOG%
type "%LOG%"
pause
exit /b 1
