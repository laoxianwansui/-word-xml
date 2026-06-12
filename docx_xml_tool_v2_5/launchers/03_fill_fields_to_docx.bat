@echo off
chcp 65001 >nul
setlocal EnableExtensions

rem 03 字段填充脚本：只能拖“已导出的工作区文件夹”，不能直接拖原始 .docx。
rem 正确流程：先运行 01_export_docx_workspace.bat 导出工作区，编辑 field_values.yaml，再运行本脚本。

set "LAUNCHER_DIR=%~dp0"
for %%I in ("%LAUNCHER_DIR%..") do set "REPO_ROOT=%%~fI"
set "TOOL=%REPO_ROOT%\docx_workspace_tool.py"
set "LOG=%REPO_ROOT%\field_fill_last_run.log"

>"%LOG%" echo DOCX field fill started at %date% %time%
>>"%LOG%" echo REPO_ROOT=%REPO_ROOT%
>>"%LOG%" echo TOOL=%TOOL%

if not exist "%TOOL%" goto TOOL_MISSING

if "%~1"=="" goto ASK_INPUT
set "WORKSPACE=%~1"
goto HAVE_INPUT

:ASK_INPUT
echo 03 字段填充脚本
echo.
echo 请拖入“工作区文件夹”，不能直接拖原始 .docx。
echo 如果你手上只有 .docx，请先运行 01_export_docx_workspace.bat。
echo 然后打开生成工作区里的 field_values.yaml，填写并保存后，再运行本脚本。
echo.
set /p "WORKSPACE=请粘贴工作区文件夹完整路径，然后按 Enter: "
goto HAVE_INPUT

:HAVE_INPUT
set "WORKSPACE=%WORKSPACE:"=%"
>>"%LOG%" echo WORKSPACE=%WORKSPACE%

if not exist "%WORKSPACE%" goto WORKSPACE_MISSING
if /i "%WORKSPACE:~-5%"==".docx" goto RAW_DOCX_INPUT
if not exist "%WORKSPACE%\original.docx" goto NOT_WORKSPACE
if not exist "%WORKSPACE%\content_map.json" goto NOT_WORKSPACE
if not exist "%WORKSPACE%\fields.yaml" goto NOT_WORKSPACE
if not exist "%WORKSPACE%\field_values.yaml" goto NOT_WORKSPACE

where py >nul 2>nul
if errorlevel 1 goto RUN_WITH_PYTHON
>>"%LOG%" echo Running with py -3
py -3 "%TOOL%" "%WORKSPACE%" --fill-fields >>"%LOG%" 2>&1
goto CHECK_RESULT

:RUN_WITH_PYTHON
>>"%LOG%" echo Running with python
python "%TOOL%" "%WORKSPACE%" --fill-fields >>"%LOG%" 2>&1
goto CHECK_RESULT

:CHECK_RESULT
if errorlevel 1 goto RUN_FAILED
echo.
echo 完成。新 DOCX 在这个目录下：%WORKSPACE%\edits
echo 日志文件：%LOG%
>>"%LOG%" echo SUCCESS
pause
exit /b 0

:TOOL_MISSING
echo 找不到工具脚本：%TOOL%
>>"%LOG%" echo ERROR: tool not found
pause
exit /b 1

:WORKSPACE_MISSING
echo 工作区文件夹不存在：%WORKSPACE%
>>"%LOG%" echo ERROR: workspace folder not found
pause
exit /b 1

:RAW_DOCX_INPUT
echo.
echo 你拖入的是原始 .docx，不是工作区文件夹。
echo 03_fill_fields_to_docx.bat 不能直接拖原始 .docx。
echo 请先运行 01_export_docx_workspace.bat 导出工作区。
echo 然后编辑工作区里的 field_values.yaml，保存后把整个工作区文件夹拖到本脚本。
>>"%LOG%" echo ERROR: raw docx input; workspace folder required
pause
exit /b 1

:NOT_WORKSPACE
echo.
echo 这不是可用于字段填充的 DOCX XML 工作区文件夹。
echo 本脚本需要拖入“工作区文件夹”，不能直接拖原始 .docx。
echo 必需文件：original.docx, content_map.json, fields.yaml, field_values.yaml
echo 如果缺少这些文件，请先运行 01_export_docx_workspace.bat。
>>"%LOG%" echo ERROR: required workspace files missing
pause
exit /b 1

:RUN_FAILED
echo.
echo 字段填充失败。日志文件：%LOG%
type "%LOG%"
pause
exit /b 1
