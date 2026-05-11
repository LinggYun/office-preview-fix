@echo off
chcp 65001 >nul
title 一键修复 Office Word Excel PPT 预览

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ==========================================
echo   一键修复 Office Word / Excel / PPT 预览
echo ==========================================
echo.

:: Windows 预览处理程序固定 GUID
set "PREVIEW_HANDLER={8895b1c6-b41f-4c1c-a562-0d564250836f}"

:: Microsoft Office 预览器固定 GUID
set "WORD_PREVIEW={84F66100-FF7C-4FB4-B0C0-02CD7FB668FE}"
set "EXCEL_PREVIEW={00020827-0000-0000-C000-000000000046}"
set "PPT_PREVIEW={65235197-874B-4A07-BDC5-E65EA825B718}"

:: 生成备份目录
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "BK=%USERPROFILE%\Desktop\OfficePreviewBackup_%TS%"
mkdir "%BK%" >nul 2>&1

echo 正在备份相关注册表项到桌面...
echo 备份目录：
echo %BK%
echo.

:: 备份扩展名关联
for %%E in (
.doc .docx .docm .dot .dotx .dotm
.xls .xlsx .xlsm .xlsb .xlt .xltx .xltm
.ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm
) do (
    reg export "HKCU\Software\Classes\%%E" "%BK%\HKCU_%%E.reg" /y >nul 2>&1
)

:: 备份 Office ProgID 关联
for %%P in (
"Word.Document.8"
"Word.Document.12"
"Word.DocumentMacroEnabled.12"
"Word.Template.8"
"Word.Template.12"
"Word.TemplateMacroEnabled.12"
"Excel.Sheet.8"
"Excel.Sheet.12"
"Excel.SheetMacroEnabled.12"
"Excel.SheetBinaryMacroEnabled.12"
"Excel.Template.8"
"Excel.Template"
"Excel.TemplateMacroEnabled"
"PowerPoint.Show.8"
"PowerPoint.Show.12"
"PowerPoint.ShowMacroEnabled.12"
"PowerPoint.SlideShow.8"
"PowerPoint.SlideShow.12"
"PowerPoint.SlideShowMacroEnabled.12"
"PowerPoint.Template.8"
"PowerPoint.Template.12"
"PowerPoint.TemplateMacroEnabled.12"
) do (
    reg export "HKCU\Software\Classes\%%~P" "%BK%\HKCU_%%~P.reg" /y >nul 2>&1
)

reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" "%BK%\HKLM_PreviewHandlers.reg" /y >nul 2>&1
reg export "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "%BK%\HKCU_Explorer_Advanced.reg" /y >nul 2>&1

echo 正在查找 Office 程序路径...
echo.

call :FindApp WINWORD.EXE WORD_EXE
call :FindApp EXCEL.EXE EXCEL_EXE
call :FindApp POWERPNT.EXE PPT_EXE

if defined WORD_EXE (
    echo 找到 Word：
    echo %WORD_EXE%
    echo 正在重新注册 Word 组件...
    start /wait "" "%WORD_EXE%" /r
    echo.
) else (
    echo 未找到 WINWORD.EXE，跳过 Word 组件重新注册。
    echo.
)

if defined EXCEL_EXE (
    echo 找到 Excel：
    echo %EXCEL_EXE%
    echo 正在重新注册 Excel 组件...
    start /wait "" "%EXCEL_EXE%" /regserver
    echo.
) else (
    echo 未找到 EXCEL.EXE，跳过 Excel 组件重新注册。
    echo.
)

if defined PPT_EXE (
    echo 找到 PowerPoint：
    echo %PPT_EXE%
    echo 正在重新注册 PowerPoint 组件...
    start /wait "" "%PPT_EXE%" /regserver
    echo.
) else (
    echo 未找到 POWERPNT.EXE，跳过 PowerPoint 组件重新注册。
    echo.
)

echo 正在写入 Word 预览处理程序...

for %%E in (.doc .docx .docm .dot .dotx .dotm) do (
    reg add "HKCU\Software\Classes\%%E\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%WORD_PREVIEW%" /f >nul
)

for %%P in (
"Word.Document.8"
"Word.Document.12"
"Word.DocumentMacroEnabled.12"
"Word.Template.8"
"Word.Template.12"
"Word.TemplateMacroEnabled.12"
) do (
    reg add "HKCU\Software\Classes\%%~P\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%WORD_PREVIEW%" /f >nul
)

echo 正在写入 Excel 预览处理程序...

for %%E in (.xls .xlsx .xlsm .xlsb .xlt .xltx .xltm) do (
    reg add "HKCU\Software\Classes\%%E\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%EXCEL_PREVIEW%" /f >nul
)

for %%P in (
"Excel.Sheet.8"
"Excel.Sheet.12"
"Excel.SheetMacroEnabled.12"
"Excel.SheetBinaryMacroEnabled.12"
"Excel.Template.8"
"Excel.Template"
"Excel.TemplateMacroEnabled"
) do (
    reg add "HKCU\Software\Classes\%%~P\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%EXCEL_PREVIEW%" /f >nul
)

echo 正在写入 PowerPoint 预览处理程序...

for %%E in (.ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm) do (
    reg add "HKCU\Software\Classes\%%E\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%PPT_PREVIEW%" /f >nul
)

for %%P in (
"PowerPoint.Show.8"
"PowerPoint.Show.12"
"PowerPoint.ShowMacroEnabled.12"
"PowerPoint.SlideShow.8"
"PowerPoint.SlideShow.12"
"PowerPoint.SlideShowMacroEnabled.12"
"PowerPoint.Template.8"
"PowerPoint.Template.12"
"PowerPoint.TemplateMacroEnabled.12"
) do (
    reg add "HKCU\Software\Classes\%%~P\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%PPT_PREVIEW%" /f >nul
)

echo 正在注册 Office 预览器名称...

reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%WORD_PREVIEW%" /t REG_SZ /d "Microsoft Word previewer" /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%EXCEL_PREVIEW%" /t REG_SZ /d "Microsoft Excel previewer" /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%PPT_PREVIEW%" /t REG_SZ /d "Microsoft PowerPoint previewer" /f >nul

:: 兼容部分 32 位 Office / 64 位系统场景
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%WORD_PREVIEW%" /t REG_SZ /d "Microsoft Word previewer" /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%EXCEL_PREVIEW%" /t REG_SZ /d "Microsoft Excel previewer" /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%PPT_PREVIEW%" /t REG_SZ /d "Microsoft PowerPoint previewer" /f >nul 2>&1

echo 正在开启资源管理器预览处理程序...

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowPreviewHandlers" /t REG_DWORD /d 1 /f >nul

echo.
echo 正在重启 Windows 资源管理器...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 >nul
start explorer.exe

echo.
echo ==========================================
echo   修复完成
echo ==========================================
echo.
echo 请重新打开文件夹，测试 Word / Excel / PPT 文件预览。
echo.
echo 备份文件已保存到：
echo %BK%
echo.
echo 如果仍然不能预览，建议重启电脑后再试。
echo 如果重启后仍无效，可能需要修复 Microsoft Office 安装。
echo.
pause
exit /b


:FindApp
set "%~2="

for /f "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%~1" /ve 2^>nul ^| findstr /i "REG_"') do set "%~2=%%b"

if not defined %~2 (
    for /f "tokens=2,*" %%a in ('reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%~1" /ve 2^>nul ^| findstr /i "REG_"') do set "%~2=%%b"
)

if not defined %~2 (
    for /f "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%~1" /ve /reg:32 2^>nul ^| findstr /i "REG_"') do set "%~2=%%b"
)

if not defined %~2 if exist "%ProgramFiles%\Microsoft Office\root\Office16\%~1" set "%~2=%ProgramFiles%\Microsoft Office\root\Office16\%~1"
if not defined %~2 if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\%~1" set "%~2=%ProgramFiles(x86)%\Microsoft Office\root\Office16\%~1"
if not defined %~2 if exist "%ProgramFiles%\Microsoft Office\Office16\%~1" set "%~2=%ProgramFiles%\Microsoft Office\Office16\%~1"
if not defined %~2 if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\%~1" set "%~2=%ProgramFiles(x86)%\Microsoft Office\Office16\%~1"
if not defined %~2 if exist "%ProgramFiles%\Microsoft Office\Office15\%~1" set "%~2=%ProgramFiles%\Microsoft Office\Office15\%~1"
if not defined %~2 if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\%~1" set "%~2=%ProgramFiles(x86)%\Microsoft Office\Office15\%~1"
if not defined %~2 if exist "%ProgramFiles%\Microsoft Office\Office14\%~1" set "%~2=%ProgramFiles%\Microsoft Office\Office14\%~1"
if not defined %~2 if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\%~1" set "%~2=%ProgramFiles(x86)%\Microsoft Office\Office14\%~1"

goto :eof