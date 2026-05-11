@echo off
chcp 65001 >nul
title Office Preview Fix - 修复 Word Excel PowerPoint 预览

:: ============================================================
:: OfficePreviewFix.bat
:: 修复 Windows 资源管理器中 Word / Excel / PowerPoint 预览功能
::
:: 功能：
:: 1. 备份相关注册表项到桌面
:: 2. 修复 Word / Excel / PowerPoint 文件预览处理程序
:: 3. 注册 Office 预览器名称
:: 4. 开启资源管理器预览处理程序
:: 5. 重启 Windows 资源管理器
::
:: 注意：
:: 本版本默认不会启动 Word / Excel / PowerPoint，
:: 避免运行过程中弹出 Excel 或其他 Office 程序。
::
:: License: MIT
:: ============================================================

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ==========================================
echo   Office Preview Fix
echo   修复 Word / Excel / PowerPoint 预览
echo ==========================================
echo.

echo 本脚本将修复以下文件在 Windows 资源管理器预览窗格中的预览功能：
echo.
echo Word:
echo   .doc .docx .docm .dot .dotx .dotm
echo.
echo Excel:
echo   .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm
echo.
echo PowerPoint:
echo   .ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm
echo.
echo 注意：本脚本会修改注册表，但会先自动备份相关注册表项。
echo.

pause

echo.
echo 正在准备修复...
echo.

:: Windows Shell Preview Handler 固定 GUID
set "PREVIEW_HANDLER={8895b1c6-b41f-4c1c-a562-0d564250836f}"

:: Microsoft Office Preview Handler 固定 GUID
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

:: 备份 Word ProgID 关联
for %%P in (
"Word.Document.8"
"Word.Document.12"
"Word.DocumentMacroEnabled.12"
"Word.Template.8"
"Word.Template.12"
"Word.TemplateMacroEnabled.12"
) do (
    reg export "HKCU\Software\Classes\%%~P" "%BK%\HKCU_%%~P.reg" /y >nul 2>&1
)

:: 备份 Excel ProgID 关联
for %%P in (
"Excel.Sheet.8"
"Excel.Sheet.12"
"Excel.SheetMacroEnabled.12"
"Excel.SheetBinaryMacroEnabled.12"
"Excel.Template.8"
"Excel.Template"
"Excel.Template.12"
"Excel.TemplateMacroEnabled"
"Excel.TemplateMacroEnabled.12"
) do (
    reg export "HKCU\Software\Classes\%%~P" "%BK%\HKCU_%%~P.reg" /y >nul 2>&1
)

:: 备份 PowerPoint ProgID 关联
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
    reg export "HKCU\Software\Classes\%%~P" "%BK%\HKCU_%%~P.reg" /y >nul 2>&1
)

:: 备份 PreviewHandlers 列表和资源管理器设置
reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" "%BK%\HKLM_PreviewHandlers.reg" /y >nul 2>&1
reg export "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" "%BK%\HKLM_WOW6432Node_PreviewHandlers.reg" /y >nul 2>&1
reg export "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "%BK%\HKCU_Explorer_Advanced.reg" /y >nul 2>&1

echo 注册表备份完成。
echo.

echo 已跳过 Office 程序重新注册步骤。
echo 为避免弹出 Word / Excel / PowerPoint，本脚本不会执行以下命令：
echo   WINWORD.EXE /r
echo   EXCEL.EXE /regserver
echo   POWERPNT.EXE /regserver
echo.

echo 正在修复 Word 预览处理程序...

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

echo Word 预览处理程序修复完成。
echo.

echo 正在修复 Excel 预览处理程序...

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
"Excel.Template.12"
"Excel.TemplateMacroEnabled"
"Excel.TemplateMacroEnabled.12"
) do (
    reg add "HKCU\Software\Classes\%%~P\shellex\%PREVIEW_HANDLER%" /ve /t REG_SZ /d "%EXCEL_PREVIEW%" /f >nul
)

echo Excel 预览处理程序修复完成。
echo.

echo 正在修复 PowerPoint 预览处理程序...

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

echo PowerPoint 预览处理程序修复完成。
echo.

echo 正在注册 Office 预览器名称...

reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%WORD_PREVIEW%" /t REG_SZ /d "Microsoft Word previewer" /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%EXCEL_PREVIEW%" /t REG_SZ /d "Microsoft Excel previewer" /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%PPT_PREVIEW%" /t REG_SZ /d "Microsoft PowerPoint previewer" /f >nul

:: 兼容部分 32 位 Office / 64 位 Windows 场景
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%WORD_PREVIEW%" /t REG_SZ /d "Microsoft Word previewer" /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%EXCEL_PREVIEW%" /t REG_SZ /d "Microsoft Excel previewer" /f >nul 2>&1
reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "%PPT_PREVIEW%" /t REG_SZ /d "Microsoft PowerPoint previewer" /f >nul 2>&1

echo Office 预览器名称注册完成。
echo.

echo 正在开启资源管理器预览处理程序...

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowPreviewHandlers" /t REG_DWORD /d 1 /f >nul

echo 资源管理器预览处理程序已开启。
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
echo 请重新打开文件夹，测试 Word / Excel / PowerPoint 文件预览。
echo.
echo 如果没有看到预览窗格，请在资源管理器中开启：
echo   查看 / 显示 / 预览窗格
echo.
echo 备份文件已保存到：
echo %BK%
echo.
echo 如果仍然不能预览，建议：
echo.
echo 1. 重启电脑后再试；
echo 2. 确认 Microsoft Office 可以正常打开；
echo 3. 如果安装了 WPS，检查 WPS 是否接管了 Office 文件预览；
echo 4. 如果仍无效，可以尝试修复 Microsoft Office 安装。
echo.
pause
exit /b