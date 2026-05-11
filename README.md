# Office Preview Fix

一个用于修复 Windows 资源管理器中 Word、Excel、PowerPoint 文件预览功能的批处理工具。

适用于 Office 文件在资源管理器预览窗格中无法预览、被 WPS 或其他软件接管预览处理程序等情况。

## 支持修复的文件类型

### Word

- `.doc`
- `.docx`
- `.docm`
- `.dot`
- `.dotx`
- `.dotm`

### Excel

- `.xls`
- `.xlsx`
- `.xlsm`
- `.xlsb`
- `.xlt`
- `.xltx`
- `.xltm`

### PowerPoint

- `.ppt`
- `.pptx`
- `.pptm`
- `.pps`
- `.ppsx`
- `.ppsm`
- `.pot`
- `.potx`
- `.potm`

## 功能说明

该脚本会执行以下操作：

1. 自动备份相关注册表项到桌面；
2. 重新注册 Word、Excel、PowerPoint 组件；
3. 修复 Office 文件扩展名对应的预览处理程序；
4. 注册 Microsoft Office 预览器；
5. 开启 Windows 资源管理器预览处理程序；
6. 重启 Windows 资源管理器。

## 使用方法

1. 下载 `OfficePreviewFix.bat`；
2. 右键点击该文件；
3. 选择“以管理员身份运行”；
4. 等待脚本执行完成；
5. 重新打开资源管理器，测试 Office 文件预览。

## 重要提示

运行前建议关闭：

- Word
- Excel
- PowerPoint
- WPS
- 正在打开预览窗格的资源管理器窗口

如果运行后仍未生效，请重启电脑后再试。

## 预览处理程序 GUID

该脚本使用的 Office 预览器 GUID 如下：

```bat
Word Preview Handler:
{84F66100-FF7C-4FB4-B0C0-02CD7FB668FE}

Excel Preview Handler:
{00020827-0000-0000-C000-000000000046}

PowerPoint Preview Handler:
{65235197-874B-4A07-BDC5-E65EA825B718}

Windows Shell Preview Handler:
{8895b1c6-b41f-4c1c-a562-0d564250836f}