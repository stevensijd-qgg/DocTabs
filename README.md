# DocTabs

轻量的 Office 文档标签栏：把 **Microsoft Word / Excel** 的多个窗口当作“一个窗口的多个标签”来切换，体验接近 WPS 的标签页。

*A lightweight WPS-style tab bar that groups Microsoft Word and Excel windows into one tabbed window. Single ~30 KB exe, ~22 MB RAM, no install, no SDK needed to build.*

## 功能

- **标签切换**：所有打开的 Word / Excel 文档显示为标签，点击切换，拖动排序，鼠标滚轮依次切换。
- **最大化时**：标签栏固定在屏幕顶部，最大化的文档窗口自动下移让出位置，标签与标题栏无缝贴合；其他程序（浏览器等）最大化时不受影响。
- **小窗口时**：标签贴在窗口标题栏正上方，随窗口移动；窗口顶到屏幕上沿放不下时自动隐藏，鼠标停在窗口上沿约 0.2 秒滑出。
- **整组联动**：一个文档最大化 / 还原 / 移动 / 缩放 / 最小化，其他文档同步为相同状态和位置。拖动窗口时其他文档暂时隐藏，松手后出现在同一位置。
- **关闭**
  - 标签上的 × 或中键：只关闭该文档
  - 窗口右上角 × 或 Alt+F4：整组文档依次关闭（未保存的会逐个弹出保存提示，点“取消”即停止）
  - 右键标签：关闭 / 关闭其他标签 / 关闭全部
- **全屏让位**：幻灯片放映、视频全屏等全屏程序在前台时自动隐藏。
- **托盘菜单**：开机自启动、退出。

只识别 Office 本体（`WINWORD.EXE` / `EXCEL.EXE`），不会收纳 WPS 的窗口。

## 下载与使用

1. 在 [Releases](https://github.com/stevensijd-qgg/DocTabs/releases) 下载 `DocTabs.exe`，放到任意文件夹。
2. 双击运行。打开 Word 或 Excel 文档后，标签栏会自动出现。
3. 如需开机自启动：右键任务栏托盘里的 DocTabs 图标 → 勾选“开机自启动”。
4. 退出：右键托盘图标 → 退出（会把调整过的窗口恢复原状）。

> 程序未做数字签名，首次运行时 Windows SmartScreen 可能提示“未知发布者”，点“更多信息 → 仍要运行”即可。也可以按下面的方法自己从源码编译。

**系统要求**：Windows 10 / 11，.NET Framework 4.7.2 及以上（Windows 10 1803 起系统自带），Microsoft Office 2013 及以上（每个文档单独一个窗口的版本）。

## 从源码编译

不需要安装 Visual Studio 或 .NET SDK，使用系统自带的 C# 编译器：

```bat
build.bat
```

会在同目录生成 `DocTabs.exe`。修改源码前请先在托盘里退出正在运行的 DocTabs，否则 exe 被占用无法覆盖。

源码只有一个文件 [`DocTabs.cs`](DocTabs.cs)（C# 5 / WinForms）。

## 实现说明

- 用 `SetWinEventHook` 监听前台切换、窗口创建/销毁、位置变化、最小化、拖动开始/结束等事件，`EnumWindows` 按窗口类（`OpusApp` / `XLMAIN`）和进程名识别文档窗口。
- 最大化：不使用 AppBar（不改变系统工作区），而是把最大化的文档窗口整体下移，并盖住 Office 最大化时顶部多留的一段空白。
- 标签栏是 `WS_EX_NOACTIVATE` 的工具窗口，点击不会抢走焦点；文档在前台时置顶，切到其他程序时只紧贴在文档窗口上一层。
- 浮动模式下标签以外的区域用 `TransparencyKey` 透明，可以点穿。
- 整组关闭通过低级鼠标/键盘钩子只识别“在当前文档的关闭按钮上松开左键”和“Alt+F4”两个动作，不记录其他输入。

## 已知限制

- 拖动窗口期间，其他文档的任务栏按钮会暂时消失，松手后恢复。
- 整组关闭依据鼠标是否点在 Office 标题栏右上角的关闭按钮区域判断，若 Office 界面布局特殊可能识别不到（此时只关闭当前文档）。
- 多显示器时，标签栏显示在当前文档所在的显示器上。

## 许可证

[MIT](LICENSE)
