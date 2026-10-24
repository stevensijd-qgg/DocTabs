# DocTabs

轻量的 Office 文档标签栏：把 **Microsoft Word / Excel** 的多个窗口当作“一个窗口的多个标签”来切换，体验接近 WPS 的标签页。

*A lightweight WPS-style tab bar that groups Microsoft Word and Excel windows into one tabbed window. Single ~50 KB exe that also registers itself as an Office COM add-in so it starts with Word / Excel, ~22 MB RAM, no install, no SDK needed to build.*

## 功能

- **标签切换**：所有打开的 Word / Excel 文档显示为标签，点击切换，拖动排序，鼠标滚轮依次切换。
- **最大化时**：标签栏固定在屏幕顶部，最大化的文档窗口自动下移让出位置，标签与标题栏无缝贴合；其他程序（浏览器等）最大化时不受影响。
- **小窗口时**：标签贴在窗口标题栏正上方，随窗口移动；窗口太矮、上方又放不下时自动隐藏，鼠标停在窗口上沿约 0.2 秒滑出。
- **整组联动**：一个文档最大化 / 还原 / 移动 / 缩放 / 最小化，其他文档同步为相同状态和位置。拖动窗口时其他文档暂时挪到屏幕外，松手后出现在同一位置。
- **拆分 / 合并窗口**：把标签竖直拖出标签栏，它就拆成一个单独的窗口（有自己的标签栏），可以和原窗口并排放、用系统分屏；把一个窗口（拖它的标签或标题栏）拖到另一个标签栏上松手，就合并回去。右键标签也有“拆分为单独窗口”“合并所有窗口”。水平拖动标签是调整顺序。
- **拖标签栏移动窗口**：按住标签栏空白处（或窗口只有一个标签时按住标签）拖动，窗口立即跟着鼠标走；松手时鼠标顶到屏幕上沿则最大化，顶到左 / 右边缘则占半屏。最大化时往下拖会先还原。双击标签栏空白处最大化 / 还原。
- **系统分屏**：窗口贴到屏幕上沿（如 Win+← 分屏）时会自动往下让出标签栏的高度，左右两个窗口都显示各自的标签。
- **随 Office 自动启动**：DocTabs 同时是一个 Office COM 插件，打开 Word 或 Excel 时自动启动标签栏，Office 全部关闭后自动退出，不用常驻后台。
- **关闭**
  - 标签上的 × 或中键：只关闭该文档
  - 窗口右上角 × 或 Alt+F4：整组文档依次关闭（未保存的会逐个弹出保存提示，点“取消”即停止）
  - 右键标签：关闭 / 关闭其他标签 / 关闭全部
- **全屏让位**：幻灯片放映、视频全屏等全屏程序在前台时自动隐藏。
- **托盘菜单**：开机自启动、退出。

只识别 Office 本体（`WINWORD.EXE` / `EXCEL.EXE`），不会收纳 WPS 的窗口。

## 下载与使用

1. 在 [Releases](https://github.com/stevensijd-qgg/DocTabs/releases) 下载 `DocTabs.exe`，放到任意文件夹。
2. 双击运行一次。打开 Word 或 Excel 文档后，标签栏会自动出现。
3. 第一次运行时会自动登记为 Office 插件（只写当前用户的注册表，不需要管理员权限），以后直接打开 Word / Excel 就会自动带起 DocTabs。可在托盘菜单里取消“打开 Word / Excel 时自动启动”。
4. 也可以改用开机自启动：右键托盘图标 → 勾选“开机自启动”。
5. 退出：右键托盘图标 → 退出（会把调整过的窗口恢复原状）。

> 想挪动 `DocTabs.exe` 的位置：挪完后运行一次即可，插件登记会自动更新。想删除它：先在托盘菜单里取消“打开 Word / Excel 时自动启动”。插件也会出现在 Office 的“文件 → 选项 → 加载项 → COM 加载项”里，名称是“DocTabs 文档标签栏”。

> 打开 Word / Excel 没有自动出现标签栏时：先在托盘菜单里确认“打开 Word / Excel 时自动启动”已勾选（DocTabs 每次启动都会检查插件登记，Office 把插件停用过也会自动恢复）；仍不行的话，把 `%LOCALAPPDATA%\DocTabs\DocTabs.log` 发给作者，里面记录了插件加载、DocTabs 启动和退出的时间与原因。装有联想电脑管家、火绒等安全软件时，它们可能限制由 Office 启动的程序或会监听全局键盘的程序（DocTabs 用它识别 Alt+F4），可把 `DocTabs.exe` 加入安全软件的信任区。

> 程序未做数字签名，首次运行时 Windows SmartScreen 可能提示“未知发布者”，点“更多信息 → 仍要运行”即可。也可以按下面的方法自己从源码编译。

**系统要求**：Windows 10 / 11，.NET Framework 4.7.2 及以上（Windows 10 1803 起系统自带），Microsoft Office 2013 及以上（每个文档单独一个窗口的版本）。

## 从源码编译

不需要安装 Visual Studio 或 .NET SDK，使用系统自带的 C# 编译器：

```bat
build.bat
```

会在同目录生成 `DocTabs.exe`。重新编译前请先在托盘里退出正在运行的 DocTabs；Word / Excel 开着时 exe 也会被 Office 插件占用，可以先关掉 Office，或者把旧的 exe 改个名字再编译。

源码只有一个文件 [`DocTabs.cs`](DocTabs.cs)（C# 5 / WinForms）。

## 实现说明

- 用 `SetWinEventHook` 监听前台切换、窗口创建/销毁、位置变化、最小化、拖动开始/结束等事件，`EnumWindows` 按窗口类（`OpusApp` / `XLMAIN`）和进程名识别文档窗口。
- 最大化：不使用 AppBar（不改变系统工作区），而是把最大化的文档窗口整体下移，并盖住 Office 最大化时顶部多留的一段空白。
- 标签栏是 `WS_EX_NOACTIVATE` 的工具窗口，点击不会抢走焦点；文档在前台时置顶，切到其他程序时只紧贴在文档窗口上一层。
- 浮动模式下标签以外的区域用 `TransparencyKey` 透明，可以点穿。
- 整组关闭通过低级鼠标/键盘钩子只识别“在当前文档的关闭按钮上松开左键”和“Alt+F4”两个动作，不记录其他输入。
- Office 插件：`DocTabs.exe` 里有一个实现 `IDTExtensibility2` 的 COM 类，登记在 `HKCU\Software\Classes\CLSID` 和 `HKCU\Software\Microsoft\Office\{Word,Excel}\Addins` 下。Office 加载它时检查 DocTabs 是否在运行，没在运行就以 `/office` 参数启动；之后每 3 秒查一次（只是检查一个互斥量），DocTabs 中途没了会把它拉回来，用户从托盘主动退出的除外。以 `/office` 启动的 DocTabs 在 Word / Excel 进程全部退出几秒后自动退出。
- 逐显示器 DPI 感知（Per-Monitor V2）：主副屏缩放不同时，所有窗口坐标都是真实像素，标签栏高度和字体随文档所在屏幕的缩放变化。
- 分组：每组一个标签栏窗口，组内文档同步状态和位置；新打开的文档加入最近使用的那一组。
- 从标签栏拖动窗口：不用系统的拖动流程（按下发生在标签栏上，Office 的拖动流程常常要等松手才开始），而是由按下的标签栏一直捕获鼠标，每条鼠标移动都异步挪一次窗口；上一次还没生效就不发下一次，Office 忙时不会排队“补跑”。松手时自己处理贴边。拖动期间同组其他文档挪到当前显示器正上方的屏幕外：Excel 隐藏 / 显示一个窗口要几百毫秒，挪动只要几毫秒。
- 关闭文档一律发送 `WM_SYSCOMMAND / SC_CLOSE`（与点标题栏 × 相同）。`WM_CLOSE` 会让 Excel 退出整个程序，连带关掉所有工作簿。
- 插件登记时把 DocTabs 加入 Office 的“始终启用”列表（`Resiliency\DoNotDisableAddinList`），避免开机后首次启动较慢时被 Office 自动停用。

## 已知限制

- 从标签栏拖动时的贴边只有最大化和左右半屏，没有系统的分屏布局提示；需要时拖 Office 自己的标题栏即可。
- 整组关闭依据鼠标是否点在 Office 标题栏右上角的关闭按钮区域判断，若 Office 界面布局特殊可能识别不到（此时只关闭当前文档）。
- 多显示器时，标签栏显示在当前文档所在的显示器上。

## 许可证

[MIT](LICENSE)
