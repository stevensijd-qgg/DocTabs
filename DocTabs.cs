// DocTabs — 轻量文档标签栏：把 Office 的 Word / Excel 窗口当作“一个窗口的多个标签”来切换（不含 WPS）。
// 分组：默认所有文档在同一组；把标签竖直拖出标签栏即拆成单独的一组（单独的窗口），拖到别的标签栏上松手则合并回去。
//      每一组有自己的标签栏，系统分屏时左右两边都能看到各自的标签。
// 最大化：标签栏固定在屏幕顶部，最大化的文档窗口整体下移让出标签栏高度（其他程序不受影响）。
// 小窗口：标签栏贴在窗口标题栏正上方随窗口移动；窗口顶到屏幕上沿（如系统分屏）时往下让出标签栏的高度。
// 同组联动：一个文档最大化/还原/移动/最小化，其他文档同步为相同状态和位置；点窗口 × 或 Alt+F4 整组依次关闭。
// 拖标签栏空白处（或只有一个标签时拖标签）即拖动窗口，交给系统自己的拖动流程，与拖标题栏一样流畅，也支持贴边分屏。
// Office 插件：程序本身同时是一个 COM 加载项，Word / Excel 启动时自动拉起 DocTabs，Office 全部关闭后自动退出。
// 编译：build.bat（使用系统自带的 .NET Framework 4.x csc.exe，无需安装 SDK）
// https://github.com/stevensijd-qgg/DocTabs — MIT License

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("DocTabs")]
[assembly: System.Reflection.AssemblyProduct("DocTabs")]
[assembly: System.Reflection.AssemblyDescription("Lightweight WPS-style tab bar for Microsoft Word and Excel windows")]
[assembly: System.Reflection.AssemblyCopyright("Copyright (c) 2026 stevensijd-qgg, MIT License")]
[assembly: System.Reflection.AssemblyVersion("1.5.1.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.5.1.0")]
[assembly: ComVisible(false)]

namespace DocTabs
{
    static class Program
    {
        public const string MutexName = "DocTabs.SingleInstance.8F2C41";

        [STAThread]
        static void Main(string[] args)
        {
            bool created;
            using (var mutex = new Mutex(true, MutexName, out created))
            {
                if (!created) return;
                // 逐显示器 DPI 感知：主副屏缩放不同时，取到的窗口坐标才是真实像素，不会被系统按主屏缩放换算
                bool perMonitor = false;
                try { perMonitor = Native.SetProcessDpiAwarenessContext(new IntPtr(-4) /* PER_MONITOR_AWARE_V2 */); }
                catch (EntryPointNotFoundException) { }
                if (!perMonitor) Native.SetProcessDPIAware();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                // 由 Office 插件拉起时带 /office：Word / Excel 全部关闭后自动退出
                bool byOffice = Array.IndexOf(args, "/office") >= 0;
                var app = new Manager(byOffice);
                Application.Run();
                GC.KeepAlive(app);
            }
        }
    }

    // ---------------- Office 插件：Word / Excel 启动时把 DocTabs 一起拉起来 ----------------

    // custom 参数必须按 VARIANT 的 SAFEARRAY 封送，否则 Office 调用 OnConnection 时 CLR 会直接崩溃
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
                                      [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                                      [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode,
                                         [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // 运行在 Word / Excel 进程里，只做一件事：保证 DocTabs 在运行
    [ComVisible(true), Guid(OfficeAddin.Clsid), ProgId(OfficeAddin.ProgId), ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            OfficeAddin.LaunchDocTabs();
            OfficeAddin.StartWatch();
        }
        public void OnDisconnection(int removeMode, ref Array custom) { OfficeAddin.StopWatch(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }

    static class OfficeAddin
    {
        public const string Clsid = "5EBC09F9-99B3-4194-BA3C-1A0AC9DDD7F9";
        public const string ProgId = "DocTabs.Connect";
        public const string SettingsKey = @"Software\DocTabs";
        // 用户从托盘主动退出时写入；DocTabs 下次启动时清除
        public const string UserExitedValue = "UserExited";
        const string ClsidKey = @"Software\Classes\CLSID\{" + Clsid + "}";
        static readonly string[] Apps = { "Word", "Excel" };

        static System.Threading.Timer watch;
        static int watchLaunches;

        public static void LaunchDocTabs()
        {
            try { if (!IsRunning()) Start(); }
            catch { } // 任何问题都不能影响 Office 本身
        }

        // Office 一直开着、而 DocTabs 中途没了（被结束、崩溃）时，几秒内把它拉回来；用户从托盘主动退出的不拉
        public static void StartWatch()
        {
            try { if (watch == null) watch = new System.Threading.Timer(delegate { Watch(); }, null, 5000, 3000); }
            catch { }
        }

        public static void StopWatch()
        {
            try { if (watch != null) { watch.Dispose(); watch = null; } }
            catch { }
        }

        static void Watch()
        {
            try
            {
                if (watchLaunches >= 5 || IsRunning()) return;
                using (var k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                    if (k != null && Equals(k.GetValue(UserExitedValue), 1)) return;
                watchLaunches++;
                Start();
            }
            catch { }
        }

        static bool IsRunning()
        {
            Mutex m;
            if (!Mutex.TryOpenExisting(Program.MutexName, out m)) return false;
            m.Dispose();
            return true;
        }

        static void Start()
        {
            string exe = typeof(Connect).Assembly.Location;
            var psi = new System.Diagnostics.ProcessStartInfo(exe, "/office")
            {
                UseShellExecute = false,
                WorkingDirectory = System.IO.Path.GetDirectoryName(exe)
            };
            System.Diagnostics.Process.Start(psi);
        }

        static string CodeBase { get { return typeof(Connect).Assembly.CodeBase; } }
        static string AssemblyName { get { return typeof(Connect).Assembly.FullName; } }

        static string AddinKey(string app) { return @"Software\Microsoft\Office\" + app + @"\Addins\" + ProgId; }

        public static bool IsRegistered()
        {
            foreach (var app in Apps)
                using (var k = Registry.CurrentUser.OpenSubKey(AddinKey(app)))
                    if (k == null) return false;
            return true;
        }

        // 程序换了位置或升级了版本后，注册信息要跟着更新
        public static bool IsCurrent()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(ClsidKey + @"\InprocServer32"))
                return k != null && Equals(k.GetValue("CodeBase"), CodeBase) && Equals(k.GetValue("Assembly"), AssemblyName);
        }

        // 只写当前用户的注册表，不需要管理员权限；32 / 64 位 Office 都登记
        public static void Register()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    using (var k = root.CreateSubKey(@"Software\Classes\" + ProgId))
                    {
                        k.SetValue("", ProgId);
                        using (var c = k.CreateSubKey("CLSID")) c.SetValue("", "{" + Clsid + "}");
                    }
                    using (var k = root.CreateSubKey(ClsidKey))
                    {
                        k.SetValue("", ProgId);
                        using (var p = k.CreateSubKey("ProgId")) p.SetValue("", ProgId);
                        using (var srv = k.CreateSubKey("InprocServer32"))
                        {
                            srv.SetValue("", "mscoree.dll");
                            srv.SetValue("ThreadingModel", "Both");
                            srv.SetValue("Class", typeof(Connect).FullName);
                            srv.SetValue("Assembly", AssemblyName);
                            srv.SetValue("RuntimeVersion", "v4.0.30319");
                            srv.SetValue("CodeBase", CodeBase);
                        }
                    }
                }
            foreach (var app in Apps)
                using (var k = Registry.CurrentUser.CreateSubKey(AddinKey(app)))
                {
                    k.SetValue("FriendlyName", "DocTabs 文档标签栏");
                    k.SetValue("Description", "打开 Word / Excel 时自动启动 DocTabs 标签栏");
                    k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
        }

        public static void Unregister()
        {
            foreach (var app in Apps) Registry.CurrentUser.DeleteSubKeyTree(AddinKey(app), false);
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    root.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false);
                    root.DeleteSubKeyTree(ClsidKey, false);
                }
        }
    }

    enum DocKind { Word, Excel }

    enum BarMode { Hidden, Maximized, Follow, AutoHide }

    class DocTab
    {
        public IntPtr Handle;
        public DocKind Kind;
        public string Title;
        public string FullTitle;
        public RectangleF Bounds;
        public RectangleF CloseBounds;
    }

    // ---------------- 全局：事件钩子、窗口扫描、分组、托盘 ----------------

    class Manager
    {
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunName = "DocTabs";

        public readonly List<TabBar> Groups = new List<TabBar>();
        // 最近一次有文档到前台的分组；新打开的文档加入这一组
        public TabBar Active;
        // 最大化后被我们下移过的窗口 / 伪最大化的窗口 / 不允许移动最大化状态的窗口
        public readonly HashSet<IntPtr> Shifted = new HashSet<IntPtr>();
        public readonly HashSet<IntPtr> PseudoMax = new HashSet<IntPtr>();
        readonly HashSet<IntPtr> noShift = new HashSet<IntPtr>();
        // 拖动/缩放某组的当前文档时，被暂时隐藏的同组其他文档
        public readonly HashSet<IntPtr> HiddenByUs = new HashSet<IntPtr>();
        public TabBar DragGroup;
        DateTime dragStart;
        // 从标签栏开始的拖动：被拖的窗口、松手后要切到前台的窗口、是否由我们自己跟着鼠标挪
        IntPtr moveWin, activateAfterMove;
        bool movePosted, selfMove, moveSeen, moveBtnUp;
        int movePosts;
        Point selfMoveOffset, lastSelfPos;
        DateTime moveStarted, moveReleased, lastMovePost;
        readonly System.Windows.Forms.Timer moveWatch;

        // 用户点了某个文档窗口的 ×（或按了 Alt+F4）：该窗口真正关掉后，它所在的整组一起关闭
        IntPtr closeIntent;
        DateTime closeIntentTime;
        // 正在逐个关闭的文档队列
        readonly List<IntPtr> cascade = new List<IntPtr>();
        IntPtr cascadeCur;
        DateTime cascadeStepTime;

        readonly List<IntPtr> hooks = new List<IntPtr>();
        readonly Native.WinEventProc winEventProc;
        readonly Native.EnumWindowsProc enumProc;
        readonly Native.LowLevelProc mouseProc;
        readonly Native.LowLevelProc keyProc;
        IntPtr mouseHook, keyHook;
        readonly Dictionary<uint, string> procNames = new Dictionary<uint, string>();
        List<DocTab> scanResult;

        readonly System.Windows.Forms.Timer scanTimer;
        readonly System.Windows.Forms.Timer fallbackTimer;
        readonly System.Windows.Forms.Timer revealTimer;
        readonly NotifyIcon tray;
        readonly MenuItem autostartItem;
        readonly MenuItem addinItem;
        readonly bool byOffice;
        int officeGoneTicks;
        bool exiting;

        public Manager(bool byOffice)
        {
            this.byOffice = byOffice;
            ClearUserExited();

            scanTimer = new System.Windows.Forms.Timer { Interval = 120 };
            scanTimer.Tick += delegate { scanTimer.Stop(); Scan(); };
            fallbackTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            fallbackTimer.Tick += delegate
            {
                // 兜底：漏掉了“拖动结束”通知时，别让文档一直处于隐藏状态
                if (DragGroup != null && (DateTime.Now - dragStart).TotalSeconds > 30) EndWindowDrag();
                if (DragGroup == null && HiddenByUs.Count > 0 && !revealTimer.Enabled) RevealHidden();
                Scan();
                if (this.byOffice) CheckOfficeGone();
            };
            fallbackTimer.Start();
            // 松手后稍等当前文档的最终位置稳定，再把其他文档放到同一位置并显示出来
            revealTimer = new System.Windows.Forms.Timer { Interval = 80 };
            revealTimer.Tick += delegate { revealTimer.Stop(); RevealHidden(); };
            moveWatch = new System.Windows.Forms.Timer { Interval = 10 };
            moveWatch.Tick += delegate { MoveWatchTick(); };

            enumProc = EnumCallback;
            winEventProc = OnWinEvent;
            const uint flags = Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS;
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_FOREGROUND, Native.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MOVESIZESTART, Native.EVENT_SYSTEM_MOVESIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MINIMIZESTART, Native.EVENT_SYSTEM_MINIMIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_DESTROY, Native.EVENT_OBJECT_HIDE, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_LOCATIONCHANGE, Native.EVENT_OBJECT_NAMECHANGE, IntPtr.Zero, winEventProc, 0, 0, flags));
            // Office 新打开的窗口加载期间是隐身（cloaked）的，显示出来时要立刻收进来
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_CLOAKED, Native.EVENT_OBJECT_UNCLOAKED, IntPtr.Zero, winEventProc, 0, 0, flags));

            tray = new NotifyIcon { Icon = MakeTrayIcon(Native.MonitorScale(IntPtr.Zero)), Text = "DocTabs 文档标签栏", Visible = true };
            autostartItem = new MenuItem("开机自启动", delegate { ToggleAutostart(); }) { Checked = IsAutostart() };
            addinItem = new MenuItem("打开 Word / Excel 时自动启动", delegate { ToggleAddin(); });
            tray.ContextMenu = new ContextMenu(new[]
            {
                new MenuItem("DocTabs 文档标签栏") { Enabled = false },
                new MenuItem("-"),
                addinItem,
                autostartItem,
                new MenuItem("退出", delegate { ExitApp(true); })
            });
            InitAddin();

            SystemEvents.DisplaySettingsChanged += OnDisplayChanged;

            // 只记录“在当前文档的关闭按钮上松开左键”和“Alt+F4”这两个动作，回调里不做别的事
            mouseProc = OnMouseLL;
            keyProc = OnKeyLL;
            IntPtr mod = Native.GetModuleHandle(null);
            mouseHook = Native.SetWindowsHookEx(Native.WH_MOUSE_LL, mouseProc, mod, 0);
            keyHook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, keyProc, mod, 0);
            Scan();
        }

        // ---------------- 分组 ----------------

        public TabBar GroupOf(IntPtr h)
        {
            if (h == IntPtr.Zero) return null;
            foreach (var g in Groups) if (g.IndexOf(h) >= 0) return g;
            return null;
        }

        public TabBar NewGroup()
        {
            var g = new TabBar(this);
            Groups.Add(g);
            return g;
        }

        void RemoveGroup(TabBar g)
        {
            Groups.Remove(g);
            if (Active == g) Active = null;
            if (DragGroup == g) DragGroup = null;
            g.Destroy();
        }

        bool IsOnAnyBar(int x, int y)
        {
            foreach (var g in Groups) if (g.Visible && g.Bounds.Contains(x, y)) return true;
            return false;
        }

        public void UpdateAll()
        {
            foreach (var g in Groups.ToArray()) g.UpdatePlacement();
        }

        // 把 src 组的文档全部并入 dst 组，src 的当前文档成为 dst 的当前文档
        public void Merge(TabBar src, TabBar dst)
        {
            IntPtr cur = src.CurDoc;
            var moved = new List<DocTab>(src.Tabs);
            src.Tabs.Clear();
            RemoveGroup(src);
            foreach (var t in moved) dst.Tabs.Add(t);
            Active = dst;
            foreach (var t in moved)
                if (t.Handle != cur && !HiddenByUs.Contains(t.Handle)) dst.ApplyGroupState(t.Handle, false);
            if (cur == IntPtr.Zero && moved.Count > 0) cur = moved[0].Handle;
            dst.LayoutTabs();
            if (cur != IntPtr.Zero) dst.ActivateDoc(cur);
            UpdateAll();
        }

        // 拖动结束时鼠标落在另一组的标签上：并进那一组。拖到屏幕最上沿是系统的“最大化”手势，不算
        void TryMergeOnDrop(TabBar src)
        {
            if (src == null || !Groups.Contains(src)) return;
            Point c = Cursor.Position;
            Rectangle mon = Screen.FromPoint(c).Bounds;
            if (c.Y <= mon.Top + 2) return;
            foreach (var g in Groups)
                if (g != src && g.Visible && g.TabStripContains(c)) { Merge(src, g); return; }
        }

        // 交给系统的拖动流程移动窗口（与拖标题栏完全相同：流畅、可贴边分屏）。调用时左键必须仍按着。
        // 按着鼠标时不能把窗口切到前台：Office 会等到松手才开始处理拖动，窗口就在松手后一直粘着鼠标。
        // 所以只把窗口提到最上面，松手后再切到前台。前台是另一个程序（线程）时系统的拖动拿不到鼠标，
        // 系统的拖动迟迟没开始时也一样，改由我们自己跟着鼠标挪
        public void StartNativeMove(IntPtr h)
        {
            Native.ReleaseCapture();
            Point c = Cursor.Position;
            IntPtr fg = Native.GetForegroundWindow();
            if (fg != h)
            {
                Native.SetWindowPos(h, Native.HWND_TOP, 0, 0, 0, 0, Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE);
                activateAfterMove = h;
            }
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            selfMoveOffset = new Point(r.Left - c.X, r.Top - c.Y);
            lastSelfPos = new Point(r.Left, r.Top);
            moveWin = h;
            moveSeen = false;
            moveBtnUp = false;
            selfMove = false;
            movePosts = 0;
            moveStarted = DateTime.Now;
            movePosted = fg == h || ThreadOf(fg) == ThreadOf(h);
            if (movePosted)
                Native.PostMessage(h, Native.WM_SYSCOMMAND, (IntPtr)(Native.SC_MOVE | Native.HTCAPTION),
                    (IntPtr)((c.Y << 16) | (c.X & 0xFFFF)));
            else StartSelfMove(h);
            moveWatch.Start();
        }

        void StartSelfMove(IntPtr h)
        {
            selfMove = true;
            var g = GroupOf(h);
            if (g != null && g.CurDoc == h && DragGroup != g) BeginWindowDrag(g);
        }

        static uint ThreadOf(IntPtr h)
        {
            uint pid;
            return h == IntPtr.Zero ? 0 : Native.GetWindowThreadProcessId(h, out pid);
        }

        // 窗口 h 是否正处在系统的拖动/缩放流程里
        static bool InMoveLoop(IntPtr h)
        {
            var gi = new Native.GUITHREADINFO { cbSize = Marshal.SizeOf(typeof(Native.GUITHREADINFO)) };
            return Native.GetGUIThreadInfo(ThreadOf(h), ref gi) && (gi.flags & Native.GUI_INMOVESIZE) != 0 && gi.hwndMoveSize == h;
        }

        // 结束窗口 h 上多余的系统拖动：先补发松开左键，不行再按回车（窗口停在当前位置）
        void EndMoveLoop(IntPtr h)
        {
            if ((DateTime.Now - lastMovePost).TotalMilliseconds < 150) return;
            lastMovePost = DateTime.Now;
            bool hard = movePosts++ >= 2;
            Native.PostMessage(h, hard ? Native.WM_KEYDOWN : Native.WM_LBUTTONUP,
                hard ? (IntPtr)Native.VK_RETURN : IntPtr.Zero, IntPtr.Zero);
        }

        // 拖动期间每 10ms 一次：自己挪窗口时跟着鼠标走；松手后系统的拖动若还没结束（窗口会一直粘着鼠标），替它结束
        void MoveWatchTick()
        {
            IntPtr h = moveWin;
            if (h == IntPtr.Zero || !Native.IsWindow(h)) { StopMoveWatch(); return; }
            int vk = SystemInformation.MouseButtonsSwapped ? Native.VK_RBUTTON : Native.VK_LBUTTON;
            bool down = (Native.GetAsyncKeyState(vk) & 0x8000) != 0;
            bool inLoop = movePosted && InMoveLoop(h);
            if (down && !moveBtnUp)
            {
                if (selfMove)
                {
                    // 系统的拖动迟到了，而我们已经在挪：结束它
                    if (inLoop) { moveSeen = true; EndMoveLoop(h); }
                    Point c = Cursor.Position;
                    var p = new Point(c.X + selfMoveOffset.X, c.Y + selfMoveOffset.Y);
                    if (p != lastSelfPos)
                    {
                        lastSelfPos = p;
                        Native.SetWindowPos(h, IntPtr.Zero, p.X, p.Y, 0, 0,
                            Native.SWP_NOSIZE | Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_ASYNCWINDOWPOS);
                    }
                }
                else if (inLoop) moveSeen = true;
                else if (!moveSeen && (DateTime.Now - moveStarted).TotalMilliseconds > 150) StartSelfMove(h);
                return;
            }
            // 松手后又按下：用户开始了新的操作，不再管
            if (down) { StopMoveWatch(); return; }
            if (!moveBtnUp)
            {
                moveBtnUp = true;
                moveReleased = DateTime.Now;
                if (selfMove)
                {
                    selfMove = false;
                    if (DragGroup != null) EndWindowDrag();
                    else TryMergeOnDrop(GroupOf(h));
                    ActivateAfterMove();
                }
            }
            double ms = (DateTime.Now - moveReleased).TotalMilliseconds;
            if (inLoop)
            {
                moveSeen = true;
                if (ms > 80) EndMoveLoop(h);
                return;
            }
            // 系统的拖动已经结束；或者还没开始——再等一会儿，防止它在松手后才开始
            if (moveSeen || !movePosted || ms > 2000) StopMoveWatch();
        }

        void StopMoveWatch()
        {
            moveWatch.Stop();
            moveWin = IntPtr.Zero;
            selfMove = false;
            ActivateAfterMove();
        }

        // 拖完了：把被拖的窗口切到前台
        void ActivateAfterMove()
        {
            IntPtr h = activateAfterMove;
            activateAfterMove = IntPtr.Zero;
            if (h == IntPtr.Zero || !Native.IsWindow(h) || Native.IsIconic(h) || Native.GetForegroundWindow() == h) return;
            BringToFront(h);
        }

        // 把窗口切到前台（绕过系统的前台锁）
        public static void BringToFront(IntPtr h)
        {
            // 发送一个空鼠标输入以解除系统的前台锁（PowerToys 同款做法），否则 SetForegroundWindow 可能只会闪烁任务栏
            var input = new Native.INPUT { type = Native.INPUT_MOUSE };
            Native.SendInput(1, new[] { input }, Marshal.SizeOf(typeof(Native.INPUT)));
            Native.SetForegroundWindow(h);
            if (Native.GetForegroundWindow() != h)
            {
                uint dummy;
                uint fgThread = Native.GetWindowThreadProcessId(Native.GetForegroundWindow(), out dummy);
                uint cur = Native.GetCurrentThreadId();
                Native.AttachThreadInput(cur, fgThread, true);
                Native.BringWindowToTop(h);
                Native.SetForegroundWindow(h);
                Native.AttachThreadInput(cur, fgThread, false);
            }
            // 最后兜底：Alt+Tab 使用的切换接口
            if (Native.GetForegroundWindow() != h) Native.SwitchToThisWindow(h, true);
        }

        // ---------------- 关闭整组 ----------------

        IntPtr OnMouseLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_LBUTTONUP)
            {
                var info = (Native.MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.MSLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                var g = GroupOf(fg);
                if (g != null && fg == g.CurDoc && !IsOnAnyBar(info.pt.X, info.pt.Y) && IsOverCloseButton(fg, info.pt.X, info.pt.Y))
                    SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        IntPtr OnKeyLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_SYSKEYDOWN)
            {
                var info = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                if (info.vkCode == 0x73 /* F4 */ && (info.flags & 0x20 /* ALT */) != 0 && GroupOf(fg) != null) SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        // Office 标题栏右上角的关闭按钮区域
        static bool IsOverCloseButton(IntPtr h, int x, int y)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            float ws = Native.MonitorScale(h);
            return x >= f.Right - 52 * ws && x <= f.Right && y >= f.Top && y <= f.Top + 56 * ws;
        }

        void SetCloseIntent(IntPtr h)
        {
            closeIntent = h;
            closeIntentTime = DateTime.Now;
        }

        public void StartCascade(IEnumerable<IntPtr> handles)
        {
            cascade.Clear();
            cascade.AddRange(handles);
            CascadeNext();
        }

        // 一次只关一个，前一个真正关掉后再关下一个，保存提示会依次弹出
        void CascadeNext()
        {
            cascadeCur = IntPtr.Zero;
            while (cascade.Count > 0)
            {
                IntPtr h = cascade[0];
                cascade.RemoveAt(0);
                if (!Native.IsWindow(h) || GroupOf(h) == null) continue;
                cascadeCur = h;
                cascadeStepTime = DateTime.Now;
                Native.PostMessage(h, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                return;
            }
        }

        void StopCascade()
        {
            cascade.Clear();
            cascadeCur = IntPtr.Zero;
        }

        // ---------------- 窗口扫描 ----------------

        public void ScheduleScan()
        {
            scanTimer.Stop();
            scanTimer.Start();
        }

        public void Scan()
        {
            if (exiting) return;
            scanResult = new List<DocTab>();
            Native.EnumWindows(enumProc, IntPtr.Zero);
            var found = new Dictionary<IntPtr, DocTab>();
            foreach (var t in scanResult) found[t.Handle] = t;

            var changed = new HashSet<TabBar>();
            foreach (var g in Groups.ToArray())
            {
                for (int i = g.Tabs.Count - 1; i >= 0; i--)
                {
                    var t = g.Tabs[i];
                    DocTab now;
                    if (!found.TryGetValue(t.Handle, out now))
                    {
                        g.Tabs.RemoveAt(i);
                        changed.Add(g);
                        if (t.Handle == closeIntent)
                        {
                            // 用户关掉了整个窗口：同组其余文档依次关闭
                            closeIntent = IntPtr.Zero;
                            var rest = new List<IntPtr>();
                            foreach (var o in g.Tabs) rest.Add(o.Handle);
                            StartCascade(rest);
                        }
                        continue;
                    }
                    if (now.FullTitle != t.FullTitle)
                    {
                        t.Title = now.Title;
                        t.FullTitle = now.FullTitle;
                        changed.Add(g);
                    }
                    found.Remove(t.Handle);
                }
            }
            if (found.Count > 0)
            {
                // 新打开的文档加入最近使用的那一组
                TabBar target = Active != null && Groups.Contains(Active) ? Active : Groups.Count > 0 ? Groups[0] : NewGroup();
                // EnumWindows 按 Z 序从上到下，倒序追加让较早的窗口排在前面
                for (int i = scanResult.Count - 1; i >= 0; i--)
                    if (found.ContainsKey(scanResult[i].Handle)) { target.Tabs.Add(scanResult[i]); changed.Add(target); }
            }
            foreach (var g in Groups.ToArray()) if (g.Tabs.Count == 0) RemoveGroup(g);

            var live = new HashSet<uint>();
            foreach (var t in scanResult) { uint pid; Native.GetWindowThreadProcessId(t.Handle, out pid); live.Add(pid); }
            foreach (var pid in new List<uint>(procNames.Keys)) if (!live.Contains(pid)) procNames.Remove(pid);

            // 离开分组的窗口（如文档已转到另一个桌面）恢复正常的最大化位置
            foreach (var h in new List<IntPtr>(Shifted))
                if (GroupOf(h) == null) { Shifted.Remove(h); if (Native.IsWindow(h)) SetFrameTop(h, 0); }
            PseudoMax.RemoveWhere(h => GroupOf(h) == null);
            noShift.RemoveWhere(h => GroupOf(h) == null);
            foreach (var g in Groups)
                foreach (var t in g.Tabs) if (!PseudoMax.Contains(t.Handle)) EnsureShift(t.Handle);

            if (cascadeCur != IntPtr.Zero && GroupOf(cascadeCur) == null) CascadeNext();
            // 长时间没关掉（保存提示一直没处理）就放弃，避免之后误关
            if (cascadeCur != IntPtr.Zero && (DateTime.Now - cascadeStepTime).TotalMinutes > 3) StopCascade();
            if (closeIntent != IntPtr.Zero && (DateTime.Now - closeIntentTime).TotalMinutes > 3) closeIntent = IntPtr.Zero;

            IntPtr fg = Native.GetForegroundWindow();
            var fgGroup = GroupOf(fg);
            if (fgGroup != null) Active = fgGroup;
            foreach (var g in Groups.ToArray()) g.AfterScan(scanResult, fg, changed.Contains(g));
        }

        bool EnumCallback(IntPtr h, IntPtr l)
        {
            if (!Native.IsWindowVisible(h) && !HiddenByUs.Contains(h)) return true;
            if (Native.GetWindow(h, Native.GW_OWNER) != IntPtr.Zero) return true;
            DocKind kind;
            if (!TryKind(Native.GetClassNameOf(h), out kind)) return true;
            if (Native.IsCloaked(h)) return true;
            string title = Native.GetWindowTextOf(h);
            if (string.IsNullOrEmpty(title)) return true;
            // 只认 Office 本体，WPS 等兼容程序可能使用相同的窗口类名
            string exe = kind == DocKind.Word ? "WINWORD.EXE" : "EXCEL.EXE";
            if (!string.Equals(ProcessName(h), exe, StringComparison.OrdinalIgnoreCase)) return true;
            scanResult.Add(new DocTab { Handle = h, Kind = kind, Title = CleanTitle(title), FullTitle = title });
            return true;
        }

        static bool TryKind(string cls, out DocKind kind)
        {
            kind = DocKind.Word;
            switch (cls)
            {
                case "OpusApp": kind = DocKind.Word; return true;
                case "XLMAIN": kind = DocKind.Excel; return true;
            }
            return false;
        }

        static string CleanTitle(string t)
        {
            // "报告.docx  -  兼容模式 - Word" → "报告.docx"
            int last = t.LastIndexOf(" - ", StringComparison.Ordinal);
            if (last > 0) t = t.Substring(0, last);
            int mode = t.LastIndexOf("  -  ", StringComparison.Ordinal);
            if (mode > 0) t = t.Substring(0, mode);
            return t.Trim();
        }

        string ProcessName(IntPtr h)
        {
            uint pid;
            Native.GetWindowThreadProcessId(h, out pid);
            string name;
            if (!procNames.TryGetValue(pid, out name))
            {
                name = System.IO.Path.GetFileName(Native.GetProcessPath(pid) ?? "");
                procNames[pid] = name;
            }
            return name;
        }

        void OnWinEvent(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            if (exiting || hwnd == IntPtr.Zero || idObject != Native.OBJID_WINDOW || idChild != 0) return;
            TabBar g;
            switch (ev)
            {
                case Native.EVENT_SYSTEM_MOVESIZESTART:
                    g = GroupOf(hwnd);
                    if (g != null && hwnd == g.CurDoc) BeginWindowDrag(g);
                    return;
                case Native.EVENT_SYSTEM_MOVESIZEEND:
                    // 我们自己挪窗口时，迟到的系统拖动被结束掉不算松手
                    if (DragGroup != null && !selfMove) EndWindowDrag();
                    return;
                case Native.EVENT_OBJECT_LOCATIONCHANGE:
                    g = GroupOf(hwnd);
                    if (g != null) g.OnDocLocation(hwnd);
                    return;
                case Native.EVENT_SYSTEM_FOREGROUND:
                    // 保存提示里点了“取消”，文档重新回到前台：不再整组关闭
                    if (hwnd == closeIntent && (DateTime.Now - closeIntentTime).TotalMilliseconds > 400) closeIntent = IntPtr.Zero;
                    if (hwnd == cascadeCur && (DateTime.Now - cascadeStepTime).TotalMilliseconds > 1500) StopCascade();
                    g = GroupOf(hwnd);
                    if (g != null)
                    {
                        Active = g;
                        g.OnDocForeground(hwnd);
                    }
                    else
                    {
                        // 刚打开的文档还没收进来：马上扫描，别让它以未下移的最大化状态盖住标签栏
                        DocKind fgKind;
                        if (TryKind(Native.GetClassNameOf(hwnd), out fgKind)) { Scan(); return; }
                    }
                    UpdateAll();
                    ScheduleScan();
                    return;
                case Native.EVENT_SYSTEM_MINIMIZESTART:
                    // 组内任一文档被最小化就全组最小化（前台可能已先切到另一个文档，不能只看当前文档）
                    g = GroupOf(hwnd);
                    if (g != null) { g.MinimizeOthers(hwnd); g.UpdatePlacement(); }
                    return;
                case Native.EVENT_OBJECT_DESTROY:
                    if (GroupOf(hwnd) != null) ScheduleScan();
                    return;
            }
            if (Native.GetAncestor(hwnd, Native.GA_ROOT) != hwnd) return;
            DocKind kind;
            if (TryKind(Native.GetClassNameOf(hwnd), out kind)) ScheduleScan();
        }

        // ---------------- 拖动窗口时的同组处理 ----------------

        void BeginWindowDrag(TabBar g)
        {
            DragGroup = g;
            dragStart = DateTime.Now;
            revealTimer.Stop();
            foreach (var t in g.Tabs)
            {
                IntPtr o = t.Handle;
                if (o == g.CurDoc || Native.IsIconic(o) || !Native.IsWindowVisible(o)) continue;
                HiddenByUs.Add(o);
                Native.ShowWindowAsync(o, Native.SW_HIDE);
            }
        }

        void EndWindowDrag()
        {
            var g = DragGroup;
            DragGroup = null;
            revealTimer.Stop();
            revealTimer.Start();
            TryMergeOnDrop(g);
        }

        void RevealHidden()
        {
            if (HiddenByUs.Count == 0) return;
            var touched = new HashSet<TabBar>();
            foreach (var o in new List<IntPtr>(HiddenByUs))
            {
                if (!Native.IsWindow(o)) continue;
                var g = GroupOf(o);
                if (g != null && g.HasState && !g.GroupMax)
                {
                    // 先摆到与当前文档相同的位置，再显示
                    var wp = Native.GetPlacement(o);
                    wp.flags = 0;
                    wp.showCmd = Native.SW_SHOWNOACTIVATE;
                    wp.rcNormalPosition = g.GroupNormal;
                    Native.SetWindowPlacement(o, ref wp);
                    Shifted.Remove(o);
                    PseudoMax.Remove(o);
                    touched.Add(g);
                }
                else
                {
                    Native.ShowWindow(o, Native.SW_SHOWNOACTIVATE);
                }
                // 放在当前文档后面，不要盖住它
                if (g != null && g.CurDoc != IntPtr.Zero && g.CurDoc != o && Native.IsWindow(g.CurDoc))
                    Native.SetWindowPos(o, g.CurDoc, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            }
            HiddenByUs.Clear();
            // 跨到缩放不同的屏幕时，Office 会按新 DPI 自行调整窗口大小，稍后再校正一次
            foreach (var g in touched) g.ScheduleSync();
            ScheduleScan();
        }

        // ---------------- 最大化窗口下移 ----------------

        // 最大化的文档窗口整体下移，把顶部让给标签栏
        public void EnsureShift(IntPtr h)
        {
            if (!Native.IsZoomed(h)) return;
            int bh = BarHeightFor(h);
            if (!SetFrameTop(h, bh)) return;
            Shifted.Add(h);
            // 个别程序不允许移动最大化窗口，改为“伪最大化”：还原后铺满标签栏以下的区域
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            if (f.Top < Screen.FromHandle(h).WorkingArea.Top + bh / 2)
            {
                noShift.Add(h);
                PseudoMaximize(h);
            }
        }

        public bool IsNoShift(IntPtr h) { return noShift.Contains(h); }

        Rectangle PseudoMaxFrame(IntPtr h)
        {
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            return Rectangle.FromLTRB(wa.Left, wa.Top + BarHeightFor(h), wa.Right, wa.Bottom);
        }

        // 窗口所在显示器上的标签栏高度
        public static int BarHeightFor(IntPtr h)
        {
            return (int)Math.Round(30 * Native.MonitorScale(h));
        }

        public void PseudoMaximize(IntPtr h)
        {
            Shifted.Remove(h);
            if (Native.IsZoomed(h) || Native.IsIconic(h)) Native.ShowWindow(h, Native.SW_SHOWNOACTIVATE);
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle want = PseudoMaxFrame(h);
            // 加上程序自身的隐形边框，让可见边框正好对齐
            int l = want.Left - (f.Left - r.Left), t = want.Top - (f.Top - r.Top);
            int rr = want.Right + (r.Right - f.Right), b = want.Bottom + (r.Bottom - f.Bottom);
            Native.SetWindowPos(h, IntPtr.Zero, l, t, rr - l, b - t,
                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            PseudoMax.Add(h);
        }

        public bool IsPseudoMax(IntPtr h)
        {
            if (!PseudoMax.Contains(h)) return false;
            if (!Native.IsZoomed(h) && !Native.IsIconic(h))
            {
                Native.RECT r;
                Native.GetWindowRect(h, out r);
                Native.RECT f = Native.FrameBounds(h, r);
                Rectangle want = PseudoMaxFrame(h);
                if (Math.Abs(f.Left - want.Left) <= 3 && Math.Abs(f.Top - want.Top) <= 3 &&
                    Math.Abs(f.Right - want.Right) <= 3 && Math.Abs(f.Bottom - want.Bottom) <= 3)
                    return true;
            }
            // 被用户拖动/缩放过，或被真正最大化了，不再算伪最大化
            PseudoMax.Remove(h);
            return false;
        }

        public bool IsMax(IntPtr h)
        {
            return Native.IsZoomed(h) || IsPseudoMax(h);
        }

        // 把最大化窗口的可见上边沿放到“工作区顶部 + offset”，底边不动；返回是否处于偏移状态
        bool SetFrameTop(IntPtr h, int offset)
        {
            if (!Native.IsZoomed(h)) return false;
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            // 按工作区算出绝对目标位置（可见边框上沿 = 工作区顶 + offset，下沿 = 工作区底），重复调用结果不变。
            // Office 最大化时会在顶部再留一段与隐形边框等高的空白（平时在屏幕外），下移时让标签栏把它也盖住，避免出现缝隙
            int border = f.Top - r.Top;
            int pad = offset != 0 ? border : 0;
            int top = wa.Top + offset - pad - border;
            int bottom = wa.Bottom + (r.Bottom - f.Bottom);
            if (Math.Abs(r.Top - top) > 1 || Math.Abs(r.Bottom - bottom) > 1)
                Native.SetWindowPos(h, IntPtr.Zero, r.Left, top, r.Right - r.Left, bottom - top,
                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            return offset != 0;
        }

        void RestoreShifted()
        {
            foreach (var h in Shifted) if (Native.IsWindow(h)) SetFrameTop(h, 0);
            Shifted.Clear();
        }

        void OnDisplayChanged(object sender, EventArgs e)
        {
            foreach (var g in Groups) foreach (var t in g.Tabs) EnsureShift(t.Handle);
            UpdateAll();
        }

        // ---------------- 托盘 / 自启动 / 插件 / 退出 ----------------

        static bool IsAutostart()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(RunKey))
                return k != null && k.GetValue(RunName) != null;
        }

        void ToggleAutostart()
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (IsAutostart()) k.DeleteValue(RunName, false);
                else k.SetValue(RunName, "\"" + Application.ExecutablePath + "\"");
            }
            autostartItem.Checked = IsAutostart();
        }

        static void ClearUserExited()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(OfficeAddin.SettingsKey, true))
                    if (k != null) k.DeleteValue(OfficeAddin.UserExitedValue, false);
            }
            catch { }
        }

        // 第一次运行时默认登记为 Office 插件；之后尊重用户在托盘里的选择，只在程序挪了位置或升级后刷新登记信息
        void InitAddin()
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(OfficeAddin.SettingsKey))
                {
                    if (k.GetValue("AddinOffered") == null)
                    {
                        k.SetValue("AddinOffered", 1, RegistryValueKind.DWord);
                        OfficeAddin.Register();
                        tray.ShowBalloonTip(5000, "DocTabs",
                            "已登记为 Office 插件，以后打开 Word / Excel 会自动启动标签栏。可在托盘菜单中关闭。", ToolTipIcon.Info);
                    }
                    else if (OfficeAddin.IsRegistered() && !OfficeAddin.IsCurrent())
                    {
                        OfficeAddin.Register();
                    }
                }
            }
            catch { }
            addinItem.Checked = OfficeAddin.IsRegistered();
        }

        void ToggleAddin()
        {
            try
            {
                if (OfficeAddin.IsRegistered()) OfficeAddin.Unregister();
                else OfficeAddin.Register();
            }
            catch (Exception ex)
            {
                MessageBox.Show("修改 Office 插件登记失败：" + ex.Message, "DocTabs", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            addinItem.Checked = OfficeAddin.IsRegistered();
        }

        void CheckOfficeGone()
        {
            bool any = false;
            foreach (var name in new[] { "WINWORD", "EXCEL" })
            {
                foreach (var p in System.Diagnostics.Process.GetProcessesByName(name))
                {
                    // 崩溃后残留、已无法访问的进程不算
                    try { if (!p.HasExited) any = true; } catch { }
                    p.Dispose();
                }
            }
            officeGoneTicks = any ? 0 : officeGoneTicks + 1;
            if (officeGoneTicks >= 3) ExitApp(false);
        }

        static Icon MakeTrayIcon(float s)
        {
            int size = Math.Max(16, (int)Math.Round(16 * s));
            using (var bmp = new Bitmap(size, size))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    float u = size / 16f;
                    using (var back = new SolidBrush(Color.FromArgb(120, 130, 145)))
                        g.FillRectangle(back, 1 * u, 3 * u, 6 * u, 4 * u);
                    using (var front = new SolidBrush(Color.FromArgb(24, 90, 189)))
                    {
                        g.FillRectangle(front, 7 * u, 3 * u, 7 * u, 4 * u);
                        g.FillRectangle(front, 1 * u, 7 * u, 14 * u, 7 * u);
                    }
                }
                IntPtr hicon = bmp.GetHicon();
                var icon = (Icon)Icon.FromHandle(hicon).Clone();
                Native.DestroyIcon(hicon);
                return icon;
            }
        }

        // byUser：用户从菜单主动退出，Office 插件在本次 Office 运行期间不再自动把它拉起来
        public void ExitApp(bool byUser)
        {
            if (exiting) return;
            exiting = true;
            if (byUser)
            {
                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(OfficeAddin.SettingsKey))
                        k.SetValue(OfficeAddin.UserExitedValue, 1, RegistryValueKind.DWord);
                }
                catch { }
            }
            fallbackTimer.Stop();
            scanTimer.Stop();
            revealTimer.Stop();
            moveWatch.Stop();
            foreach (var h in hooks) if (h != IntPtr.Zero) Native.UnhookWinEvent(h);
            if (mouseHook != IntPtr.Zero) Native.UnhookWindowsHookEx(mouseHook);
            if (keyHook != IntPtr.Zero) Native.UnhookWindowsHookEx(keyHook);
            mouseHook = keyHook = IntPtr.Zero;
            hooks.Clear();
            SystemEvents.DisplaySettingsChanged -= OnDisplayChanged;
            DragGroup = null;
            RevealHidden();
            RestoreShifted();
            foreach (var g in Groups.ToArray()) g.Destroy();
            Groups.Clear();
            tray.Visible = false;
            tray.Dispose();
            Application.Exit();
        }
    }

    // ---------------- 一个分组：一个标签栏窗口 + 组内文档的联动 ----------------

    class TabBar : Form
    {
        static readonly Color BarBg = Color.FromArgb(222, 226, 231);
        static readonly Color BarLine = Color.FromArgb(200, 205, 212);
        static readonly Color ActiveBg = Color.White;
        static readonly Color HoverBg = Color.FromArgb(236, 239, 242);
        static readonly Color TextColor = Color.FromArgb(32, 32, 32);
        static readonly Color DimText = Color.FromArgb(80, 84, 90);
        static readonly Color SepColor = Color.FromArgb(180, 186, 194);
        // 浮动模式下标签之外的区域用这个颜色绘制，配合 TransparencyKey 变成透明且可点穿
        static readonly Color KeyColor = Color.FromArgb(201, 203, 207);

        readonly Manager mgr;
        public readonly List<DocTab> Tabs = new List<DocTab>();
        // 组内的当前文档（最近一次处于前台的文档窗口）
        public IntPtr CurDoc;
        // 同组共享的窗口状态
        public bool HasState;
        public bool GroupMax;
        public Native.RECT GroupNormal;

        BarMode mode = BarMode.Hidden;
        bool isTopmost;
        Rectangle lastRect;
        bool reassertZ;
        bool overlayShown;
        int hoverTicks, leaveTicks;
        bool destroying;

        int hoverIndex = -1;
        bool hoverClose;
        int pressIndex = -1;
        Point pressPoint;
        Point pressScreen;
        bool pressOnBar;
        bool dragging;
        // 当前文档所在显示器的缩放比例；主副屏缩放不同时随文档所在的屏幕变化
        float s;
        int barHeight;

        Font tabFont;
        Font badgeFont;
        readonly System.Windows.Forms.Timer syncTimer;
        readonly System.Windows.Forms.Timer hoverTimer;
        readonly System.Windows.Forms.Timer pressTimer;
        readonly System.Windows.Forms.Timer roomTimer;
        readonly ToolTip tip;

        public TabBar(Manager mgr)
        {
            this.mgr = mgr;
            UpdateScale(IntPtr.Zero);

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = BarBg;
            TransparencyKey = KeyColor;
            Text = "DocTabs";
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Bounds = new Rectangle(0, -barHeight * 3, 200, barHeight);
            CreateHandle();

            tip = new ToolTip { ShowAlways = true, InitialDelay = 600, ReshowDelay = 200 };
            // 拖动/缩放结束后再把其他文档同步到同一位置，避免拖动过程中反复重排
            syncTimer = new System.Windows.Forms.Timer { Interval = 150 };
            syncTimer.Tick += delegate { syncTimer.Stop(); SyncOthersNormal(); };
            hoverTimer = new System.Windows.Forms.Timer { Interval = 100 };
            hoverTimer.Tick += delegate { OnHoverTick(); };
            // 按住标签栏期间轮询鼠标：标签栏从不激活，鼠标一离开它就收不到鼠标消息，拖快了也不能漏判
            pressTimer = new System.Windows.Forms.Timer { Interval = 10 };
            pressTimer.Tick += delegate { PressTick(); };
            // 窗口贴到屏幕上沿后，稍等位置稳定（系统分屏动画结束）再往下让出标签栏
            roomTimer = new System.Windows.Forms.Timer { Interval = 250 };
            roomTimer.Tick += delegate { roomTimer.Stop(); MakeRoom(); };
        }

        public void Destroy()
        {
            destroying = true;
            syncTimer.Stop();
            hoverTimer.Stop();
            pressTimer.Stop();
            roomTimer.Stop();
            tip.Dispose();
            if (tabFont != null) tabFont.Dispose();
            if (badgeFont != null) badgeFont.Dispose();
            Dispose();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= Native.WS_EX_TOOLWINDOW | Native.WS_EX_NOACTIVATE;
                return cp;
            }
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_MOUSEACTIVATE) { m.Result = (IntPtr)Native.MA_NOACTIVATE; return; }
            // 标签栏的位置和大小完全由我们自己按文档所在屏幕计算，不接受系统按 DPI 建议的尺寸
            if (m.Msg == Native.WM_DPICHANGED) { m.Result = IntPtr.Zero; return; }
            // 外部发来的关闭请求（如结束任务）：整个程序正常退出，把调整过的窗口恢复原状
            if (m.Msg == Native.WM_CLOSE && !destroying) { mgr.ExitApp(false); return; }
            base.WndProc(ref m);
        }

        public int IndexOf(IntPtr h)
        {
            if (h == IntPtr.Zero) return -1;
            for (int i = 0; i < Tabs.Count; i++) if (Tabs[i].Handle == h) return i;
            return -1;
        }

        public bool TabStripContains(Point screen)
        {
            if (!Bounds.Contains(screen)) return false;
            Point p = PointToClient(screen);
            if (Tabs.Count == 0) return false;
            return p.X <= Tabs[Tabs.Count - 1].Bounds.Right + 40 * s;
        }

        // 扫描之后：确定当前文档、刷新布局和位置
        public void AfterScan(List<DocTab> zOrder, IntPtr fg, bool changed)
        {
            if (IndexOf(CurDoc) < 0) CurDoc = IntPtr.Zero;
            bool fgHere = IndexOf(fg) >= 0;
            if (fgHere && fg != CurDoc) { OnDocForeground(fg); changed = true; }
            else if (!fgHere && mgr.DragGroup != this && !pressOnBar)
            {
                // 本组没有文档在前台（例如从资源管理器打开了新文档，但它没能抢到前台）：
                // 以 Z 序最上面、看得见的文档作为当前文档，标签栏贴在它上面，避免被它盖住
                foreach (var t in zOrder)
                {
                    if (IndexOf(t.Handle) < 0 || !Native.IsWindowVisible(t.Handle) || Native.IsIconic(t.Handle)) continue;
                    if (t.Handle != CurDoc) { CurDoc = t.Handle; changed = true; }
                    break;
                }
            }
            if (CurDoc == IntPtr.Zero && Tabs.Count > 0) { CurDoc = Tabs[0].Handle; changed = true; }
            if (!HasState && CurDoc != IntPtr.Zero) CaptureGroup(CurDoc);
            if (changed) LayoutTabs();
            reassertZ = true;
            UpdatePlacement();
            if (changed) Invalidate();
        }

        // ---------------- 同组窗口联动 ----------------

        public void OnDocForeground(IntPtr h)
        {
            if (h == CurDoc) return;
            CurDoc = h;
            // 已被最小化（例如全组最小化时前台恰好切到它）就不要再把它恢复回来
            if (Native.IsIconic(h)) { Invalidate(); return; }
            if (!HasState) CaptureGroup(h);
            else ApplyGroupState(h, true);
            Invalidate();
        }

        void CaptureGroup(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            GroupMax = mgr.IsMax(h);
            GroupNormal = Native.GetPlacement(h).rcNormalPosition;
            HasState = true;
        }

        // 让窗口 h 与分组状态一致：最大化 / 相同的还原位置
        public void ApplyGroupState(IntPtr h, bool activating)
        {
            if (!HasState)
            {
                if (Native.IsIconic(h)) Native.ShowWindow(h, Native.SW_RESTORE);
                return;
            }
            if (GroupMax)
            {
                if (mgr.IsNoShift(h)) { if (!mgr.IsPseudoMax(h)) mgr.PseudoMaximize(h); return; }
                if (!Native.IsZoomed(h)) Native.ShowWindow(h, Native.SW_MAXIMIZE);
                mgr.EnsureShift(h);
                return;
            }
            var wp = Native.GetPlacement(h);
            if (Native.IsZoomed(h) || Native.IsIconic(h) || !wp.rcNormalPosition.Equals(GroupNormal))
            {
                wp.flags = 0;
                wp.showCmd = activating ? Native.SW_SHOWNORMAL : Native.SW_SHOWNOACTIVATE;
                wp.rcNormalPosition = GroupNormal;
                Native.SetWindowPlacement(h, ref wp);
                mgr.Shifted.Remove(h);
                mgr.PseudoMax.Remove(h);
            }
        }

        public void OnDocLocation(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            bool zoomed = Native.IsZoomed(h);
            bool wasPseudo = mgr.PseudoMax.Contains(h);
            if (h != CurDoc)
            {
                // 组内其他文档（多半是我们自己同步过去的）只需保证最大化时让出标签栏
                if (zoomed) mgr.EnsureShift(h);
                return;
            }
            if (zoomed && wasPseudo && HasState)
            {
                // 伪最大化窗口显示的是“最大化”按钮，用户点它（或双击标题栏）就当作“还原”
                mgr.PseudoMax.Remove(h);
                GroupMax = false;
                ApplyGroupState(h, true);
                ScheduleSync();
                UpdatePlacement();
                return;
            }
            bool inDrag = mgr.DragGroup == this;
            if (zoomed)
            {
                mgr.EnsureShift(h);
                GroupMax = true;
            }
            else if (mgr.IsPseudoMax(h))
            {
                GroupMax = true;
            }
            else
            {
                mgr.Shifted.Remove(h);
                var np = Native.GetPlacement(h).rcNormalPosition;
                bool wasMax = GroupMax;
                bool changed = GroupMax || !np.Equals(GroupNormal);
                GroupMax = false;
                GroupNormal = np;
                if (changed && !inDrag)
                {
                    // 缩放过程中实时带动其他文档，结束后再由 syncTimer 精确校正
                    if (!wasMax) LiveSync(h);
                    ScheduleSync();
                }
                // 窗口顶到了屏幕上沿（例如系统分屏）：稍后往下让出标签栏
                if (!inDrag && NeedsRoom(h)) { roomTimer.Stop(); roomTimer.Start(); }
            }
            HasState = true;
            UpdatePlacement();
        }

        public void ScheduleSync()
        {
            syncTimer.Stop();
            syncTimer.Start();
        }

        void LiveSync(IntPtr h)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return;
            Native.RECT f = Native.FrameBounds(h, r);
            foreach (var t in Tabs)
            {
                IntPtr o = t.Handle;
                if (o == h || Native.IsIconic(o) || Native.IsZoomed(o) || mgr.PseudoMax.Contains(o)) continue;
                Native.RECT ro;
                if (!Native.GetWindowRect(o, out ro)) continue;
                Native.RECT fo = Native.FrameBounds(o, ro);
                int l = f.Left - (fo.Left - ro.Left), tp = f.Top - (fo.Top - ro.Top);
                int rr = f.Right + (ro.Right - fo.Right), b = f.Bottom + (ro.Bottom - fo.Bottom);
                if (l == ro.Left && tp == ro.Top && rr == ro.Right && b == ro.Bottom) continue;
                // 异步下发，Office 重排再慢也不会卡住标签栏自己
                Native.SetWindowPos(o, IntPtr.Zero, l, tp, rr - l, b - tp,
                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER | Native.SWP_ASYNCWINDOWPOS);
            }
        }

        void SyncOthersNormal()
        {
            if (!HasState || GroupMax || mgr.DragGroup == this) return;
            foreach (var t in Tabs)
            {
                IntPtr h = t.Handle;
                if (h == CurDoc || Native.IsIconic(h) || mgr.HiddenByUs.Contains(h)) continue;
                ApplyGroupState(h, false);
            }
        }

        public void MinimizeOthers(IntPtr h)
        {
            foreach (var t in new List<DocTab>(Tabs))
                if (t.Handle != h && !Native.IsIconic(t.Handle))
                    Native.ShowWindow(t.Handle, Native.SW_SHOWMINNOACTIVE);
        }

        // 窗口上方放不下标签栏（顶到了屏幕上沿）
        bool NeedsRoom(IntPtr h)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            int overlap = (int)Math.Round(2 * s);
            return f.Top < wa.Top + Manager.BarHeightFor(h) - overlap - (int)Math.Round(3 * s);
        }

        // 往下让出标签栏的高度（底边不动）；窗口太矮时保持自动隐藏
        void MakeRoom()
        {
            IntPtr h = CurDoc;
            if (h == IntPtr.Zero || mgr.DragGroup != null || pressOnBar || Native.IsIconic(h) || mgr.IsMax(h)) return;
            if (!NeedsRoom(h)) return;
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            int want = wa.Top + Manager.BarHeightFor(h) - (int)Math.Round(2 * s);
            int dy = want - f.Top;
            if (f.Bottom - want < 240 * s) return;
            Native.SetWindowPos(h, IntPtr.Zero, r.Left, r.Top + dy, r.Right - r.Left, r.Bottom - r.Top - dy,
                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
        }

        // 标签栏按当前文档所在显示器的缩放比例绘制
        void UpdateScale(IntPtr h)
        {
            float ns = Native.MonitorScale(h);
            if (tabFont != null && Math.Abs(ns - s) < 0.01f) return;
            s = ns;
            barHeight = (int)Math.Round(30 * s);
            // 按像素建字体：进程是逐显示器 DPI 感知的，按“磅”建的字体在不同缩放的屏幕上大小会不对
            if (tabFont != null) tabFont.Dispose();
            if (badgeFont != null) badgeFont.Dispose();
            tabFont = new Font("Microsoft YaHei UI", 12f * s, GraphicsUnit.Pixel);
            badgeFont = new Font("Segoe UI", 10f * s, FontStyle.Bold, GraphicsUnit.Pixel);
            LayoutTabs();
            Invalidate();
        }

        // ---------------- 标签栏位置 ----------------

        bool IsFullscreen(IntPtr h)
        {
            if (h == IntPtr.Zero || h == Handle) return false;
            string cls = Native.GetClassNameOf(h);
            if (cls == "Progman" || cls == "WorkerW" || cls == "Shell_TrayWnd") return false;
            int style = Native.GetWindowLong(h, Native.GWL_STYLE);
            if ((style & Native.WS_CAPTION) == Native.WS_CAPTION) return false;
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Rectangle m = Screen.FromHandle(h).Bounds;
            return r.Left <= m.Left && r.Top <= m.Top && r.Right >= m.Right && r.Bottom >= m.Bottom;
        }

        public void UpdatePlacement()
        {
            if (destroying) return;
            IntPtr fg = Native.GetForegroundWindow();
            BarMode m;
            Rectangle rect = Rectangle.Empty;
            if (CurDoc == IntPtr.Zero || Tabs.Count == 0 || !Native.IsWindow(CurDoc) || Native.IsIconic(CurDoc) ||
                !Native.IsWindowVisible(CurDoc) || (fg == CurDoc && IsFullscreen(fg)))
            {
                m = BarMode.Hidden;
            }
            else if (mgr.IsMax(CurDoc))
            {
                UpdateScale(CurDoc);
                Rectangle wa = Screen.FromHandle(CurDoc).WorkingArea;
                m = BarMode.Maximized;
                rect = new Rectangle(wa.Left, wa.Top, wa.Width, barHeight);
            }
            else
            {
                UpdateScale(CurDoc);
                Native.RECT wr;
                Native.GetWindowRect(CurDoc, out wr);
                Native.RECT f = Native.FrameBounds(CurDoc, wr);
                Rectangle wa = Screen.FromHandle(CurDoc).WorkingArea;
                int w = Math.Max(1, Math.Min(f.Right - f.Left, DesiredWidth()));
                // 向下压住标题栏顶部的边框线，让当前标签和标题栏连成一体
                int overlap = (int)Math.Round(2 * s);
                if (f.Top - barHeight + overlap >= wa.Top - (int)Math.Round(3 * s))
                {
                    m = BarMode.Follow;
                    rect = new Rectangle(f.Left, Math.Max(wa.Top, f.Top - barHeight + overlap), w, barHeight);
                }
                else
                {
                    m = BarMode.AutoHide;
                    rect = new Rectangle(f.Left, f.Top, w, barHeight);
                }
            }

            if (m != BarMode.AutoHide)
            {
                overlayShown = false;
                hoverTimer.Stop();
            }
            else if (!hoverTimer.Enabled) { hoverTicks = 0; leaveTicks = 0; hoverTimer.Start(); }

            bool modeChanged = m != mode;
            mode = m;
            if (m == BarMode.Hidden || (m == BarMode.AutoHide && !overlayShown))
            {
                if (Visible) Hide();
                lastRect = Rectangle.Empty;
                return;
            }

            // 本组文档在前台时置顶，Office 被点击激活也盖不住标签栏；切到其他程序后取消置顶，只紧贴在文档上面一层
            bool docInFront = IndexOf(fg) >= 0 || fg == Handle;
            bool wantTopmost = m == BarMode.AutoHide || docInFront;
            bool moved = rect != lastRect || !Visible;
            if (!Visible) Show();
            const uint moveFlags = Native.SWP_NOACTIVATE | Native.SWP_NOZORDER;
            if (wantTopmost)
            {
                // 已经置顶时只移动位置，不重复调整 Z 序（拖动窗口时每帧都会走到这里）
                if (!isTopmost || reassertZ)
                    Native.SetWindowPos(Handle, Native.HWND_TOPMOST, rect.Left, rect.Top, rect.Width, rect.Height, Native.SWP_NOACTIVATE);
                else if (moved)
                    Native.SetWindowPos(Handle, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height, moveFlags);
            }
            else
            {
                if (isTopmost)
                    Native.SetWindowPos(Handle, Native.HWND_NOTOPMOST, rect.Left, rect.Top, rect.Width, rect.Height, Native.SWP_NOACTIVATE);
                else if (moved)
                    Native.SetWindowPos(Handle, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height, moveFlags);
                // 确认标签栏紧贴在当前文档上面一层：新打开或被提到前面的文档不能把它盖住
                IntPtr above = Native.GetWindow(CurDoc, Native.GW_HWNDPREV);
                if (above != Handle)
                {
                    bool aboveTopmost = above != IntPtr.Zero &&
                        (Native.GetWindowLong(above, Native.GWL_EXSTYLE) & Native.WS_EX_TOPMOST) != 0;
                    Native.SetWindowPos(Handle, above == IntPtr.Zero || aboveTopmost ? Native.HWND_TOP : above, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE);
                }
            }
            isTopmost = wantTopmost;
            reassertZ = false;
            lastRect = rect;
            if (modeChanged) { LayoutTabs(); Invalidate(); }
        }

        int DesiredWidth()
        {
            int n = Math.Max(1, Tabs.Count);
            return (int)Math.Ceiling(12 * s + n * 210 * s + (n - 1) * 2 * s);
        }

        bool Floating
        {
            get { return mode == BarMode.Follow || mode == BarMode.AutoHide; }
        }

        // 小窗口且上方没空间时：鼠标停在窗口上沿约 0.2 秒滑出标签栏，离开 0.5 秒后收起
        void OnHoverTick()
        {
            if (mode != BarMode.AutoHide || CurDoc == IntPtr.Zero) { hoverTimer.Stop(); return; }
            Point p = Cursor.Position;
            if (!overlayShown)
            {
                Native.RECT wr;
                Native.GetWindowRect(CurDoc, out wr);
                Native.RECT f = Native.FrameBounds(CurDoc, wr);
                int zone = (int)(4 * s);
                bool inZone = p.X >= f.Left && p.X < f.Right && p.Y >= f.Top - zone && p.Y <= f.Top + zone &&
                              Control.MouseButtons == MouseButtons.None;
                hoverTicks = inZone ? hoverTicks + 1 : 0;
                if (hoverTicks >= 2) { overlayShown = true; leaveTicks = 0; UpdatePlacement(); }
            }
            else
            {
                Rectangle r = Bounds;
                r.Inflate((int)(6 * s), (int)(6 * s));
                leaveTicks = (r.Contains(p) || pressOnBar) ? 0 : leaveTicks + 1;
                if (leaveTicks >= 5) { overlayShown = false; hoverTicks = 0; UpdatePlacement(); }
            }
        }

        // ---------------- 绘制 ----------------

        public void LayoutTabs()
        {
            float pad = Floating ? 6 * s : 4 * s, gap = Floating ? 2 * s : 1 * s, top = 4 * s;
            float h = barHeight - top;
            int n = Tabs.Count;
            if (n == 0) return;
            float avail = ClientSize.Width - pad * 2;
            float w = Math.Max(1, Math.Min(210 * s, (avail - gap * (n - 1)) / n));
            float x = pad;
            foreach (var t in Tabs)
            {
                t.Bounds = new RectangleF(x, top, w, h);
                float cs = 16 * s;
                t.CloseBounds = new RectangleF(x + w - cs - 6 * s, top + (h - cs) / 2, cs, cs);
                x += w + gap;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutTabs();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            bool floating = Floating;
            g.Clear(floating ? KeyColor : BarBg);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            for (int i = 0; i < Tabs.Count; i++)
            {
                var t = Tabs[i];
                bool active = t.Handle == CurDoc;
                bool hover = i == hoverIndex;
                var b = t.Bounds;

                if (active || hover || floating)
                {
                    using (var path = TopRounded(b, 6 * s))
                    using (var br = new SolidBrush(active ? ActiveBg : hover ? HoverBg : BarBg))
                        g.FillPath(br, path);
                }
                else if (i + 1 < Tabs.Count && Tabs[i + 1].Handle != CurDoc && i + 1 != hoverIndex)
                {
                    using (var pen = new Pen(SepColor, Math.Max(1f, s)))
                        g.DrawLine(pen, b.Right, b.Top + b.Height * 0.28f, b.Right, b.Bottom - b.Height * 0.28f);
                }

                // 应用标识
                float bw = 16 * s, bh = 16 * s;
                var badge = new RectangleF(b.Left + 8 * s, b.Top + (b.Height - bh) / 2, bw, bh);
                using (var path = Rounded(badge, 3 * s))
                using (var br = new SolidBrush(KindColor(t.Kind)))
                    g.FillPath(br, path);
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(t.Kind == DocKind.Word ? "W" : "X", badgeFont, Brushes.White, badge, sf);

                // 标题
                bool showClose = active || hover;
                float textLeft = badge.Right + 6 * s;
                float textRight = showClose ? t.CloseBounds.Left - 2 * s : b.Right - 8 * s;
                var textRect = Rectangle.Round(new RectangleF(textLeft, b.Top, Math.Max(0, textRight - textLeft), b.Height));
                TextRenderer.DrawText(g, t.Title, tabFont, textRect, active ? TextColor : DimText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

                // 关闭按钮
                if (showClose)
                {
                    var c = t.CloseBounds;
                    if (hover && hoverClose)
                        using (var br = new SolidBrush(Color.FromArgb(214, 218, 224))) g.FillEllipse(br, c);
                    float d = 4 * s;
                    float cx = c.Left + c.Width / 2, cy = c.Top + c.Height / 2;
                    using (var pen = new Pen(DimText, 1.2f * s))
                    {
                        g.DrawLine(pen, cx - d, cy - d, cx + d, cy + d);
                        g.DrawLine(pen, cx + d, cy - d, cx - d, cy + d);
                    }
                }
            }

            if (floating) return;
            using (var pen = new Pen(BarLine, Math.Max(1f, s)))
            {
                float y = ClientSize.Height - pen.Width / 2;
                int ai = IndexOf(CurDoc);
                if (ai < 0) g.DrawLine(pen, 0, y, ClientSize.Width, y);
                else
                {
                    var ab = Tabs[ai].Bounds;
                    g.DrawLine(pen, 0, y, ab.Left, y);
                    g.DrawLine(pen, ab.Right, y, ClientSize.Width, y);
                }
            }
        }

        static Color KindColor(DocKind k)
        {
            return k == DocKind.Word ? Color.FromArgb(24, 90, 189) : Color.FromArgb(16, 124, 65);
        }

        static GraphicsPath TopRounded(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = rad * 2;
            p.AddLine(r.Left, r.Bottom, r.Left, r.Top + rad);
            p.AddArc(r.Left, r.Top, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            p.AddLine(r.Right, r.Top + rad, r.Right, r.Bottom);
            p.CloseFigure();
            return p;
        }

        static GraphicsPath Rounded(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = rad * 2;
            p.AddArc(r.Left, r.Top, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ---------------- 鼠标交互 ----------------

        int HitTest(Point p)
        {
            for (int i = 0; i < Tabs.Count; i++) if (Tabs[i].Bounds.Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (pressOnBar && e.Button == MouseButtons.Left && CheckStartDrag(Cursor.Position)) return;
            if (pressOnBar && pressIndex >= 0 && e.Button == MouseButtons.Left)
            {
                // 水平拖动标签：调整顺序
                if (!dragging && Math.Abs(e.X - pressPoint.X) > 6 * s) dragging = true;
                if (dragging)
                {
                    int target = HitTest(new Point(e.X, (int)Tabs[pressIndex].Bounds.Top + 2));
                    if (target >= 0 && target != pressIndex)
                    {
                        var t = Tabs[pressIndex];
                        Tabs.RemoveAt(pressIndex);
                        Tabs.Insert(target, t);
                        pressIndex = target;
                        hoverIndex = target;
                        LayoutTabs();
                        Invalidate();
                    }
                    return;
                }
            }
            int idx = HitTest(e.Location);
            bool overClose = idx >= 0 && Tabs[idx].CloseBounds.Contains(e.Location);
            if (idx != hoverIndex || overClose != hoverClose)
            {
                if (idx != hoverIndex) tip.SetToolTip(this, idx >= 0 ? Tabs[idx].FullTitle : null);
                hoverIndex = idx;
                hoverClose = overClose;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (!pressOnBar) { hoverIndex = -1; hoverClose = false; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int idx = HitTest(e.Location);
            if (e.Button == MouseButtons.Left)
            {
                pressOnBar = true;
                pressScreen = Cursor.Position;
                pressIndex = idx;
                pressPoint = e.Location;
                dragging = false;
                pressTimer.Start();
            }
            else if (e.Button == MouseButtons.Middle && idx >= 0) CloseTab(Tabs[idx]);
            else if (e.Button == MouseButtons.Right) ShowTabMenu(idx, e.Location);
        }

        // 双击标签栏空白处：最大化 / 还原，与双击标题栏一致
        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left || HitTest(e.Location) >= 0) return;
            if (IndexOf(CurDoc) < 0) return;
            ActivateDoc(CurDoc);
            if (Native.IsZoomed(CurDoc)) Native.ShowWindow(CurDoc, Native.SW_RESTORE);
            else if (mgr.IsPseudoMax(CurDoc)) { GroupMax = false; mgr.PseudoMax.Remove(CurDoc); ApplyGroupState(CurDoc, true); }
            else Native.ShowWindow(CurDoc, Native.SW_MAXIMIZE);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            bool wasPress = pressOnBar;
            int pressed = pressIndex;
            bool wasDrag = dragging;
            ResetPress();
            if (!wasPress || wasDrag || pressed < 0 || pressed >= Tabs.Count) return;
            var t = Tabs[pressed];
            if (!t.Bounds.Contains(e.Location)) return;
            if (t.CloseBounds.Contains(e.Location) && (t.Handle == CurDoc || pressed == hoverIndex)) CloseTab(t);
            else ActivateDoc(t.Handle);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (Tabs.Count == 0) return;
            int cur = IndexOf(CurDoc);
            int next = e.Delta > 0 ? cur - 1 : cur + 1;
            if (cur < 0) next = 0;
            next = (next + Tabs.Count) % Tabs.Count;
            ActivateDoc(Tabs[next].Handle);
        }

        void ShowTabMenu(int idx, Point p)
        {
            var items = new List<MenuItem>();
            if (idx >= 0)
            {
                var t = Tabs[idx];
                items.Add(new MenuItem("关闭", delegate { CloseTab(t); }));
                items.Add(new MenuItem("关闭其他标签", delegate
                {
                    var rest = new List<IntPtr>();
                    foreach (var o in Tabs) if (o != t) rest.Add(o.Handle);
                    mgr.StartCascade(rest);
                }));
                items.Add(new MenuItem("关闭全部", delegate
                {
                    var all = new List<IntPtr>();
                    foreach (var o in Tabs) if (o != t) all.Add(o.Handle);
                    all.Add(t.Handle);
                    mgr.StartCascade(all);
                }));
                if (Tabs.Count > 1)
                    items.Add(new MenuItem("拆分为单独窗口", delegate { int i = IndexOf(t.Handle); if (i >= 0) DetachToSide(i); }));
                if (mgr.Groups.Count > 1)
                    items.Add(new MenuItem("合并所有窗口", delegate { MergeAllInto(); }));
                items.Add(new MenuItem("-"));
            }
            items.Add(new MenuItem("退出 DocTabs", delegate { mgr.ExitApp(true); }));
            var menu = new ContextMenu(items.ToArray());
            Native.SetForegroundWindow(Handle);
            menu.Show(this, p);
        }

        // ---------------- 拖动：拆分标签 / 移动窗口 ----------------

        // 按住标签栏期间每 10ms 一次：判断是否开始拖动、发现在别处松手
        void PressTick()
        {
            if (!pressOnBar) { pressTimer.Stop(); return; }
            // 读物理按键状态：在别处松手时消息不会发给标签栏
            int vk = SystemInformation.MouseButtonsSwapped ? Native.VK_RBUTTON : Native.VK_LBUTTON;
            bool down = (Native.GetAsyncKeyState(vk) & 0x8000) != 0;
            Point c = Cursor.Position;
            if (!down)
            {
                // 在标签栏上松手的交给 OnMouseUp 处理（点击切换）；在别处松手的直接收尾
                if (!Bounds.Contains(c)) ResetPress();
                return;
            }
            CheckStartDrag(c);
        }

        // 多个标签时把标签竖直拖出标签栏：拆成单独的窗口；水平拖动标签：调整顺序（在 OnMouseMove 里）。
        // 拖空白处、或只有一个标签时拖标签：移动窗口
        bool CheckStartDrag(Point c)
        {
            if (!pressOnBar) return false;
            int dx = Math.Abs(c.X - pressScreen.X), dy = Math.Abs(c.Y - pressScreen.Y);
            if (pressIndex >= 0 && pressIndex < Tabs.Count && Tabs.Count > 1)
            {
                bool detach = dragging ? dy > barHeight * 0.8f : dy > barHeight * 0.6f && dy > dx;
                if (!detach) return false;
                Detach(pressIndex, c);
                return true;
            }
            if (Math.Max(dx, dy) <= 4 * s) return false;
            BeginBarMove(c);
            return true;
        }

        void ResetPress()
        {
            pressOnBar = false;
            pressIndex = -1;
            dragging = false;
            pressTimer.Stop();
        }

        // 移动整个窗口：最大化时先在鼠标下还原，然后交给系统的拖动流程
        void BeginBarMove(Point c)
        {
            IntPtr h = pressIndex >= 0 && pressIndex < Tabs.Count ? Tabs[pressIndex].Handle : CurDoc;
            ResetPress();
            if (IndexOf(h) < 0 || Native.IsIconic(h)) return;
            hoverIndex = -1;
            hoverClose = false;
            tip.Hide(this);
            // 按着鼠标时不切前台（见 StartNativeMove），只更新分组状态，松手后再切
            if (h != CurDoc) { ApplyGroupState(h, true); CurDoc = h; }
            mgr.Active = this;
            if (mgr.IsMax(h)) RestoreUnderCursor(h, c, pressScreen.X - Left, pressScreen.Y - Top);
            Invalidate();
            mgr.StartNativeMove(h);
        }

        // 从最大化状态拖出：还原成普通大小，并让鼠标仍落在标签栏里原来抓住的位置
        void RestoreUnderCursor(IntPtr h, Point c, int grabX, int grabY)
        {
            var wp = Native.GetPlacement(h);
            Native.RECT n = HasState ? GroupNormal : wp.rcNormalPosition;
            int w = Math.Max(200, n.Right - n.Left), ht = Math.Max(150, n.Bottom - n.Top);
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            int side = f.Left - r.Left; // 左侧隐形边框，最大化和普通状态下一样宽
            int frameW = w - side * 2;
            Rectangle wa = Screen.FromPoint(c).WorkingArea;
            if (grabX > frameW - 60 * s) grabX = (int)((long)frameW * (c.X - wa.Left) / Math.Max(1, wa.Width));
            PlaceNormal(h, this, c.X - grabX - side, c.Y - grabY, w, ht, c);
        }

        // 把窗口 h 以普通状态放到屏幕坐标 (left, barTop + 标签栏高度) 处，并记为分组 g 的还原位置
        void PlaceNormal(IntPtr h, TabBar g, int left, int barTop, int w, int ht, Point c)
        {
            Screen scr = Screen.FromPoint(c);
            Rectangle wa = scr.WorkingArea, mon = scr.Bounds;
            int overlap = (int)Math.Round(2 * s);
            int top = Math.Max(wa.Top + barHeight - overlap, barTop + barHeight - overlap);
            // rcNormalPosition 用的是工作区坐标（相对于工作区左上角）
            int ox = wa.Left - mon.Left, oy = wa.Top - mon.Top;
            var n = new Native.RECT { Left = left - ox, Top = top - oy, Right = left - ox + w, Bottom = top - oy + ht };
            mgr.Shifted.Remove(h);
            mgr.PseudoMax.Remove(h);
            g.GroupMax = false;
            g.GroupNormal = n;
            g.HasState = true;
            var wp = Native.GetPlacement(h);
            wp.flags = 0;
            wp.showCmd = Native.SW_SHOWNORMAL;
            wp.rcNormalPosition = n;
            Native.SetWindowPlacement(h, ref wp);
        }

        // 把第 index 个标签拆成单独一组，窗口出现在鼠标下面并继续跟着鼠标走
        void Detach(int index, Point c)
        {
            var tab = Tabs[index];
            IntPtr h = tab.Handle;
            float grabX = Math.Max(10 * s, Math.Min(pressPoint.X - tab.Bounds.Left, 180 * s));
            int grabY = Math.Max(0, Math.Min(pressPoint.Y, barHeight - 1));
            ResetPress();
            hoverIndex = -1;
            hoverClose = false;
            tip.Hide(this);

            int w, ht;
            WindowSizeFor(h, out w, out ht);
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            int side = f.Left - r.Left;

            var ng = SplitOff(index);
            // 新的标签栏只有一个标签，它的左边距是 6 像素（按缩放）
            int frameLeft = c.X - (int)(6 * s + grabX);
            PlaceNormal(h, ng, frameLeft - side, c.Y - grabY, w, ht, c);
            // 不在这里切前台（见 StartNativeMove），松手后再切
            mgr.StartNativeMove(h);
            ng.UpdatePlacement();
        }

        // 右键菜单“拆分为单独窗口”：在原窗口旁边错开一点放
        void DetachToSide(int index)
        {
            var tab = Tabs[index];
            IntPtr h = tab.Handle;
            int w, ht;
            WindowSizeFor(h, out w, out ht);
            Rectangle wa = Screen.FromHandle(CurDoc != IntPtr.Zero ? CurDoc : h).WorkingArea;
            int off = (int)(40 * s);
            int left = GroupMax ? wa.Left + wa.Width / 4 : 0, barTop = 0;
            if (!GroupMax && HasState)
            {
                Native.RECT cr;
                Native.GetWindowRect(CurDoc, out cr);
                left = cr.Left + off;
                barTop = cr.Top - barHeight + off;
            }
            else barTop = wa.Top + wa.Height / 6;
            var ng = SplitOff(index);
            PlaceNormal(h, ng, left, barTop, w, ht, new Point(left + w / 2, barTop + barHeight));
            Manager.BringToFront(h);
            ng.UpdatePlacement();
        }

        // 从本组移出第 index 个标签，放进新建的一组
        TabBar SplitOff(int index)
        {
            var tab = Tabs[index];
            Tabs.RemoveAt(index);
            if (CurDoc == tab.Handle) CurDoc = Tabs[Math.Min(index, Tabs.Count - 1)].Handle;
            LayoutTabs();
            Invalidate();
            var ng = mgr.NewGroup();
            ng.Tabs.Add(tab);
            ng.CurDoc = tab.Handle;
            mgr.Active = ng;
            ng.LayoutTabs();
            UpdatePlacement();
            return ng;
        }

        // 拆出去的窗口用原分组还原时的大小
        void WindowSizeFor(IntPtr h, out int w, out int ht)
        {
            Native.RECT n;
            if (HasState) n = GroupNormal;
            else n = Native.GetPlacement(h).rcNormalPosition;
            w = Math.Max((int)(400 * s), n.Right - n.Left);
            ht = Math.Max((int)(300 * s), n.Bottom - n.Top);
        }

        void MergeAllInto()
        {
            foreach (var g in mgr.Groups.ToArray()) if (g != this) mgr.Merge(g, this);
        }

        // ---------------- 窗口操作 ----------------

        public void ActivateDoc(IntPtr h)
        {
            // 先把目标文档调成与当前文档相同的状态/位置，再切过去，看起来就像同一个窗口换了内容
            ApplyGroupState(h, true);
            Manager.BringToFront(h);
            CurDoc = h;
            mgr.Active = this;
            mgr.UpdateAll();
            Invalidate();
        }

        void CloseTab(DocTab t)
        {
            // Office 会自行弹出“是否保存”提示
            Native.PostMessage(t.Handle, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            mgr.ScheduleScan();
        }
    }

    static class Native
    {
        public const int WS_EX_TOOLWINDOW = 0x80, WS_EX_NOACTIVATE = 0x08000000;
        public const int WS_CAPTION = 0x00C00000, GWL_STYLE = -16, GWL_EXSTYLE = -20, WS_EX_TOPMOST = 0x8;
        public const int WM_MOUSEACTIVATE = 0x21, MA_NOACTIVATE = 3, WM_CLOSE = 0x10;
        public const int WM_SYSCOMMAND = 0x112, SC_MOVE = 0xF010, HTCAPTION = 2;
        public const uint GW_OWNER = 4, GW_HWNDPREV = 3, GA_ROOT = 2;
        public const int SW_SHOWNORMAL = 1, SW_MAXIMIZE = 3, SW_SHOWNOACTIVATE = 4, SW_SHOWMINNOACTIVE = 7, SW_RESTORE = 9;
        public const uint SWP_NOSIZE = 1, SWP_NOMOVE = 2, SWP_NOZORDER = 4, SWP_NOACTIVATE = 0x10, SWP_NOOWNERZORDER = 0x200, SWP_ASYNCWINDOWPOS = 0x4000;
        public static readonly IntPtr HWND_TOP = IntPtr.Zero, HWND_TOPMOST = new IntPtr(-1), HWND_NOTOPMOST = new IntPtr(-2);
        public const uint EVENT_SYSTEM_FOREGROUND = 3, EVENT_SYSTEM_MOVESIZESTART = 0x0A, EVENT_SYSTEM_MOVESIZEEND = 0x0B;
        public const uint EVENT_SYSTEM_MINIMIZESTART = 0x16, EVENT_SYSTEM_MINIMIZEEND = 0x17;
        public const int SW_HIDE = 0;
        public const uint EVENT_OBJECT_DESTROY = 0x8001, EVENT_OBJECT_HIDE = 0x8003;
        public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B, EVENT_OBJECT_NAMECHANGE = 0x800C;
        public const uint EVENT_OBJECT_CLOAKED = 0x8017, EVENT_OBJECT_UNCLOAKED = 0x8018;
        public const uint WINEVENT_OUTOFCONTEXT = 0, WINEVENT_SKIPOWNPROCESS = 2;
        public const int OBJID_WINDOW = 0;
        public const uint INPUT_MOUSE = 0;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
            public bool Equals(RECT o) { return Left == o.Left && Top == o.Top && Right == o.Right && Bottom == o.Bottom; }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X, Y; }

        public const int GUI_INMOVESIZE = 2, WM_KEYDOWN = 0x100, VK_RETURN = 0x0D;

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize, flags;
            public IntPtr hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO info);

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx, dy;
            public uint mouseData, dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        public delegate void WinEventProc(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);
        public delegate IntPtr LowLevelProc(int code, IntPtr wParam, IntPtr lParam);

        public const int WH_KEYBOARD_LL = 13, WH_MOUSE_LL = 14, WM_LBUTTONUP = 0x202, WM_SYSKEYDOWN = 0x104;
        public const int VK_LBUTTON = 1, VK_RBUTTON = 2;
        [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vk);
        [DllImport("user32.dll")] public static extern bool ReleaseCapture();

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData, flags, time; public IntPtr extra; }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr extra; }

        [DllImport("user32.dll")] public static extern IntPtr SetWindowsHookEx(int id, LowLevelProc proc, IntPtr mod, uint thread);
        [DllImport("user32.dll")] public static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetModuleHandle(string name);

        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
        [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr h, uint flags);
        [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(POINT pt, uint flags);
        [DllImport("shcore.dll")] static extern int GetDpiForMonitor(IntPtr mon, int type, out uint dpiX, out uint dpiY);
        public const int WM_DPICHANGED = 0x02E0;

        // 窗口所在显示器的缩放比例（96 DPI = 1.0）；h 为空时取主显示器
        public static float MonitorScale(IntPtr h)
        {
            try
            {
                IntPtr mon = h != IntPtr.Zero ? MonitorFromWindow(h, 2 /* NEAREST */) : MonitorFromPoint(new POINT(), 1 /* PRIMARY */);
                uint x, y;
                if (mon != IntPtr.Zero && GetDpiForMonitor(mon, 0 /* EFFECTIVE */, out x, out y) == 0 && x > 0) return x / 96f;
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) return g.DpiX / 96f;
        }
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
        [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr h, bool altTab);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] static extern bool GetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
        [DllImport("user32.dll")] public static extern bool SetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
        [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int index);
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr mod, WinEventProc proc, uint pid, uint tid, uint flags);
        [DllImport("user32.dll")] public static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("user32.dll")] public static extern uint SendInput(uint count, INPUT[] inputs, int size);
        [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr h);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out int value, int size);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out RECT value, int size);
        [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr h, int flags, StringBuilder sb, ref int size);

        public static string GetClassNameOf(IntPtr h)
        {
            var sb = new StringBuilder(64);
            GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string GetWindowTextOf(IntPtr h)
        {
            int len = GetWindowTextLength(h);
            if (len <= 0) return "";
            var sb = new StringBuilder(len + 1);
            GetWindowText(h, sb, sb.Capacity);
            return sb.ToString();
        }

        public static bool IsCloaked(IntPtr h)
        {
            int v;
            return DwmGetWindowAttribute(h, 14 /* DWMWA_CLOAKED */, out v, 4) == 0 && v != 0;
        }

        // 窗口可见边框（不含 Win10/11 的隐形缩放边框）
        public static RECT FrameBounds(IntPtr h, RECT fallback)
        {
            RECT r;
            if (DwmGetWindowAttribute(h, 9 /* DWMWA_EXTENDED_FRAME_BOUNDS */, out r, Marshal.SizeOf(typeof(RECT))) == 0 && r.Right > r.Left)
                return r;
            return fallback;
        }

        public static WINDOWPLACEMENT GetPlacement(IntPtr h)
        {
            var wp = new WINDOWPLACEMENT();
            wp.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
            GetWindowPlacement(h, ref wp);
            return wp;
        }

        public static string GetProcessPath(uint pid)
        {
            IntPtr hp = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, pid);
            if (hp == IntPtr.Zero) return null;
            try
            {
                var sb = new StringBuilder(1024);
                int size = sb.Capacity;
                return QueryFullProcessImageName(hp, 0, sb, ref size) ? sb.ToString() : null;
            }
            finally { CloseHandle(hp); }
        }
    }
}
