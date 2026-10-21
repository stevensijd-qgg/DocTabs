// DocTabs — 轻量文档标签栏：把 Office 的 Word / Excel 窗口当作“一个窗口的多个标签”来切换（不含 WPS）。
// 最大化：标签栏固定在屏幕顶部，最大化的文档窗口整体下移让出标签栏高度（其他程序不受影响）。
// 小窗口：标签栏贴在窗口标题栏正上方随窗口移动；上方没有空间时自动隐藏，鼠标移到窗口上沿时滑出。
// 同组联动：一个文档最大化/还原/移动/最小化，其他文档同步为相同状态和位置；点窗口 × 或 Alt+F4 整组依次关闭。
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
[assembly: System.Reflection.AssemblyVersion("1.3.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.3.0.0")]

namespace DocTabs
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            bool created;
            using (var mutex = new Mutex(true, "DocTabs.SingleInstance.8F2C41", out created))
            {
                if (!created) return;
                Native.SetProcessDPIAware();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var bar = new TabBar();
                Application.Run();
                GC.KeepAlive(bar);
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

        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunName = "DocTabs";

        readonly List<DocTab> tabs = new List<DocTab>();
        readonly List<IntPtr> hooks = new List<IntPtr>();
        readonly Dictionary<uint, string> procNames = new Dictionary<uint, string>();
        readonly HashSet<IntPtr> shifted = new HashSet<IntPtr>();
        readonly HashSet<IntPtr> pseudoMax = new HashSet<IntPtr>();
        readonly HashSet<IntPtr> noShift = new HashSet<IntPtr>();
        // 拖动/缩放当前文档时被暂时隐藏的其他文档
        readonly HashSet<IntPtr> hiddenByUs = new HashSet<IntPtr>();
        bool windowDrag;
        DateTime windowDragStart;

        // 用户点了某个文档窗口的 ×（或按了 Alt+F4）：该窗口真正关掉后，整组一起关闭
        IntPtr closeIntent;
        DateTime closeIntentTime;
        // 正在逐个关闭的文档队列
        readonly List<IntPtr> cascade = new List<IntPtr>();
        IntPtr cascadeCur;
        DateTime cascadeStepTime;
        readonly Native.LowLevelProc mouseProc;
        readonly Native.LowLevelProc keyProc;
        IntPtr mouseHook, keyHook;
        readonly Native.WinEventProc winEventProc;
        readonly Native.EnumWindowsProc enumProc;
        List<DocTab> scanResult;

        // 当前文档（最近一次处于前台的文档窗口）
        IntPtr curDoc;
        BarMode mode = BarMode.Hidden;
        bool isTopmost;
        bool overlayShown;
        int hoverTicks, leaveTicks;

        // 同组共享的窗口状态
        bool hasGroup;
        bool groupMax;
        Native.RECT groupNormal;

        int hoverIndex = -1;
        bool hoverClose;
        int pressIndex = -1;
        Point pressPoint;
        bool dragging;
        readonly float s;
        readonly int barHeight;

        readonly Font tabFont;
        readonly Font badgeFont;
        readonly System.Windows.Forms.Timer scanTimer;
        readonly System.Windows.Forms.Timer fallbackTimer;
        readonly System.Windows.Forms.Timer syncTimer;
        readonly System.Windows.Forms.Timer hoverTimer;
        readonly System.Windows.Forms.Timer revealTimer;
        readonly NotifyIcon tray;
        readonly ToolTip tip;
        MenuItem autostartItem;

        public TabBar()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) s = g.DpiX / 96f;
            barHeight = (int)Math.Round(30 * s);

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = BarBg;
            TransparencyKey = KeyColor;
            Text = "DocTabs";
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Bounds = new Rectangle(0, -barHeight * 3, 200, barHeight);

            tabFont = new Font("Microsoft YaHei UI", 9f);
            badgeFont = new Font("Segoe UI", 7.5f, FontStyle.Bold);

            CreateHandle();

            tip = new ToolTip { ShowAlways = true, InitialDelay = 600, ReshowDelay = 200 };

            scanTimer = new System.Windows.Forms.Timer { Interval = 120 };
            scanTimer.Tick += delegate { scanTimer.Stop(); Scan(); };
            fallbackTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            fallbackTimer.Tick += delegate
            {
                // 兜底：漏掉了“拖动结束”通知时，别让文档一直处于隐藏状态
                if (windowDrag && (DateTime.Now - windowDragStart).TotalSeconds > 30) EndWindowDrag();
                if (!windowDrag && hiddenByUs.Count > 0 && !revealTimer.Enabled) RevealHidden();
                Scan();
            };
            fallbackTimer.Start();
            // 拖动/缩放结束后再把其他文档同步到同一位置，避免拖动过程中反复重排
            syncTimer = new System.Windows.Forms.Timer { Interval = 150 };
            syncTimer.Tick += delegate { syncTimer.Stop(); SyncOthersNormal(); };
            hoverTimer = new System.Windows.Forms.Timer { Interval = 100 };
            hoverTimer.Tick += delegate { OnHoverTick(); };
            // 松手后稍等当前文档的最终位置稳定，再把其他文档放到同一位置并显示出来
            revealTimer = new System.Windows.Forms.Timer { Interval = 80 };
            revealTimer.Tick += delegate { revealTimer.Stop(); RevealHidden(); };

            enumProc = EnumCallback;
            winEventProc = OnWinEvent;
            const uint flags = Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS;
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_FOREGROUND, Native.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MOVESIZESTART, Native.EVENT_SYSTEM_MOVESIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_SYSTEM_MINIMIZESTART, Native.EVENT_SYSTEM_MINIMIZEEND, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_DESTROY, Native.EVENT_OBJECT_HIDE, IntPtr.Zero, winEventProc, 0, 0, flags));
            hooks.Add(Native.SetWinEventHook(Native.EVENT_OBJECT_LOCATIONCHANGE, Native.EVENT_OBJECT_NAMECHANGE, IntPtr.Zero, winEventProc, 0, 0, flags));

            tray = new NotifyIcon { Icon = MakeTrayIcon(), Text = "DocTabs 文档标签栏", Visible = true };
            autostartItem = new MenuItem("开机自启动", delegate { ToggleAutostart(); }) { Checked = IsAutostart() };
            tray.ContextMenu = new ContextMenu(new[]
            {
                new MenuItem("DocTabs 文档标签栏") { Enabled = false },
                new MenuItem("-"),
                autostartItem,
                new MenuItem("退出", delegate { ExitApp(); })
            });

            SystemEvents.DisplaySettingsChanged += OnDisplayChanged;

            // 只记录“在当前文档的关闭按钮上松开左键”和“Alt+F4”这两个动作，回调里不做别的事
            mouseProc = OnMouseLL;
            keyProc = OnKeyLL;
            IntPtr mod = Native.GetModuleHandle(null);
            mouseHook = Native.SetWindowsHookEx(Native.WH_MOUSE_LL, mouseProc, mod, 0);
            keyHook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, keyProc, mod, 0);
            Scan();
        }

        IntPtr OnMouseLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_LBUTTONUP && curDoc != IntPtr.Zero)
            {
                var info = (Native.MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.MSLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                bool onBar = Visible && Bounds.Contains(info.pt.X, info.pt.Y);
                if (fg == curDoc && !onBar && IsOverCloseButton(fg, info.pt.X, info.pt.Y)) SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        IntPtr OnKeyLL(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam.ToInt32() == Native.WM_SYSKEYDOWN)
            {
                var info = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));
                IntPtr fg = Native.GetForegroundWindow();
                if (info.vkCode == 0x73 /* F4 */ && (info.flags & 0x20 /* ALT */) != 0 && IndexOf(fg) >= 0) SetCloseIntent(fg);
            }
            return Native.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        // Office 标题栏右上角的关闭按钮区域
        bool IsOverCloseButton(IntPtr h, int x, int y)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return false;
            Native.RECT f = Native.FrameBounds(h, r);
            return x >= f.Right - 52 * s && x <= f.Right && y >= f.Top && y <= f.Top + 56 * s;
        }

        void SetCloseIntent(IntPtr h)
        {
            closeIntent = h;
            closeIntentTime = DateTime.Now;
        }

        void StartCascade(IEnumerable<IntPtr> handles)
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
                if (!Native.IsWindow(h) || IndexOf(h) < 0) continue;
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
            base.WndProc(ref m);
        }

        // ---------------- 窗口扫描 ----------------

        void ScheduleScan()
        {
            scanTimer.Stop();
            scanTimer.Start();
        }

        void Scan()
        {
            scanResult = new List<DocTab>();
            Native.EnumWindows(enumProc, IntPtr.Zero);
            var found = new Dictionary<IntPtr, DocTab>();
            foreach (var t in scanResult) found[t.Handle] = t;

            bool changed = false;
            for (int i = tabs.Count - 1; i >= 0; i--)
            {
                DocTab now;
                if (!found.TryGetValue(tabs[i].Handle, out now))
                {
                    IntPtr gone = tabs[i].Handle;
                    tabs.RemoveAt(i);
                    changed = true;
                    if (gone == closeIntent)
                    {
                        // 用户关掉了整个窗口：组里其余文档依次关闭
                        closeIntent = IntPtr.Zero;
                        var rest = new List<IntPtr>();
                        foreach (var t in tabs) rest.Add(t.Handle);
                        StartCascade(rest);
                    }
                    continue;
                }
                if (now.FullTitle != tabs[i].FullTitle)
                {
                    tabs[i].Title = now.Title;
                    tabs[i].FullTitle = now.FullTitle;
                    changed = true;
                }
                found.Remove(tabs[i].Handle);
            }
            // EnumWindows 按 Z 序从上到下，倒序追加让较早的窗口排在前面
            for (int i = scanResult.Count - 1; i >= 0; i--)
                if (found.ContainsKey(scanResult[i].Handle)) { tabs.Add(scanResult[i]); changed = true; }

            var live = new HashSet<uint>();
            foreach (var t in tabs) { uint pid; Native.GetWindowThreadProcessId(t.Handle, out pid); live.Add(pid); }
            foreach (var pid in new List<uint>(procNames.Keys)) if (!live.Contains(pid)) procNames.Remove(pid);

            // 离开分组的窗口（如文档已转到另一个桌面）恢复正常的最大化位置
            foreach (var h in new List<IntPtr>(shifted))
                if (IndexOf(h) < 0) { shifted.Remove(h); if (Native.IsWindow(h)) SetFrameTop(h, 0); }
            pseudoMax.RemoveWhere(h => IndexOf(h) < 0);
            noShift.RemoveWhere(h => IndexOf(h) < 0);
            foreach (var t in tabs) if (!pseudoMax.Contains(t.Handle)) EnsureShift(t.Handle);

            if (cascadeCur != IntPtr.Zero && IndexOf(cascadeCur) < 0) CascadeNext();
            // 长时间没关掉（保存提示一直没处理）就放弃，避免之后误关
            if (cascadeCur != IntPtr.Zero && (DateTime.Now - cascadeStepTime).TotalMinutes > 3) StopCascade();
            if (closeIntent != IntPtr.Zero && (DateTime.Now - closeIntentTime).TotalMinutes > 3) closeIntent = IntPtr.Zero;
            if (IndexOf(curDoc) < 0) curDoc = IntPtr.Zero;
            // 还没有文档到过前台时，以 Z 序最上面的文档作为当前文档
            if (curDoc == IntPtr.Zero && scanResult.Count > 0)
            {
                curDoc = scanResult[0].Handle;
                if (!hasGroup) CaptureGroup(curDoc);
                changed = true;
            }
            IntPtr fg = Native.GetForegroundWindow();
            if (IndexOf(fg) >= 0 && fg != curDoc) { OnDocForeground(fg); changed = true; }

            if (changed) LayoutTabs();
            UpdatePlacement();
            if (changed) Invalidate();
        }

        bool EnumCallback(IntPtr h, IntPtr l)
        {
            if (!Native.IsWindowVisible(h) && !hiddenByUs.Contains(h)) return true;
            if (Native.GetWindow(h, Native.GW_OWNER) != IntPtr.Zero) return true;
            DocKind kind;
            if (!TryKind(Native.GetClassNameOf(h), out kind)) return true;
            if (Native.IsCloaked(h)) return true;
            string title = Native.GetWindowTextOf(h);
            if (string.IsNullOrEmpty(title)) return true;
            // 只认 Office 本体，WPS 等兼容程序可能使用相同的窗口类名
            string exe = kind == DocKind.Word ? "WINWORD.EXE" : "EXCEL.EXE";
            if (!string.Equals(ProcessName(h), exe, StringComparison.OrdinalIgnoreCase)) return true;
            scanResult.Add(new DocTab { Handle = h, Kind = kind, Title = CleanTitle(title, kind), FullTitle = title });
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

        static string CleanTitle(string t, DocKind kind)
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

        int IndexOf(IntPtr h)
        {
            if (h == IntPtr.Zero) return -1;
            for (int i = 0; i < tabs.Count; i++) if (tabs[i].Handle == h) return i;
            return -1;
        }

        void OnWinEvent(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            if (hwnd == IntPtr.Zero || idObject != Native.OBJID_WINDOW || idChild != 0) return;
            switch (ev)
            {
                case Native.EVENT_SYSTEM_MOVESIZESTART:
                    if (hwnd == curDoc) BeginWindowDrag();
                    return;
                case Native.EVENT_SYSTEM_MOVESIZEEND:
                    if (windowDrag) EndWindowDrag();
                    return;
                case Native.EVENT_OBJECT_LOCATIONCHANGE:
                    if (IndexOf(hwnd) >= 0) OnDocLocation(hwnd);
                    return;
                case Native.EVENT_SYSTEM_FOREGROUND:
                    // 保存提示里点了“取消”，文档重新回到前台：不再整组关闭
                    if (hwnd == closeIntent && (DateTime.Now - closeIntentTime).TotalMilliseconds > 400) closeIntent = IntPtr.Zero;
                    if (hwnd == cascadeCur && (DateTime.Now - cascadeStepTime).TotalMilliseconds > 1500) StopCascade();
                    if (IndexOf(hwnd) >= 0) OnDocForeground(hwnd);
                    UpdatePlacement();
                    ScheduleScan();
                    return;
                case Native.EVENT_SYSTEM_MINIMIZESTART:
                    // 组内任一文档被最小化就全组最小化（前台可能已先切到另一个文档，不能只看 curDoc）
                    if (IndexOf(hwnd) >= 0) MinimizeOthers(hwnd);
                    UpdatePlacement();
                    return;
                case Native.EVENT_OBJECT_DESTROY:
                    if (IndexOf(hwnd) >= 0) ScheduleScan();
                    return;
            }
            if (Native.GetAncestor(hwnd, Native.GA_ROOT) != hwnd) return;
            DocKind kind;
            if (TryKind(Native.GetClassNameOf(hwnd), out kind)) ScheduleScan();
        }

        // ---------------- 同组窗口联动 ----------------

        void OnDocForeground(IntPtr h)
        {
            if (h == curDoc) return;
            curDoc = h;
            // 已被最小化（例如全组最小化时前台恰好切到它）就不要再把它恢复回来
            if (Native.IsIconic(h)) { Invalidate(); return; }
            if (!hasGroup) CaptureGroup(h);
            else ApplyGroupState(h, true);
            Invalidate();
        }

        void CaptureGroup(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            groupMax = IsMax(h);
            groupNormal = Native.GetPlacement(h).rcNormalPosition;
            hasGroup = true;
        }

        // 让窗口 h 与分组状态一致：最大化 / 相同的还原位置
        void ApplyGroupState(IntPtr h, bool activating)
        {
            if (!hasGroup)
            {
                if (Native.IsIconic(h)) Native.ShowWindow(h, Native.SW_RESTORE);
                return;
            }
            if (groupMax)
            {
                if (noShift.Contains(h)) { if (!IsPseudoMax(h)) PseudoMaximize(h); return; }
                if (!Native.IsZoomed(h)) Native.ShowWindow(h, Native.SW_MAXIMIZE);
                EnsureShift(h);
                return;
            }
            var wp = Native.GetPlacement(h);
            if (Native.IsZoomed(h) || Native.IsIconic(h) || !wp.rcNormalPosition.Equals(groupNormal))
            {
                wp.flags = 0;
                wp.showCmd = activating ? Native.SW_SHOWNORMAL : Native.SW_SHOWNOACTIVATE;
                wp.rcNormalPosition = groupNormal;
                Native.SetWindowPlacement(h, ref wp);
                shifted.Remove(h);
                pseudoMax.Remove(h);
            }
        }

        void OnDocLocation(IntPtr h)
        {
            if (Native.IsIconic(h)) return;
            bool zoomed = Native.IsZoomed(h);
            bool wasPseudo = pseudoMax.Contains(h);
            if (h != curDoc || h != Native.GetForegroundWindow())
            {
                // 后台文档（多半是我们自己同步过去的）只需保证最大化时让出标签栏
                if (zoomed) EnsureShift(h);
                return;
            }
            if (zoomed && wasPseudo && hasGroup)
            {
                // 伪最大化窗口显示的是“最大化”按钮，用户点它（或双击标题栏）就当作“还原”
                pseudoMax.Remove(h);
                groupMax = false;
                ApplyGroupState(h, true);
                syncTimer.Stop(); syncTimer.Start();
                UpdatePlacement();
                return;
            }
            if (zoomed)
            {
                EnsureShift(h);
                groupMax = true;
            }
            else if (IsPseudoMax(h))
            {
                groupMax = true;
            }
            else
            {
                shifted.Remove(h);
                var np = Native.GetPlacement(h).rcNormalPosition;
                bool wasMax = groupMax;
                bool changed = groupMax || !np.Equals(groupNormal);
                groupMax = false;
                groupNormal = np;
                if (changed)
                {
                    // 拖动/缩放过程中实时带动其他文档，结束后再由 syncTimer 精确校正
                    if (!windowDrag)
                    {
                        if (!wasMax) LiveSync(h);
                        syncTimer.Stop(); syncTimer.Start();
                    }
                }
            }
            hasGroup = true;
            UpdatePlacement();
        }

        void LiveSync(IntPtr h)
        {
            Native.RECT r;
            if (!Native.GetWindowRect(h, out r)) return;
            Native.RECT f = Native.FrameBounds(h, r);
            foreach (var t in tabs)
            {
                IntPtr o = t.Handle;
                if (o == h || Native.IsIconic(o) || Native.IsZoomed(o) || pseudoMax.Contains(o)) continue;
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
            if (!hasGroup || groupMax || windowDrag) return;
            foreach (var t in tabs)
            {
                IntPtr h = t.Handle;
                if (h == curDoc || Native.IsIconic(h) || hiddenByUs.Contains(h)) continue;
                ApplyGroupState(h, false);
            }
        }

        void BeginWindowDrag()
        {
            windowDrag = true;
            windowDragStart = DateTime.Now;
            revealTimer.Stop();
            foreach (var t in tabs)
            {
                IntPtr o = t.Handle;
                if (o == curDoc || Native.IsIconic(o) || !Native.IsWindowVisible(o)) continue;
                hiddenByUs.Add(o);
                Native.ShowWindowAsync(o, Native.SW_HIDE);
            }
        }

        void EndWindowDrag()
        {
            windowDrag = false;
            revealTimer.Stop();
            revealTimer.Start();
        }

        void RevealHidden()
        {
            if (hiddenByUs.Count == 0) return;
            foreach (var o in new List<IntPtr>(hiddenByUs))
            {
                if (!Native.IsWindow(o)) continue;
                if (hasGroup && !groupMax)
                {
                    // 先摆到与当前文档相同的位置，再显示
                    var wp = Native.GetPlacement(o);
                    wp.flags = 0;
                    wp.showCmd = Native.SW_SHOWNOACTIVATE;
                    wp.rcNormalPosition = groupNormal;
                    Native.SetWindowPlacement(o, ref wp);
                    shifted.Remove(o);
                    pseudoMax.Remove(o);
                }
                else
                {
                    Native.ShowWindow(o, Native.SW_SHOWNOACTIVATE);
                }
                // 放在当前文档后面，不要盖住它
                if (curDoc != IntPtr.Zero && Native.IsWindow(curDoc))
                    Native.SetWindowPos(o, curDoc, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE | Native.SWP_NOOWNERZORDER);
            }
            hiddenByUs.Clear();
            ScheduleScan();
        }

        void MinimizeOthers(IntPtr h)
        {
            foreach (var t in new List<DocTab>(tabs))
                if (t.Handle != h && !Native.IsIconic(t.Handle))
                    Native.ShowWindow(t.Handle, Native.SW_SHOWMINNOACTIVE);
        }

        // 最大化的文档窗口整体下移，把顶部让给标签栏
        void EnsureShift(IntPtr h)
        {
            if (!Native.IsZoomed(h)) return;
            if (!SetFrameTop(h, barHeight)) return;
            shifted.Add(h);
            // 个别程序不允许移动最大化窗口，改为“伪最大化”：还原后铺满标签栏以下的区域
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            Native.RECT f = Native.FrameBounds(h, r);
            if (f.Top < Screen.FromHandle(h).WorkingArea.Top + barHeight / 2)
            {
                noShift.Add(h);
                PseudoMaximize(h);
            }
        }

        Rectangle PseudoMaxFrame(IntPtr h)
        {
            Rectangle wa = Screen.FromHandle(h).WorkingArea;
            return Rectangle.FromLTRB(wa.Left, wa.Top + barHeight, wa.Right, wa.Bottom);
        }

        void PseudoMaximize(IntPtr h)
        {
            shifted.Remove(h);
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
            pseudoMax.Add(h);
        }

        bool IsPseudoMax(IntPtr h)
        {
            if (!pseudoMax.Contains(h)) return false;
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
            pseudoMax.Remove(h);
            return false;
        }

        bool IsMax(IntPtr h)
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

        // ---------------- 标签栏位置 ----------------

        void UpdatePlacement()
        {
            IntPtr fg = Native.GetForegroundWindow();
            BarMode m;
            Rectangle rect = Rectangle.Empty;
            if (curDoc == IntPtr.Zero || tabs.Count == 0 || !Native.IsWindow(curDoc) || Native.IsIconic(curDoc) ||
                !Native.IsWindowVisible(curDoc) || (fg == curDoc && IsFullscreen(fg)))
            {
                m = BarMode.Hidden;
            }
            else if (IsMax(curDoc))
            {
                Rectangle wa = Screen.FromHandle(curDoc).WorkingArea;
                m = BarMode.Maximized;
                rect = new Rectangle(wa.Left, wa.Top, wa.Width, barHeight);
            }
            else
            {
                Native.RECT wr;
                Native.GetWindowRect(curDoc, out wr);
                Native.RECT f = Native.FrameBounds(curDoc, wr);
                Rectangle wa = Screen.FromHandle(curDoc).WorkingArea;
                int w = Math.Max(1, Math.Min(f.Right - f.Left, DesiredWidth()));
                if (f.Top - barHeight >= wa.Top)
                {
                    m = BarMode.Follow;
                    // 向下压住标题栏顶部的边框线，让当前标签和标题栏连成一体
                    int overlap = (int)Math.Round(2 * s);
                    rect = new Rectangle(f.Left, f.Top - barHeight + overlap, w, barHeight);
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
                return;
            }

            // 文档在前台时置顶，Office 被点击激活也盖不住标签栏；切到其他程序后取消置顶，只紧贴在文档上面一层
            bool docInFront = IndexOf(fg) >= 0 || fg == Handle;
            bool wantTopmost = m == BarMode.AutoHide || docInFront;
            if (!Visible) Show();
            if (wantTopmost)
            {
                Native.SetWindowPos(Handle, Native.HWND_TOPMOST, rect.Left, rect.Top, rect.Width, rect.Height, Native.SWP_NOACTIVATE);
            }
            else if (isTopmost)
            {
                Native.SetWindowPos(Handle, Native.HWND_NOTOPMOST, rect.Left, rect.Top, rect.Width, rect.Height, Native.SWP_NOACTIVATE);
                IntPtr above = Native.GetWindow(curDoc, Native.GW_HWNDPREV);
                if (above != Handle)
                    Native.SetWindowPos(Handle, above, 0, 0, 0, 0,
                        Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE);
            }
            else
            {
                Native.SetWindowPos(Handle, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height,
                    Native.SWP_NOACTIVATE | Native.SWP_NOZORDER);
            }
            isTopmost = wantTopmost;
            if (modeChanged) { LayoutTabs(); Invalidate(); }
        }

        int DesiredWidth()
        {
            int n = Math.Max(1, tabs.Count);
            return (int)Math.Ceiling(12 * s + n * 210 * s + (n - 1) * 2 * s);
        }

        bool Floating
        {
            get { return mode == BarMode.Follow || mode == BarMode.AutoHide; }
        }

        // 小窗口且上方没空间时：鼠标停在窗口上沿约 0.2 秒滑出标签栏，离开 0.5 秒后收起
        void OnHoverTick()
        {
            if (mode != BarMode.AutoHide || curDoc == IntPtr.Zero) { hoverTimer.Stop(); return; }
            Point p = Cursor.Position;
            if (!overlayShown)
            {
                Native.RECT wr;
                Native.GetWindowRect(curDoc, out wr);
                Native.RECT f = Native.FrameBounds(curDoc, wr);
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
                leaveTicks = (r.Contains(p) || pressIndex >= 0) ? 0 : leaveTicks + 1;
                if (leaveTicks >= 5) { overlayShown = false; hoverTicks = 0; UpdatePlacement(); }
            }
        }

        void OnDisplayChanged(object sender, EventArgs e)
        {
            foreach (var t in tabs) EnsureShift(t.Handle);
            UpdatePlacement();
        }

        // ---------------- 绘制 ----------------

        void LayoutTabs()
        {
            float pad = Floating ? 6 * s : 4 * s, gap = Floating ? 2 * s : 1 * s, top = 4 * s;
            float h = barHeight - top;
            int n = tabs.Count;
            if (n == 0) return;
            float avail = ClientSize.Width - pad * 2;
            float w = Math.Max(1, Math.Min(210 * s, (avail - gap * (n - 1)) / n));
            float x = pad;
            foreach (var t in tabs)
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

            for (int i = 0; i < tabs.Count; i++)
            {
                var t = tabs[i];
                bool active = t.Handle == curDoc;
                bool hover = i == hoverIndex;
                var b = t.Bounds;

                if (active || hover || floating)
                {
                    using (var path = TopRounded(b, 6 * s))
                    using (var br = new SolidBrush(active ? ActiveBg : hover ? HoverBg : BarBg))
                        g.FillPath(br, path);
                }
                else if (i + 1 < tabs.Count && tabs[i + 1].Handle != curDoc && i + 1 != hoverIndex)
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
                int ai = IndexOf(curDoc);
                if (ai < 0) g.DrawLine(pen, 0, y, ClientSize.Width, y);
                else
                {
                    var ab = tabs[ai].Bounds;
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
            for (int i = 0; i < tabs.Count; i++) if (tabs[i].Bounds.Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (pressIndex >= 0 && e.Button == MouseButtons.Left)
            {
                if (!dragging && Math.Abs(e.X - pressPoint.X) > 6 * s) dragging = true;
                if (dragging)
                {
                    int target = HitTest(new Point(e.X, (int)tabs[pressIndex].Bounds.Top + 2));
                    if (target >= 0 && target != pressIndex)
                    {
                        var t = tabs[pressIndex];
                        tabs.RemoveAt(pressIndex);
                        tabs.Insert(target, t);
                        pressIndex = target;
                        hoverIndex = target;
                        LayoutTabs();
                        Invalidate();
                    }
                    return;
                }
            }
            int idx = HitTest(e.Location);
            bool overClose = idx >= 0 && tabs[idx].CloseBounds.Contains(e.Location);
            if (idx != hoverIndex || overClose != hoverClose)
            {
                if (idx != hoverIndex) tip.SetToolTip(this, idx >= 0 ? tabs[idx].FullTitle : null);
                hoverIndex = idx;
                hoverClose = overClose;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (pressIndex < 0) { hoverIndex = -1; hoverClose = false; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int idx = HitTest(e.Location);
            if (e.Button == MouseButtons.Left && idx >= 0) { pressIndex = idx; pressPoint = e.Location; dragging = false; }
            else if (e.Button == MouseButtons.Middle && idx >= 0) CloseTab(tabs[idx]);
            else if (e.Button == MouseButtons.Right) ShowTabMenu(idx, e.Location);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            int pressed = pressIndex;
            bool wasDrag = dragging;
            pressIndex = -1;
            dragging = false;
            if (wasDrag || pressed < 0 || pressed >= tabs.Count) return;
            var t = tabs[pressed];
            if (!t.Bounds.Contains(e.Location)) return;
            if (t.CloseBounds.Contains(e.Location) && (t.Handle == curDoc || pressed == hoverIndex)) CloseTab(t);
            else ActivateTab(t);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (tabs.Count == 0) return;
            int cur = IndexOf(curDoc);
            int next = e.Delta > 0 ? cur - 1 : cur + 1;
            if (cur < 0) next = 0;
            next = (next + tabs.Count) % tabs.Count;
            ActivateTab(tabs[next]);
        }

        void ShowTabMenu(int idx, Point p)
        {
            var items = new List<MenuItem>();
            if (idx >= 0)
            {
                var t = tabs[idx];
                items.Add(new MenuItem("关闭", delegate { CloseTab(t); }));
                items.Add(new MenuItem("关闭其他标签", delegate
                {
                    var rest = new List<IntPtr>();
                    foreach (var o in tabs) if (o != t) rest.Add(o.Handle);
                    StartCascade(rest);
                }));
                items.Add(new MenuItem("关闭全部", delegate
                {
                    var all = new List<IntPtr>();
                    foreach (var o in tabs) if (o != t) all.Add(o.Handle);
                    all.Add(t.Handle);
                    StartCascade(all);
                }));
                items.Add(new MenuItem("-"));
            }
            items.Add(new MenuItem("退出 DocTabs", delegate { ExitApp(); }));
            var menu = new ContextMenu(items.ToArray());
            Native.SetForegroundWindow(Handle);
            menu.Show(this, p);
        }

        // ---------------- 窗口操作 ----------------

        void ActivateTab(DocTab t)
        {
            IntPtr h = t.Handle;
            // 先把目标文档调成与当前文档相同的状态/位置，再切过去，看起来就像同一个窗口换了内容
            ApplyGroupState(h, true);
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
            curDoc = h;
            UpdatePlacement();
            Invalidate();
        }

        void CloseTab(DocTab t)
        {
            // Office 会自行弹出“是否保存”提示
            Native.PostMessage(t.Handle, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            ScheduleScan();
        }

        // ---------------- 托盘 / 自启动 / 退出 ----------------

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

        Icon MakeTrayIcon()
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

        void RestoreShifted()
        {
            foreach (var h in shifted) if (Native.IsWindow(h)) SetFrameTop(h, 0);
            shifted.Clear();
        }

        void ExitApp()
        {
            fallbackTimer.Stop();
            scanTimer.Stop();
            syncTimer.Stop();
            hoverTimer.Stop();
            foreach (var h in hooks) if (h != IntPtr.Zero) Native.UnhookWinEvent(h);
            if (mouseHook != IntPtr.Zero) Native.UnhookWindowsHookEx(mouseHook);
            if (keyHook != IntPtr.Zero) Native.UnhookWindowsHookEx(keyHook);
            mouseHook = keyHook = IntPtr.Zero;
            hooks.Clear();
            SystemEvents.DisplaySettingsChanged -= OnDisplayChanged;
            windowDrag = false;
            RevealHidden();
            RestoreShifted();
            tray.Visible = false;
            tray.Dispose();
            Close();
            Application.Exit();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            windowDrag = false;
            RevealHidden();
            RestoreShifted();
            base.OnFormClosed(e);
        }
    }

    static class Native
    {
        public const int WS_EX_TOOLWINDOW = 0x80, WS_EX_NOACTIVATE = 0x08000000;
        public const int WS_CAPTION = 0x00C00000, GWL_STYLE = -16;
        public const int WM_MOUSEACTIVATE = 0x21, MA_NOACTIVATE = 3, WM_CLOSE = 0x10;
        public const uint GW_OWNER = 4, GW_HWNDPREV = 3, GA_ROOT = 2;
        public const int SW_SHOWNORMAL = 1, SW_MAXIMIZE = 3, SW_SHOWNOACTIVATE = 4, SW_SHOWMINNOACTIVE = 7, SW_RESTORE = 9;
        public const uint SWP_NOSIZE = 1, SWP_NOMOVE = 2, SWP_NOZORDER = 4, SWP_NOACTIVATE = 0x10, SWP_NOOWNERZORDER = 0x200, SWP_ASYNCWINDOWPOS = 0x4000;
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1), HWND_NOTOPMOST = new IntPtr(-2);
        public const uint EVENT_SYSTEM_FOREGROUND = 3, EVENT_SYSTEM_MOVESIZESTART = 0x0A, EVENT_SYSTEM_MOVESIZEEND = 0x0B;
        public const uint EVENT_SYSTEM_MINIMIZESTART = 0x16, EVENT_SYSTEM_MINIMIZEEND = 0x17;
        public const int SW_HIDE = 0;
        public const uint EVENT_OBJECT_DESTROY = 0x8001, EVENT_OBJECT_HIDE = 0x8003;
        public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B, EVENT_OBJECT_NAMECHANGE = 0x800C;
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

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData, flags, time; public IntPtr extra; }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr extra; }

        [DllImport("user32.dll")] public static extern IntPtr SetWindowsHookEx(int id, LowLevelProc proc, IntPtr mod, uint thread);
        [DllImport("user32.dll")] public static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetModuleHandle(string name);

        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
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
