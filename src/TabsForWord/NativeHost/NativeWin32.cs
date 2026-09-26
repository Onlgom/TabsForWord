using System;
using System.Runtime.InteropServices;
using System.Text;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// Win32 P/Invoke for the in-window tab host.
    /// DECLARATIONS and thin safe wrappers ONLY - no business logic here.
    ///
    /// Every HWND is an IntPtr (never int: truncating a pointer on x64 is not allowed).
    /// GWL access goes through the Get/SetWindowLongPtrSafe wrappers, which are correct
    /// in both 32- and 64-bit processes (32-bit user32 has no *LongPtr exports).
    /// </summary>
    internal static class NativeWin32
    {
        // ------------------------------------------------------------------
        // Structures
        // ------------------------------------------------------------------

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public int Width { get { return Right - Left; } }
            public int Height { get { return Bottom - Top; } }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int X;
            public int Y;
        }

        // ------------------------------------------------------------------
        // Window enumeration
        // ------------------------------------------------------------------

        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // ------------------------------------------------------------------
        // Geometry and coordinates
        // ------------------------------------------------------------------

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        // ------------------------------------------------------------------
        // Hierarchy and state
        // ------------------------------------------------------------------

        internal const uint GA_PARENT = 1;
        internal const uint GA_ROOT = 2;
        internal const uint GA_ROOTOWNER = 3;   // the root, then up the owner chain (a dialog -> its Word window)

        [DllImport("user32.dll")]
        internal static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowEnabled(IntPtr hWnd);

        // ------------------------------------------------------------------
        // Window management
        // ------------------------------------------------------------------

        internal static readonly IntPtr HWND_TOP = IntPtr.Zero;

        internal const uint SWP_NOSIZE = 0x0001;
        internal const uint SWP_NOMOVE = 0x0002;
        internal const uint SWP_NOZORDER = 0x0004;
        internal const uint SWP_NOACTIVATE = 0x0010;
        internal const uint SWP_SHOWWINDOW = 0x0040;

        internal const int SW_HIDE = 0;

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        internal static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        // ------------------------------------------------------------------
        // Activation (tab switching) and input state
        // ------------------------------------------------------------------

        internal const int WM_SYSCOMMAND = 0x0112;
        internal const int SC_RESTORE = 0xF120;

        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        /// <summary>Screen position of the cursor for the message being processed (packed x/y).</summary>
        [DllImport("user32.dll")]
        internal static extern uint GetMessagePos();

        [StructLayout(LayoutKind.Sequential)]
        private struct POINTER_INFO
        {
            public uint pointerType;
            public uint pointerId;
            public uint frameId;
            public uint pointerFlags;
            public IntPtr sourceDevice;
            public IntPtr hwndTarget;
            public POINT ptPixelLocation;
            public POINT ptHimetricLocation;
            public POINT ptPixelLocationRaw;
            public POINT ptHimetricLocationRaw;
            public uint dwTime;
            public uint historyCount;
            public int InputData;
            public uint dwKeyStates;
            public ulong PerformanceCount;
            public int ButtonChangeType;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetPointerInfo(uint pointerId, ref POINTER_INFO pointerInfo);

        /// <summary>Screen position of a touch/pen pointer (WM_POINTERACTIVATE); false if unavailable.</summary>
        internal static bool TryGetPointerPixelLocation(uint pointerId, out System.Drawing.Point point)
        {
            point = System.Drawing.Point.Empty;
            try
            {
                var info = new POINTER_INFO();
                if (!GetPointerInfo(pointerId, ref info)) return false;
                point = new System.Drawing.Point(info.ptPixelLocation.X, info.ptPixelLocation.Y);
                return true;
            }
            catch
            {
                return false;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        /// <summary>
        /// Environment.TickCount-compatible time of the last keyboard or mouse input in
        /// the session (any process); null if the call fails.
        /// </summary>
        internal static int? GetLastInputTickSafe()
        {
            try
            {
                var lii = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO)) };
                if (!GetLastInputInfo(ref lii)) return null;
                return unchecked((int)lii.dwTime);
            }
            catch
            {
                return null;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public uint cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        private const uint GUI_INMOVESIZE = 0x00000002;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        /// <summary>
        /// Is the left or right mouse button physically held right now (any window, any
        /// process)? Both are checked, so swapped buttons do not matter.
        /// </summary>
        internal static bool IsMouseButtonDownSafe()
        {
            try
            {
                return ((GetAsyncKeyState(0x01) | GetAsyncKeyState(0x02)) & 0x8000) != 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Is the calling thread (Word UI thread) inside the modal move/size loop right
        /// now - the user is dragging a window or its frame with the mouse? Stateless:
        /// asked at the moment, so nothing can get stuck if an end event is missed.
        /// </summary>
        internal static bool IsInMoveSizeLoopSafe()
        {
            try
            {
                var gti = new GUITHREADINFO { cbSize = (uint)Marshal.SizeOf(typeof(GUITHREADINFO)) };
                if (!GetGUIThreadInfo(GetCurrentThreadId(), ref gti)) return false;
                return (gti.flags & GUI_INMOVESIZE) != 0;
            }
            catch
            {
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Window styles (x64-safe wrappers)
        // ------------------------------------------------------------------

        internal const int GWL_STYLE = -16;
        internal const int GWL_EXSTYLE = -20;

        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        /// <summary>GetWindowLongPtr that also works in a 32-bit process (and does not truncate in a 64-bit one).</summary>
        internal static IntPtr GetWindowLongPtrSafe(IntPtr hWnd, int nIndex)
        {
            return IntPtr.Size == 8
                ? GetWindowLongPtr64(hWnd, nIndex)
                : new IntPtr(GetWindowLong32(hWnd, nIndex));
        }

        /// <summary>SetWindowLongPtr that also works in a 32-bit process.</summary>
        internal static IntPtr SetWindowLongPtrSafe(IntPtr hWnd, int nIndex, IntPtr value)
        {
            return IntPtr.Size == 8
                ? SetWindowLongPtr64(hWnd, nIndex, value)
                : new IntPtr(SetWindowLong32(hWnd, nIndex, value.ToInt32()));
        }

        // ------------------------------------------------------------------
        // DPI
        // ------------------------------------------------------------------

        [DllImport("user32.dll", EntryPoint = "GetDpiForWindow")]
        private static extern uint GetDpiForWindowNative(IntPtr hWnd);

        private static bool _dpiApiMissing;

        /// <summary>
        /// DPI of a particular window (Windows 10 1607+). On older systems, or on
        /// failure, returns the fallback (usually the control DeviceDpi, or 96).
        /// </summary>
        internal static int GetDpiForWindowSafe(IntPtr hWnd, int fallback)
        {
            if (_dpiApiMissing) return fallback;
            try
            {
                uint dpi = GetDpiForWindowNative(hWnd);
                return dpi >= 48 && dpi <= 768 ? (int)dpi : fallback;
            }
            catch (EntryPointNotFoundException)
            {
                _dpiApiMissing = true;
                return fallback;
            }
            catch
            {
                return fallback;
            }
        }

        // Word juggles the DPI context of its own UI thread (Office mixed-mode DPI
        // hosting): our callbacks may run in a system-aware context where every window
        // coordinate is VIRTUALISED (scaled to the system DPI). All strip geometry has
        // to be computed in physical pixels - so for the duration of the calculation the
        // thread is switched to per-monitor-v2 (Windows 10 1703+).
        internal static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

        [DllImport("user32.dll", EntryPoint = "SetThreadDpiAwarenessContext")]
        private static extern IntPtr SetThreadDpiAwarenessContextNative(IntPtr dpiContext);

        private static bool _threadDpiApiMissing;

        /// <summary>
        /// Switches the DPI context of the thread; returns the previous one (IntPtr.Zero
        /// means the API is unavailable or the call failed: older Windows keep the old behaviour).
        /// </summary>
        internal static IntPtr SetThreadDpiAwarenessContextSafe(IntPtr dpiContext)
        {
            if (_threadDpiApiMissing) return IntPtr.Zero;
            try
            {
                return SetThreadDpiAwarenessContextNative(dpiContext);
            }
            catch (EntryPointNotFoundException)
            {
                _threadDpiApiMissing = true;
                return IntPtr.Zero;
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        // ------------------------------------------------------------------
        // WinEvent hook (re-layout when Word windows move or are resized).
        // This is neither a global SetWindowsHookEx hook nor DLL injection:
        // WINEVENT_OUTOFCONTEXT plus a filter on our own process id; the callback
        // arrives asynchronously on the installing thread through its message queue.
        // ------------------------------------------------------------------

        internal const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
        internal const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        internal const int OBJID_WINDOW = 0;

        internal delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate pfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentProcessId();

        // ------------------------------------------------------------------
        // Reading the Word background colour (auto-matching the strip background)
        // ------------------------------------------------------------------

        internal const uint CLR_INVALID = 0xFFFFFFFF;

        [DllImport("user32.dll")]
        internal static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        internal static extern uint GetPixel(IntPtr hdc, int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        // ------------------------------------------------------------------
        // Safe string wrappers
        // ------------------------------------------------------------------

        internal static string GetClassNameSafe(IntPtr hWnd)
        {
            try
            {
                var sb = new StringBuilder(256);
                int n = GetClassName(hWnd, sb, sb.Capacity);
                return n > 0 ? sb.ToString() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal static string GetWindowTextSafe(IntPtr hWnd, int maxLength)
        {
            try
            {
                var sb = new StringBuilder(Math.Max(2, maxLength + 1));
                int n = GetWindowText(hWnd, sb, sb.Capacity);
                return n > 0 ? sb.ToString() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
