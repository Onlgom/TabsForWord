using System;
using System.Runtime.InteropServices;

namespace TabsForWord
{
    internal static class NativeMethods
    {
        // Window.Activate() makes a window active inside Word object model, but in
        // SDI Word it does not always bring the top-level window to the front.
        // The click happens in our process => the process is allowed to set foreground.
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        // SetForegroundWindow can fail silently during activation transitions
        // (Windows holds a foreground lock). BringWindowToTop additionally raises
        // the window in the z-order - together they are reliable.
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool BringWindowToTop(IntPtr hWnd);

        // ------------------------------------------------------------------
        // Local keyboard hook (Ctrl+Tab): Word UI thread only
        // ------------------------------------------------------------------

        internal const int WH_KEYBOARD = 2;

        // Hook codes. HC_ACTION means the message is being REMOVED from the queue;
        // HC_NOREMOVE means someone is peeking at it and it will arrive again. Acting
        // on a peek would switch the tab twice for one keypress.
        internal const int HC_ACTION = 0;

        internal const int VK_TAB = 0x09;
        internal const int VK_SHIFT = 0x10;
        internal const int VK_CONTROL = 0x11;
        internal const int VK_MENU = 0x12;

        internal delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        internal static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        internal static extern short GetKeyState(int nVirtKey);

        // ------------------------------------------------------------------
        // Look and feel of the all-tabs popup
        // ------------------------------------------------------------------

        internal const int CS_DROPSHADOW = 0x00020000;
        internal const int WS_EX_TOOLWINDOW = 0x00000080;

        private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        private const int DWMWCP_ROUND = 2;

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        /// <summary>Rounds the window corners (Windows 11; silently ignored on Windows 10).</summary>
        internal static void TryRoundCorners(IntPtr hwnd)
        {
            try
            {
                int pref = DWMWCP_ROUND;
                DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));
            }
            catch { }
        }

        private const int EM_SETCUEBANNER = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        /// <summary>Grey hint text in an empty search box.</summary>
        internal static void TrySetCueBanner(IntPtr textBoxHandle, string text)
        {
            try
            {
                SendMessage(textBoxHandle, EM_SETCUEBANNER, (IntPtr)1, text);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsProcessDPIAware();

        /// <summary>DPI awareness of the Word process (for the diagnostic log).</summary>
        internal static bool IsProcessDpiAwareSafe()
        {
            try { return IsProcessDPIAware(); }
            catch { return false; }
        }
    }
}
