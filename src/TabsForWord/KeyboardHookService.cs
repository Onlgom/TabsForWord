using System;

namespace TabsForWord
{
    /// <summary>
    /// Ctrl+Tab / Ctrl+Shift+Tab - next / previous tab (specification section 8).
    ///
    /// A local WH_KEYBOARD hook on Word UI thread only (not global, not
    /// low-level): only keys addressed to Word own windows are intercepted.
    /// The delegate lives in a field - the GC must not collect it while the hook is set.
    ///
    /// Known limitation: Word own Ctrl+Tab (a tab character inside a table cell)
    /// is shadowed - documented in docs/KNOWN_ISSUES.md.
    /// </summary>
    public sealed class KeyboardHookService : IDisposable
    {
        private readonly Action _nextTab;
        private readonly Action _prevTab;
        private readonly Func<bool> _shouldHandle;   // null = always; false = the key goes on untouched
        private NativeMethods.HookProc _proc;   // GC root
        private IntPtr _hook = IntPtr.Zero;

        public KeyboardHookService(Action nextTab, Action prevTab, Func<bool> shouldHandle = null)
        {
            _nextTab = nextTab;
            _prevTab = prevTab;
            _shouldHandle = shouldHandle;
        }

        /// <summary>Installs the hook. Must be called from Word UI thread.</summary>
        public void Install()
        {
            if (_hook != IntPtr.Zero) return;
            try
            {
                _proc = HookCallback;
                _hook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD, _proc,
                    IntPtr.Zero, NativeMethods.GetCurrentThreadId());
                LoggingService.Info(_hook != IntPtr.Zero
                    ? "Keyboard hook installed (Ctrl+Tab)"
                    : "Keyboard hook install failed (SetWindowsHookEx returned NULL)");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Keyboard hook install failed", ex);
            }
        }

        private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                // HC_ACTION only: on HC_NOREMOVE the message is merely being peeked at
                // and will come back, so handling it there would fire the switch twice.
                if (code == NativeMethods.HC_ACTION && wParam.ToInt32() == NativeMethods.VK_TAB)
                {
                    bool ctrl = (NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) & 0x8000) != 0;
                    bool alt = (NativeMethods.GetKeyState(NativeMethods.VK_MENU) & 0x8000) != 0;
                    if (ctrl && !alt && (_shouldHandle == null || _shouldHandle()))
                    {
                        // Bit 31 of lParam: 0 = key down, 1 = key up. We act on key down
                        // (auto-repeat included) and swallow the key up as well.
                        bool keyUp = (lParam.ToInt64() & 0x80000000L) != 0;
                        if (!keyUp)
                        {
                            bool shift = (NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0;
                            var action = shift ? _prevTab : _nextTab;
                            if (action != null) action();
                        }
                        return (IntPtr)1; // the key is not passed on to Word
                    }
                }
            }
            catch (Exception ex)
            {
                // An exception escaping a hook would bring Word down - always swallow.
                LoggingService.Error("Keyboard hook callback failed", ex);
            }
            return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
        }

        public void Dispose()
        {
            if (_hook == IntPtr.Zero) return;
            try
            {
                NativeMethods.UnhookWindowsHookEx(_hook);
                LoggingService.Info("Keyboard hook removed");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Keyboard hook remove failed", ex);
            }
            finally
            {
                _hook = IntPtr.Zero;
                _proc = null;
            }
        }
    }
}
