using System;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// Instant strip re-layout when Word windows move or are resized.
    ///
    /// Uses SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE) with WINEVENT_OUTOFCONTEXT
    /// and a filter on our own process id:
    ///  - this is NOT a global low-level hook and NOT DLL injection into other
    ///    processes (out-of-context events are delivered asynchronously to our
    ///    thread through its message queue);
    ///  - the alternatives are worse: subclassing Word windows (risky, banned in v1)
    ///    or polling the geometry aggressively on a timer (a stop condition of the experiment).
    ///
    /// The callback has to be cheap: filtering down to a few HWND comparisons, then
    /// an idempotent UpdateLayout (SetWindowPos only when something actually changed).
    /// The delegate lives in a field - the GC must not collect it while the hook is set.
    /// </summary>
    internal sealed class WinEventHookService : IDisposable
    {
        private readonly Action<IntPtr> _onWindowLocationChanged;
        private NativeWin32.WinEventDelegate _proc;   // GC root
        private IntPtr _hook = IntPtr.Zero;

        public WinEventHookService(Action<IntPtr> onWindowLocationChanged)
        {
            _onWindowLocationChanged = onWindowLocationChanged;
        }

        /// <summary>Installs the hook. Call from Word UI thread (it pumps the messages).</summary>
        public void Install()
        {
            if (_hook != IntPtr.Zero) return;
            try
            {
                _proc = Callback;
                _hook = NativeWin32.SetWinEventHook(
                    NativeWin32.EVENT_OBJECT_LOCATIONCHANGE,
                    NativeWin32.EVENT_OBJECT_LOCATIONCHANGE,
                    IntPtr.Zero, _proc,
                    NativeWin32.GetCurrentProcessId(), 0,
                    NativeWin32.WINEVENT_OUTOFCONTEXT);

                LoggingService.Info(_hook != IntPtr.Zero
                    ? "WinEvent hook installed (location change, own process)"
                    : "WinEvent hook install failed (SetWinEventHook returned NULL)");
            }
            catch (Exception ex)
            {
                LoggingService.Error("WinEvent hook install failed", ex);
                _hook = IntPtr.Zero;
                _proc = null;
            }
        }

        private void Callback(IntPtr hook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint thread, uint time)
        {
            try
            {
                // Only windows are of interest (OBJID_WINDOW); the caret, the cursor and
                // other objects generate the bulk of the events - filter them out.
                if (idObject != NativeWin32.OBJID_WINDOW || hwnd == IntPtr.Zero) return;

                var handler = _onWindowLocationChanged;
                if (handler != null) handler(hwnd);
            }
            catch (Exception ex)
            {
                // An exception from a hook callback must never reach Word.
                LoggingService.Error("WinEvent callback failed", ex);
            }
        }

        public void Dispose()
        {
            if (_hook == IntPtr.Zero)
            {
                _proc = null;
                return;
            }
            try
            {
                NativeWin32.UnhookWinEvent(_hook);
                LoggingService.Info("WinEvent hook removed");
            }
            catch (Exception ex)
            {
                LoggingService.Error("WinEvent unhook failed", ex);
            }
            finally
            {
                _hook = IntPtr.Zero;
                _proc = null;
            }
        }
    }
}
