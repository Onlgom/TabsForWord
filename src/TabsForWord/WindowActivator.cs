using System;
using System.Windows.Forms;
using TabsForWord.NativeHost;

namespace TabsForWord
{
    /// <summary>
    /// The Win32 half of a tab switch: un-minimise, raise, and a short check afterwards
    /// that the target really stayed in front.
    ///
    /// Why the check exists: in the field (2026-09-17 log) the same tab was clicked
    /// three times in seven seconds - the switch was being undone and the old log line
    /// "Window activated" was written unconditionally, so nothing showed it. The check
    /// logs who took the foreground back and re-raises the target ONCE, and only when
    /// it is safe: the foreground went back to the very Word window the switch started
    /// from and the user has not touched the mouse or keyboard since the click (so a
    /// deliberate click back into the old window, Alt+Tab or the taskbar is never
    /// fought, and neither is anything else that brings some other window forward).
    ///
    /// Window titles are never read: they carry document names.
    /// </summary>
    internal static class WindowActivator
    {
        /// <summary>
        /// SC_RESTORE is what the taskbar button sends: a window that was maximised
        /// before it was minimised comes back maximised. (The COM route used before,
        /// WindowState = wdWindowStateNormal, brought such a window back as an
        /// ordinary one - the "some maximised, some not" mix the user saw.)
        /// SetForegroundWindow on its own never un-minimises.
        /// </summary>
        internal static void RestoreIfMinimized(IntPtr hwnd)
        {
            if (!NativeWin32.IsIconic(hwnd)) return;
            NativeWin32.SendMessage(hwnd, NativeWin32.WM_SYSCOMMAND, (IntPtr)NativeWin32.SC_RESTORE, IntPtr.Zero);
        }

        /// <summary>SetForegroundWindow + BringWindowToTop; returns what SetForegroundWindow said.</summary>
        internal static bool Raise(IntPtr hwnd)
        {
            bool ok = NativeMethods.SetForegroundWindow(hwnd);
            NativeMethods.BringWindowToTop(hwnd);
            return ok;
        }

        /// <summary>A Word document frame (OpusApp) of THIS process.</summary>
        internal static bool IsOurWordFrame(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !NativeWin32.IsWindow(hwnd)) return false;
            uint pid;
            NativeWin32.GetWindowThreadProcessId(hwnd, out pid);
            return pid == NativeWin32.GetCurrentProcessId()
                && NativeWin32.GetClassNameSafe(hwnd) == "OpusApp";
        }

        /// <summary>Hwnd, class and state for the log - never the title.</summary>
        internal static string Describe(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero) return "none";
                string hex = "0x" + hwnd.ToInt64().ToString("X");
                if (!NativeWin32.IsWindow(hwnd)) return hex + " dead";
                uint pid;
                NativeWin32.GetWindowThreadProcessId(hwnd, out pid);
                if (pid != NativeWin32.GetCurrentProcessId()) return hex + " other-process";
                string state = NativeWin32.IsIconic(hwnd) ? "min" : (NativeWin32.IsZoomed(hwnd) ? "max" : "normal");
                return hex + " " + NativeWin32.GetClassNameSafe(hwnd) + " " + state
                    + (NativeWin32.IsWindowEnabled(hwnd) ? "" : " disabled");
            }
            catch
            {
                return "?";
            }
        }
    }

    /// <summary>Pure decisions of the post-activation check (unit tested).</summary>
    internal static class ActivationPolicy
    {
        /// <summary>First look after the switch, and the second (last) one.</summary>
        internal const int FirstCheckMs = 250;
        internal const int LastCheckMs = 1200;

        /// <summary>
        /// Re-raise the target? Only when the foreground went back to the Word window the
        /// switch started from (the switch was undone - never another program, a dialog or
        /// some third window), the user has not done anything since the click, and this
        /// switch has not been retried yet.
        /// </summary>
        internal static bool ShouldRetry(bool fgIsSourceWordWindow, bool userInputSinceSwitch, bool alreadyRetried)
        {
            return fgIsSourceWordWindow && !userInputSinceSwitch && !alreadyRetried;
        }

        /// <summary>
        /// Was there input after the switch started? Both ticks come from the same
        /// wrapping millisecond counter; unknown (null) counts as "yes" - be conservative.
        /// </summary>
        internal static bool InputSince(int? lastInputTick, int switchTick)
        {
            if (!lastInputTick.HasValue) return true;
            return unchecked(lastInputTick.Value - switchTick) > 0;
        }
    }

    /// <summary>
    /// Watches one switch: at +250 ms and +1200 ms checks that the target is still in
    /// front. Lives on Word UI thread (WinForms timer). A new switch replaces the old watch.
    /// </summary>
    internal sealed class ActivationWatch : IDisposable
    {
        private readonly Timer _timer = new Timer();
        private IntPtr _target;
        private IntPtr _source;
        private int _requestTick;   // when the user asked for the switch: input after it is "new"
        private int _startTick;     // when the switch was done: for the elapsed time in the log
        private bool _retried;
        private bool _firstDone;

        public ActivationWatch()
        {
            _timer.Tick += OnTick;
        }

        /// <param name="target">the window that must stay in front</param>
        /// <param name="source">the window that was in front when the switch started</param>
        /// <param name="requestTick">Environment.TickCount of the click / key press that asked for it</param>
        public void Start(IntPtr target, IntPtr source, int requestTick)
        {
            _target = target;
            _source = source;
            _requestTick = requestTick;
            _startTick = Environment.TickCount;
            _retried = false;
            _firstDone = false;
            _timer.Stop();
            _timer.Interval = ActivationPolicy.FirstCheckMs;
            _timer.Start();
        }

        public void Cancel()
        {
            _timer.Stop();
            _target = IntPtr.Zero;
        }

        private void OnTick(object sender, EventArgs e)
        {
            try
            {
                _timer.Stop();
                var target = _target;
                if (target == IntPtr.Zero || !NativeWin32.IsWindow(target)) { Cancel(); return; }

                var fg = NativeWin32.GetForegroundWindow();
                // The target in front, or one of its own popups (a Word dialog owned by that window).
                bool ok = fg == target || (fg != IntPtr.Zero && NativeWin32.GetAncestor(fg, NativeWin32.GA_ROOTOWNER) == target);
                int elapsed = unchecked(Environment.TickCount - _startTick);

                if (!ok)
                {
                    bool input = ActivationPolicy.InputSince(NativeWin32.GetLastInputTickSafe(), _requestTick);
                    bool ours = WindowActivator.IsOurWordFrame(fg);
                    bool backToSource = ours && fg == _source;
                    LoggingService.Warn("Activate check: target " + WindowActivator.Describe(target) +
                        " is not in front at +" + elapsed + " ms; foreground " + WindowActivator.Describe(fg) +
                        (backToSource ? " (the window the switch started from)" : (ours ? "" : " (not a Word document window)")) +
                        (input ? ", user input since the click" : ", no input since the click"));

                    if (ActivationPolicy.ShouldRetry(backToSource, input, _retried))
                    {
                        _retried = true;
                        WindowActivator.RestoreIfMinimized(target);
                        bool sfw = WindowActivator.Raise(target);
                        LoggingService.Info("Activate retry: sfw=" + (sfw ? 1 : 0) + ", foreground now " +
                            WindowActivator.Describe(NativeWin32.GetForegroundWindow()));
                    }
                }
                else if (_retried && _firstDone)
                {
                    LoggingService.Info("Activate check: target stayed in front after the retry");
                }

                if (!_firstDone)
                {
                    _firstDone = true;
                    _timer.Interval = ActivationPolicy.LastCheckMs - ActivationPolicy.FirstCheckMs;
                    _timer.Start();
                    return;
                }
                Cancel();
            }
            catch (Exception ex)
            {
                Cancel();
                LoggingService.Error("Activate check failed", ex);
            }
        }

        public void Dispose()
        {
            try { _timer.Stop(); _timer.Dispose(); } catch { }
        }
    }
}
