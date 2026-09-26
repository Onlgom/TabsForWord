using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// The in-window strip manager: a dictionary of "Word window HWND -> NativeTabHost".
    ///
    /// It implements the same ITabPaneSync contract as TabPaneManager, so
    /// DocumentWindowManager (the tab model) knows nothing about this mode at all.
    /// It keeps an ordinary TabPaneManager inside as its fallback: on a critical
    /// failure every host is destroyed and the very next synchronisation creates
    /// ordinary Custom Task Panes - Word carries on working.
    ///
    /// Every operation is idempotent: EnsureHostForWindow never creates duplicates,
    /// PruneDeadHosts survives already-destroyed windows, a repeated Dispose is safe.
    /// </summary>
    public sealed class NativeTabHostManager : ITabPaneSync, IDisposable
    {
        private const int MaxConsecutiveFailures = 5;

        private readonly TabPaneManager _fallback;
        private readonly Dictionary<int, NativeTabHost> _hosts = new Dictionary<int, NativeTabHost>();
        private readonly HashSet<long> _dumpedTrees = new HashSet<long>();

        private DocumentWindowManager _manager;
        private WinEventHookService _winEvents;
        private List<DocumentTabModel> _lastTabs = new List<DocumentTabModel>();
        private bool _fallbackActive;
        private bool _collapsed;
        private int _consecutiveFailures;
        private bool _disposed;

        public NativeTabHostManager(TabPaneManager fallback)
        {
            _fallback = fallback;
        }

        /// <summary>Whether the strip is collapsed (shared by every window, as in CTP mode).</summary>
        public bool Collapsed { get { return _collapsed; } }

        /// <summary>Whether the emergency fallback to Custom Task Panes is active.</summary>
        public bool FallbackActive { get { return _fallbackActive; } }

        public void Init(DocumentWindowManager manager)
        {
            _manager = manager;
            // After every switch (click, Enter, the all-tabs menu, Ctrl+Tab) the focus goes
            // back to the document of the activated window if our strip still holds it.
            _manager.SetAfterActivate(RestoreFocusToDocument);
            _winEvents = new WinEventHookService(OnWindowLocationChanged);
            _winEvents.Install();
            LoggingService.Info("Native host mode enabled");
        }

        // ------------------------------------------------------------------
        // ITabPaneSync
        // ------------------------------------------------------------------

        public void SyncPanes(List<Word.Window> liveWindows, List<DocumentTabModel> tabs, bool tabsChanged)
        {
            if (_disposed) return;
            if (_fallbackActive)
            {
                _fallback.SyncPanes(liveWindows, tabs, tabsChanged);
                return;
            }

            _lastTabs = tabs;
            try
            {
                PruneDeadHosts(liveWindows);

                int attempts = 0, failures = 0;
                foreach (var w in liveWindows)
                {
                    int hwnd;
                    try { hwnd = w.Hwnd; }
                    catch (COMException) { continue; }

                    attempts++;
                    var host = EnsureHostForWindow(hwnd);
                    if (host == null) { failures++; continue; }

                    host.PushTabs(tabs);
                    // A layout failure (document zone not found and so on) counts as a
                    // native-mode failure too: without this the main failure scenario
                    // (Office updated, window classes changed) would never reach the fallback.
                    if (!host.UpdateLayout()) failures++;
                }

                // A complete failure across every window in a row (creation, zone or
                // layout) - switch over to CTP.
                if (attempts > 0 && failures == attempts) _consecutiveFailures++;
                else _consecutiveFailures = 0;

                if (_consecutiveFailures >= MaxConsecutiveFailures)
                    ActivateFallback("native host unusable for every window " +
                        _consecutiveFailures + " syncs in a row (create/zone/layout)");
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHostManager.SyncPanes failed", ex);
                if (++_consecutiveFailures >= MaxConsecutiveFailures)
                    ActivateFallback("repeated SyncPanes failures: " + ex.Message);
            }
        }

        public void PushTabs(List<DocumentTabModel> tabs)
        {
            if (_disposed) return;
            if (_fallbackActive)
            {
                _fallback.PushTabs(tabs);
                return;
            }
            _lastTabs = tabs;
            foreach (var host in SnapshotHosts())
            {
                try { host.PushTabs(tabs); }
                catch (Exception ex) { LoggingService.Error("PushTabs to native host failed", ex); }
            }
        }

        // ------------------------------------------------------------------
        // Host lifecycle
        // ------------------------------------------------------------------

        /// <summary>Idempotent: a live host of the window is reused, duplicates are impossible.</summary>
        internal NativeTabHost EnsureHostForWindow(int hwnd)
        {
            NativeTabHost existing;
            if (_hosts.TryGetValue(hwnd, out existing)) return existing;

            var raw = new IntPtr(hwnd);
            if (!NativeWin32.IsWindow(raw)) return null; // the window died between the snapshot and the sync

            // Window.Hwnd in SDI Word is the top-level window, but normalise it just in case.
            var root = NativeWin32.GetAncestor(raw, NativeWin32.GA_ROOT);
            if (root == IntPtr.Zero) root = raw;

            LoggingService.Info("Word HWND discovered: om=0x" + raw.ToInt64().ToString("X") +
                ", root=0x" + root.ToInt64().ToString("X") +
                " \"" + NativeWin32.GetClassNameSafe(root) + "\"");

            if (TabHostSettings.DumpWindowTree && _dumpedTrees.Add(root.ToInt64()))
                WindowTreeDiagnostics.DumpToFile(root, "host creation (dump=1)");

            var host = new NativeTabHost(this, root);
            if (!host.EnsureCreated())
            {
                host.Dispose();
                return null;
            }
            _hosts[hwnd] = host;
            return host;
        }

        /// <summary>Idempotent: a missing host is not an error.</summary>
        internal void RemoveHostForWindow(int hwnd)
        {
            NativeTabHost host;
            if (!_hosts.TryGetValue(hwnd, out host)) return;
            _hosts.Remove(hwnd);
            try { host.Dispose(); }
            catch (Exception ex) { LoggingService.Error("Host dispose failed", ex); }
        }

        private void PruneDeadHosts(List<Word.Window> liveWindows)
        {
            var liveKeys = new HashSet<int>();
            foreach (var w in liveWindows)
            {
                try { liveKeys.Add(w.Hwnd); }
                catch (COMException) { }
            }

            List<int> dead = null;
            foreach (var kv in _hosts)
            {
                bool alive = liveKeys.Contains(kv.Key) && NativeWin32.IsWindow(kv.Value.WordWindowHandle);
                if (!alive)
                {
                    if (dead == null) dead = new List<int>();
                    dead.Add(kv.Key);
                }
            }
            if (dead == null) return;

            foreach (var key in dead)
            {
                LoggingService.Info("Word HWND no longer valid: hwnd=" + key);
                RemoveHostForWindow(key);
            }
        }

        private List<NativeTabHost> SnapshotHosts()
        {
            var list = new List<NativeTabHost>(_hosts.Count);
            foreach (var kv in _hosts) list.Add(kv.Value);
            return list;
        }

        // ------------------------------------------------------------------
        // Instant reaction to windows moving (WinEvent)
        // ------------------------------------------------------------------

        private void OnWindowLocationChanged(IntPtr hwnd)
        {
            if (_disposed || _fallbackActive) return;

            // The callback fires on every move of every window of the process -
            // the filtering has to be cheap (a few comparisons per host).
            foreach (var kv in _hosts)
            {
                var host = kv.Value;
                if (hwnd == host.HostHandle) continue; // our own window - ignore (echo guard)
                if (hwnd == host.WordWindowHandle || hwnd == host.AnchorHwnd || hwnd == host.CanvasHwnd)
                {
                    host.UpdateLayout();
                    // Two hosts could share an HWND only in theory - stop searching.
                    return;
                }
            }
        }

        // ------------------------------------------------------------------
        // Emergency fallback
        // ------------------------------------------------------------------

        private void ActivateFallback(string reason)
        {
            if (_fallbackActive) return;
            _fallbackActive = true;
            LoggingService.Error("Native host failed; falling back to CustomTaskPane: " + reason);
            try
            {
                if (_winEvents != null) { _winEvents.Dispose(); _winEvents = null; }
                foreach (var host in SnapshotHosts())
                {
                    try { host.Dispose(); } catch { }
                }
                _hosts.Clear();
            }
            catch (Exception ex)
            {
                LoggingService.Error("Fallback cleanup failed", ex);
            }
            // The CTP strips are created by the next Reconcile (events + the ~1 Hz poll).
        }

        // ------------------------------------------------------------------
        // User actions coming from TabStripControl (the same ones as in TabPaneManager)
        // ------------------------------------------------------------------

        internal void HandleTabActivate(int hwnd)
        {
            // Posted: the switch runs once the click is fully over (RequestActivate), and
            // the focus is returned to the document right after it (SetAfterActivate in Init).
            if (_manager != null) _manager.RequestActivate(hwnd, "tab");
        }

        internal void HandleTabClose(int hwnd)
        {
            if (_manager != null) _manager.RequestCloseWindow(hwnd);
        }

        internal void HandleNewDocument()
        {
            if (_manager != null) _manager.CreateNewDocument();
        }

        internal void HandleCloseOthers(int hwnd)
        {
            if (_manager != null) _manager.CloseOtherWindows(hwnd);
        }

        internal void HandleOpenFolder(int hwnd)
        {
            if (_manager != null) _manager.OpenContainingFolder(hwnd);
        }

        internal void HandleReorder(int hwnd, int newIndex)
        {
            if (_manager != null) _manager.ReorderTab(hwnd, newIndex);
            RestoreFocusToActiveDocument();
        }

        internal void HandleTabColorChange(int hwnd, System.Drawing.Color? color)
        {
            if (_manager != null) _manager.SetTabColor(hwnd, color);
            RestoreFocusToActiveDocument();
        }

        internal void HandleTabPinChange(int hwnd, bool pinned)
        {
            if (_manager != null) _manager.SetTabPinned(hwnd, pinned);
            RestoreFocusToActiveDocument();
        }

        internal void HandleHotkeyToggle(bool enabled)
        {
            if (_manager != null) _manager.SetHotkeyEnabled(enabled);
            RestoreFocusToActiveDocument();
        }

        /// <summary>Interface language: save it and rebuild the captions in every window.</summary>
        internal void HandleLanguageChange(UiLang? lang)
        {
            UiLanguage.Save(lang);
            foreach (var host in SnapshotHosts())
            {
                try { host.RefreshLocalizedUi(); }
                catch (Exception ex) { LoggingService.Error("Apply language to host failed", ex); }
            }
            RestoreFocusToActiveDocument();
        }

        /// <summary>Tab size: save it and apply it to every window at once.</summary>
        internal void HandleTabSizeChange(float scale)
        {
            TabSizeSettings.Save(scale);
            foreach (var host in SnapshotHosts())
            {
                try { host.SetUserScale(scale); }
                catch (Exception ex) { LoggingService.Error("Apply tab size to host failed", ex); }
            }
            RestoreFocusToActiveDocument();
        }

        internal void HandleCollapseToggle()
        {
            _collapsed = !_collapsed;
            LoggingService.Info(_collapsed ? "Tab panel collapsed (native)" : "Tab panel expanded (native)");
            foreach (var host in SnapshotHosts())
            {
                host.SetCollapsed(_collapsed);
                host.UpdateLayout();
            }
            RestoreFocusToActiveDocument();
        }

        /// <summary>
        /// Returns focus to the document of the CURRENTLY active Word window. Needed
        /// after strip actions that do not switch tabs: collapse/expand, changing the
        /// size, the colour or the pinning, dragging a tab. Clicking one of our buttons
        /// takes the input focus onto the strip, and the next character typed (or
        /// Ctrl+F1) went nowhere - measured on a live Word: after clicking collapse the
        /// character never reached the document.
        /// </summary>
        private void RestoreFocusToActiveDocument()
        {
            try
            {
                var fg = NativeWin32.GetForegroundWindow();
                if (fg == IntPtr.Zero) return;
                foreach (var kv in _hosts)
                {
                    if (kv.Value != null && kv.Value.WordWindowHandle == fg)
                    {
                        RestoreFocusToDocument(kv.Key);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("RestoreFocusToActiveDocument failed", ex);
            }
        }

        /// <summary>
        /// A click on a tab may have left the input focus on our strip - return it to
        /// the document canvas of the activated window (same thread, so SetFocus is
        /// allowed). Does nothing when the focus is not ours anyway.
        /// </summary>
        private void RestoreFocusToDocument(int activatedHwnd)
        {
            try
            {
                var focus = NativeWin32.GetFocus();
                if (focus == IntPtr.Zero) return;

                bool ours = false;
                foreach (var kv in _hosts)
                {
                    if (kv.Value.OwnsFocus(focus)) { ours = true; break; }
                }
                if (!ours) return;

                NativeTabHost target;
                if (!_hosts.TryGetValue(activatedHwnd, out target)) return;

                // SetFocus into a background window would create a "focus behind the back"
                // of the active one - move the focus only if the target is really in front.
                if (NativeWin32.GetForegroundWindow() != target.WordWindowHandle) return;

                var canvas = target.CanvasHwnd;
                if (canvas != IntPtr.Zero && NativeWin32.IsWindow(canvas))
                    NativeWin32.SetFocus(canvas);
            }
            catch (Exception ex)
            {
                LoggingService.Error("RestoreFocusToDocument failed", ex);
            }
        }

        // ------------------------------------------------------------------

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                if (_winEvents != null) { _winEvents.Dispose(); _winEvents = null; }
                foreach (var host in SnapshotHosts())
                {
                    try { host.Dispose(); } catch { }
                }
                _hosts.Clear();
                LoggingService.Info("All native tab hosts disposed");
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHostManager dispose failed", ex);
            }
            // _fallback belongs to Connect, which disposes it.
        }
    }
}
