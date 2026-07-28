using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord
{
    /// <summary>
    /// Manages the Custom Task Panes: in SDI Word every pane belongs to exactly one
    /// document window, so there is one pane per Word.Window (keyed by Hwnd).
    ///
    /// Word destroys a pane itself when its window closes; no event is raised for that,
    /// and our RCW simply goes "dead" (a COMException on any access). That is why every
    /// synchronisation starts by probing for and cleaning up dead panes.
    /// </summary>
    public sealed class TabPaneManager : IDisposable, ITabPaneSync
    {
        private const string ControlProgId = "TabsForWord.TabStripControl";

        // Per specification 2b the "Documents" title line is removed. Word own CTP title
        // strip cannot be removed through the official API, so it is at least left
        // without any text (see docs/KNOWN_ISSUES.md, O2).
        private const string PaneTitle = " ";

        // Word enforces a minimum CTP height, and its own title eats part of that height.
        // We start from a sensible value and then make up the shortfall from the actual
        // control height (DesiredContentHeight).
        private const int InitialHeightPx = 72;

        private sealed class PaneEntry
        {
            public Office.CustomTaskPane Ctp;
            public TabStripControl Control;
            public int LastDpi;   // DPI of the Word window at the last sync (0 = not set yet)
        }

        private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();
        private Office.ICTPFactory _factory;
        private DocumentWindowManager _manager;
        private bool _heightLogged;
        private bool _collapsed;   // collapse state, shared by every strip (per session)
        private bool _disposed;

        public void Init(Office.ICTPFactory factory, DocumentWindowManager manager)
        {
            _factory = factory;
            _manager = manager;
        }

        /// <summary>
        /// Brings the strips in step with the live windows: drop the dead ones, create the
        /// missing ones, hand the current tab list to every control.
        /// </summary>
        public void SyncPanes(List<Word.Window> liveWindows, List<DocumentTabModel> tabs, bool tabsChanged)
        {
            if (_disposed || _factory == null) return;

            PruneDeadPanes();

            foreach (var w in liveWindows)
            {
                int hwnd;
                try { hwnd = w.Hwnd; }
                catch (COMException) { continue; }

                if (!_panes.ContainsKey(hwnd))
                    CreatePane(w, hwnd);

                // The window may have moved to a monitor with a different scale - the strip
                // follows the DPI of the Word window (SyncPanes runs at about 1 Hz).
                PaneEntry entry;
                if (_panes.TryGetValue(hwnd, out entry))
                    SyncPaneDpi(entry, hwnd);
            }

            // Handing out the snapshot: UpdateTabs does not repaint when nothing changed,
            // so we always hand it out (new strips then get the list immediately).
            foreach (var entry in _panes.Values.ToList())
            {
                try
                {
                    entry.Control.UpdateTabs(tabs);
                }
                catch (Exception ex)
                {
                    LoggingService.Error("UpdateTabs push failed", ex);
                }
            }
        }

        private void CreatePane(Word.Window window, int hwnd)
        {
            try
            {
                var ctp = _factory.CreateCTP(ControlProgId, PaneTitle, window);
                ctp.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionTop;
                ctp.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                try
                {
                    ctp.Height = InitialHeightPx;
                }
                catch (COMException)
                {
                    // Word may reject the height - accept its value and correct below.
                }

                var control = ctp.ContentControl as TabStripControl;
                if (control == null)
                {
                    LoggingService.Error("ContentControl is not TabStripControl (registration problem?)");
                    try { ctp.Delete(); } catch { }
                    return;
                }

                control.TabActivateRequested += OnTabActivateRequested;
                control.TabCloseRequested += OnTabCloseRequested;
                control.NewDocumentRequested += OnNewDocumentRequested;
                control.CloseOthersRequested += OnCloseOthersRequested;
                control.OpenFolderRequested += OnOpenFolderRequested;
                control.ReorderRequested += OnReorderRequested;
                control.CollapseToggleRequested += OnCollapseToggleRequested;
                control.TabColorChangeRequested += OnTabColorChangeRequested;
                control.TabSizeChangeRequested += OnTabSizeChangeRequested;
                control.TabPinChangeRequested += OnTabPinChangeRequested;
                control.HotkeyToggleRequested += OnHotkeyToggleRequested;
                control.LanguageChangeRequested += OnLanguageChangeRequested;

                control.SetCollapsed(_collapsed);
                ctp.Visible = true;

                // The strip scale follows the DPI of the CONTROL own window (the space its
                // pixels live in; the CTP host and ctp.Height are in the same space, so no
                // conversion is needed). It is assigned BEFORE the first height fix-up,
                // which uses DesiredContentHeight.
                int paneDpi = control.IsHandleCreated
                    ? NativeHost.NativeWin32.GetDpiForWindowSafe(control.Handle, 0)
                    : 0;
                if (paneDpi > 0) control.SetDpiOverride(paneDpi);

                // Word own pane title eats part of the pane height. Right after Visible the
                // control is not measured yet (Height=0), so the height is corrected once,
                // when the host really lays the control out (the first SizeChanged with Height>0).
                AttachOneShotHeightFix(ctp, control);

                _panes[hwnd] = new PaneEntry { Ctp = ctp, Control = control, LastDpi = paneDpi };
                LoggingService.Info("Tab pane created: hwnd=" + hwnd);
            }
            catch (COMException ex)
            {
                // A failure on one pane must not bring the others down.
                LoggingService.Error("CreateCTP failed for hwnd=" + hwnd, ex);
            }
            catch (Exception ex)
            {
                LoggingService.Error("CreatePane failed for hwnd=" + hwnd, ex);
            }
        }

        private void AttachOneShotHeightFix(Office.CustomTaskPane ctp, TabStripControl control)
        {
            bool done = false;
            EventHandler handler = null;
            handler = delegate
            {
                if (done) return;
                try
                {
                    int contentH = control.Height;
                    if (contentH <= 0) return; // not laid out yet - wait for the next SizeChanged

                    done = true;
                    control.SizeChanged -= handler;

                    int diff = contentH - control.DesiredContentHeight;
                    if (diff != 0)
                    {
                        try { ctp.Height = ctp.Height - diff; }
                        catch (COMException) { /* Word rejected it (its minimum) - accept as is */ }
                    }

                    if (!_heightLogged)
                    {
                        _heightLogged = true;
                        try { LoggingService.Info("CTP height: pane=" + ctp.Height + ", content=" + control.Height); }
                        catch (COMException) { }
                    }
                }
                catch (Exception ex)
                {
                    done = true;
                    try { control.SizeChanged -= handler; } catch { }
                    LoggingService.Error("Height fix failed", ex);
                }
            };
            control.SizeChanged += handler;
        }

        /// <summary>
        /// The strip scale = DPI of the Word window (per-monitor). When the DPI changes
        /// (the window moved to another monitor) the control is given a new scale and the
        /// CTP height is corrected by the delta between the actual and the desired height.
        /// The correction runs only ON A DPI CHANGE - otherwise, whenever Word refuses
        /// (its CTP minimum), the fix-up would repeat every second.
        /// </summary>
        private void SyncPaneDpi(PaneEntry entry, int hwnd)
        {
            try
            {
                if (!entry.Control.IsHandleCreated) return;
                int dpi = NativeHost.NativeWin32.GetDpiForWindowSafe(entry.Control.Handle, 0);
                if (dpi <= 0 || dpi == entry.LastDpi) return; // no API (old Windows) or nothing changed

                bool first = entry.LastDpi == 0;
                entry.LastDpi = dpi;
                entry.Control.SetDpiOverride(dpi);

                // On creation the height is settled by AttachOneShotHeightFix - here we
                // only correct later monitor changes.
                if (first) return;
                int diff = entry.Control.Height - entry.Control.DesiredContentHeight;
                if (diff != 0)
                {
                    try { entry.Ctp.Height = entry.Ctp.Height - diff; }
                    catch (COMException) { /* Word rejected it (its CTP minimum) - accept as is */ }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("SyncPaneDpi failed", ex);
            }
        }

        /// <summary>Pushes the current tab order to every live strip (after a reorder).</summary>
        public void PushTabs(List<DocumentTabModel> tabs)
        {
            if (_disposed) return;
            foreach (var entry in _panes.Values.ToList())
            {
                try { entry.Control.UpdateTabs(tabs); }
                catch (Exception ex) { LoggingService.Error("PushTabs failed", ex); }
            }
        }

        private void OnTabActivateRequested(int hwnd)
        {
            if (_manager != null) _manager.ActivateWindow(hwnd);
        }

        private void OnTabCloseRequested(int hwnd)
        {
            if (_manager != null) _manager.RequestCloseWindow(hwnd);
        }

        private void OnNewDocumentRequested()
        {
            if (_manager != null) _manager.CreateNewDocument();
        }

        private void OnCloseOthersRequested(int hwnd)
        {
            if (_manager != null) _manager.CloseOtherWindows(hwnd);
        }

        private void OnOpenFolderRequested(int hwnd)
        {
            if (_manager != null) _manager.OpenContainingFolder(hwnd);
        }

        private void OnReorderRequested(int hwnd, int newIndex)
        {
            if (_manager != null) _manager.ReorderTab(hwnd, newIndex);
        }

        private void OnTabColorChangeRequested(int hwnd, System.Drawing.Color? color)
        {
            if (_manager != null) _manager.SetTabColor(hwnd, color);
        }

        private void OnTabPinChangeRequested(int hwnd, bool pinned)
        {
            if (_manager != null) _manager.SetTabPinned(hwnd, pinned);
        }

        private void OnHotkeyToggleRequested(bool enabled)
        {
            if (_manager != null) _manager.SetHotkeyEnabled(enabled);
        }

        /// <summary>
        /// Interface language: save the choice and rebuild the captions in every strip
        /// (null = back to auto-detection from Word own language).
        /// </summary>
        private void OnLanguageChangeRequested(UiLang? lang)
        {
            try
            {
                UiLanguage.Save(lang);
                foreach (var entry in _panes.Values.ToList())
                {
                    try { entry.Control.RefreshLocalizedUi(); }
                    catch (Exception ex) { LoggingService.Error("Apply language to pane failed", ex); }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnLanguageChangeRequested failed", ex);
            }
        }

        /// <summary>
        /// Tab size: save the multiplier and apply it to every strip at once (the CTP
        /// height is corrected by the delta, exactly as when collapsing).
        /// </summary>
        private void OnTabSizeChangeRequested(float scale)
        {
            try
            {
                TabSizeSettings.Save(scale);
                foreach (var entry in _panes.Values.ToList())
                {
                    try
                    {
                        entry.Control.SetUserScale(scale);
                        int diff = entry.Control.Height - entry.Control.DesiredContentHeight;
                        if (diff != 0)
                        {
                            try { entry.Ctp.Height = entry.Ctp.Height - diff; }
                            catch (COMException) { /* Word rejected it (its CTP minimum) - accept as is */ }
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggingService.Error("Apply tab size to pane failed", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnTabSizeChangeRequested failed", ex);
            }
        }

        /// <summary>
        /// Collapse/expand the strip in EVERY window. The CTP height changes by the delta
        /// of the actual control height; if Word refuses (its CTP minimum), the control
        /// still draws the collapsed sliver.
        /// </summary>
        private void OnCollapseToggleRequested()
        {
            _collapsed = !_collapsed;
            LoggingService.Info(_collapsed ? "Tab panel collapsed" : "Tab panel expanded");
            foreach (var entry in _panes.Values.ToList())
                ApplyCollapsed(entry);
        }

        private void ApplyCollapsed(PaneEntry entry)
        {
            try
            {
                entry.Control.SetCollapsed(_collapsed);
                int diff = entry.Control.Height - entry.Control.DesiredContentHeight;
                if (diff != 0)
                {
                    try { entry.Ctp.Height = entry.Ctp.Height - diff; }
                    catch (COMException) { /* Word rejected it (its CTP minimum) - accept as is */ }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("ApplyCollapsed failed", ex);
            }
        }

        private void PruneDeadPanes()
        {
            List<int> dead = null;
            foreach (var kv in _panes)
            {
                bool alive;
                try
                {
                    // Any access to a dead pane throws a COMException.
                    alive = kv.Value.Ctp.Visible || true;
                }
                catch (COMException)
                {
                    alive = false;
                }
                catch (InvalidComObjectException)
                {
                    alive = false;
                }
                if (!alive)
                {
                    if (dead == null) dead = new List<int>();
                    dead.Add(kv.Key);
                }
            }
            if (dead == null) return;

            foreach (var hwnd in dead)
            {
                var entry = _panes[hwnd];
                _panes.Remove(hwnd);
                try { Marshal.ReleaseComObject(entry.Ctp); } catch { }
                LoggingService.Info("Dead tab pane removed: hwnd=" + hwnd);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var entry in _panes.Values)
            {
                // By unload time Office may already have destroyed the panes - every risky
                // operation gets its own try/catch.
                try { entry.Ctp.Delete(); } catch { }
                try { Marshal.ReleaseComObject(entry.Ctp); } catch { }
            }
            _panes.Clear();
            LoggingService.Info("All tab panes disposed");
        }
    }
}
