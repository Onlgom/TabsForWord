using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord
{
    /// <summary>
    /// Owner of the tab model. The only way it is updated is the idempotent
    /// Reconcile(): a diff of the live Word collections (Windows + ProtectedViewWindows)
    /// against the current model (see ADR-004).
    ///
    /// Runs strictly on Word main STA thread: Word events, WinForms timers and clicks
    /// on the strip all arrive on the same thread, so no locking is needed.
    /// </summary>
    public sealed class DocumentWindowManager : IDisposable
    {
        // Word error code for "command cancelled by the user" (cancel in the save dialog).
        private const int WdErrCancelledByUser = 4198;

        private readonly Word.Application _word;
        private readonly ITabPaneSync _paneManager;
        private readonly TabColorStore _colors = new TabColorStore();
        private readonly TabOrderStore _order = new TabOrderStore();

        private List<DocumentTabModel> _tabs = new List<DocumentTabModel>();
        private Timer _debounceTimer;   // one-shot tick after a "hint" (BeforeClose and friends)
        private Timer _pollTimer;       // background reconcile: Save As, the Saved flag, missed events
        private KeyboardHookService _keyboardHook;  // Ctrl+Tab / Ctrl+Shift+Tab; null = switched off in settings
        private bool _startupOrderApplied;          // the stored order is restored once, at startup
        private bool _disposed;

        public DocumentWindowManager(Word.Application word, ITabPaneSync paneManager)
        {
            _word = word;
            _paneManager = paneManager;
        }

        public void Start()
        {
            _debounceTimer = new Timer { Interval = 300 };
            _debounceTimer.Tick += OnDebounceTick;

            _pollTimer = new Timer { Interval = 1000 };
            _pollTimer.Tick += OnPollTick;
            _pollTimer.Start();

            // Ctrl+Tab / Ctrl+Shift+Tab (specification 2b, section 8) - can be switched off
            // in the settings window (docs/KNOWN_ISSUES.md O8: it shadows the ordinary tab
            // key inside tables). Start() runs on Word UI thread - the hook goes onto it.
            if (HotkeySettings.CtrlTabEnabled) InstallHotkey();

            Reconcile("startup");
        }

        private void InstallHotkey()
        {
            if (_keyboardHook != null) return;
            _keyboardHook = new KeyboardHookService(() => ActivateAdjacent(1), () => ActivateAdjacent(-1));
            _keyboardHook.Install();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                if (_debounceTimer != null) { _debounceTimer.Stop(); _debounceTimer.Dispose(); _debounceTimer = null; }
                if (_pollTimer != null) { _pollTimer.Stop(); _pollTimer.Dispose(); _pollTimer = null; }
                if (_keyboardHook != null) { _keyboardHook.Dispose(); _keyboardHook = null; }
            }
            catch (Exception ex)
            {
                LoggingService.Error("DocumentWindowManager dispose failed", ex);
            }
        }

        /// <summary>Immediate reconcile (called from Word events).</summary>
        public void RequestReconcile(string reason)
        {
            if (_disposed) return;
            Reconcile(reason);
        }

        /// <summary>
        /// Deferred reconcile (about 300 ms): for "hint" events after which Word still
        /// shows a dialog (DocumentBeforeClose, DocumentBeforeSave) - the outcome only
        /// becomes known once the dialog is done.
        /// </summary>
        public void RequestDeferredReconcile(string reason)
        {
            if (_disposed || _debounceTimer == null) return;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void OnDebounceTick(object sender, EventArgs e)
        {
            try
            {
                _debounceTimer.Stop();
                Reconcile("deferred");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Deferred reconcile failed", ex);
            }
        }

        private void OnPollTick(object sender, EventArgs e)
        {
            try
            {
                Reconcile(null); // null = do not log a reason (background tick)
            }
            catch (Exception ex)
            {
                LoggingService.Error("Poll reconcile failed", ex);
            }
        }

        // ------------------------------------------------------------------
        // Reconcile
        // ------------------------------------------------------------------

        private void Reconcile(string reason)
        {
            if (_disposed) return;

            List<DocumentTabModel> snapshot;
            List<Word.Window> liveWindows;
            try
            {
                snapshot = BuildSnapshot(out liveWindows);
            }
            catch (Exception ex)
            {
                // Word may be in a transient state (closing, a modal dialog).
                LoggingService.Error("BuildSnapshot failed" + (reason != null ? " (" + reason + ")" : ""), ex);
                return;
            }

            // Tab order is stable: existing tabs keep their position, new ones go to the end.
            var ordered = new List<DocumentTabModel>();
            foreach (var old in _tabs)
            {
                var live = snapshot.FirstOrDefault(t => t.Hwnd == old.Hwnd);
                if (live != null) ordered.Add(live);
            }
            foreach (var t in snapshot)
            {
                if (ordered.All(x => x.Hwnd != t.Hwnd)) ordered.Add(t);
            }

            // First reconcile after Word starts: restore the stored order (otherwise the
            // initial order is whatever Word.Windows enumeration returns, which is
            // arbitrary). After that the order lives in memory as usual - existing tabs
            // keep their position and new ones go to the end.
            if (!_startupOrderApplied)
            {
                _startupOrderApplied = true;
                ordered = ordered.OrderBy(t => _order.RankOf(t.FullPath) ?? int.MaxValue).ToList();
            }

            // Pinned tabs always come first (OrderBy is stable: the relative order inside
            // each group is preserved). The same invariant lives in ReorderTab and in the
            // drag preview (TabStripControl).
            ordered = ordered.OrderBy(t => t.IsPinned ? 0 : 1).ToList();

            bool changed = !DocumentTabModel.ListsEqual(_tabs, ordered);
            if (changed)
            {
                foreach (var t in ordered.Where(t => _tabs.All(x => x.Hwnd != t.Hwnd)))
                    LoggingService.Info("Tab added: hwnd=" + t.Hwnd + ", name=" + t.Caption + (t.IsProtectedView ? " [PV]" : ""));
                foreach (var t in _tabs.Where(t => ordered.All(x => x.Hwnd != t.Hwnd)))
                {
                    LoggingService.Info("Tab removed: hwnd=" + t.Hwnd + ", name=" + t.Caption);
                    _colors.ForgetWindow(t.Hwnd); // session colour of an unsaved document
                    _order.ForgetWindow(t.Hwnd);  // session pinning of an unsaved document
                }
                _tabs = ordered;
                _order.SaveOrder(_tabs);
            }

            // The strips are maintained always (a window can appear without the model changing, and vice versa).
            _paneManager.SyncPanes(liveWindows, _tabs, changed);
        }

        private List<DocumentTabModel> BuildSnapshot(out List<Word.Window> liveWindows)
        {
            var result = new List<DocumentTabModel>();
            liveWindows = new List<Word.Window>();

            int activeHwnd = 0;
            try
            {
                if (_word.Windows.Count > 0) activeHwnd = _word.ActiveWindow.Hwnd;
            }
            catch (COMException) { /* no active window (only PV, or nothing at all) */ }

            // Ordinary document windows
            foreach (Word.Window w in _word.Windows)
            {
                try
                {
                    var doc = w.Document;
                    int hwnd = w.Hwnd;
                    string name = doc.Name;
                    string path = doc.Path;
                    string fullPath = string.IsNullOrEmpty(path) ? name : doc.FullName;
                    bool saved = doc.Saved;

                    string caption = name;
                    try
                    {
                        if (doc.Windows.Count > 1) caption = name + ":" + w.WindowNumber;
                    }
                    catch (COMException) { }

                    // First save / Save As: move the tab colour and pinning onto the new key
                    // BEFORE Resolve/IsPinned below - otherwise they would vanish after Ctrl+S.
                    var prev = _tabs.FirstOrDefault(x => x.Hwnd == hwnd);
                    if (prev != null)
                    {
                        _colors.HandlePathChanged(hwnd, prev.FullPath, fullPath);
                        _order.HandlePathChanged(hwnd, prev.FullPath, fullPath);
                    }

                    result.Add(new DocumentTabModel
                    {
                        Hwnd = hwnd,
                        Caption = caption,
                        FullPath = fullPath,
                        Saved = saved,
                        IsActive = hwnd == activeHwnd,
                        IsProtectedView = false,
                        TabColor = _colors.Resolve(fullPath, hwnd),
                        IsPinned = _order.IsPinned(fullPath, hwnd)
                    });
                    liveWindows.Add(w);
                }
                catch (COMException)
                {
                    // The window is in a transient state (closing) - skip it, the next
                    // Reconcile will see a settled state.
                }
            }

            // Protected View windows (a separate collection; they are not in Documents).
            // A ProtectedViewWindow has NO Hwnd property - we use a synthetic negative key
            // derived from the full source path (it can never collide with a real hwnd).
            try
            {
                string pvActiveKeySource = null;
                try
                {
                    if (_word.ProtectedViewWindows.Count > 0)
                    {
                        // During the PV -> editing transition this property can return null
                        // (not a COMException!) - the null guard is mandatory.
                        var apv = _word.ActiveProtectedViewWindow;
                        if (apv != null) pvActiveKeySource = PvSourceFullPath(apv);
                    }
                }
                catch (Exception) { }

                foreach (Word.ProtectedViewWindow pv in _word.ProtectedViewWindows)
                {
                    try
                    {
                        if (pv == null) continue;
                        string full = PvSourceFullPath(pv);
                        string src = pv.SourceName;
                        int pvKey = PvKey(full);
                        result.Add(new DocumentTabModel
                        {
                            Hwnd = pvKey,
                            Caption = src,
                            FullPath = full,
                            Saved = true,
                            IsActive = pvActiveKeySource != null &&
                                       string.Equals(full, pvActiveKeySource, StringComparison.OrdinalIgnoreCase),
                            IsProtectedView = true,
                            TabColor = _colors.Resolve(full, pvKey),
                            IsPinned = _order.IsPinned(full, pvKey)
                        });
                    }
                    catch (Exception) { /* the PV window is in a transient state - skip it */ }
                }
            }
            catch (Exception) { }

            return result;
        }

        private static string PvSourceFullPath(Word.ProtectedViewWindow pv)
        {
            string src = pv.SourceName;
            string srcPath = pv.SourcePath;
            return string.IsNullOrEmpty(srcPath) ? src : srcPath + "\\" + src;
        }

        /// <summary>Synthetic key of a PV tab: always negative, stable for a given path.</summary>
        internal static int PvKey(string fullPath)
        {
            int h = StringComparer.OrdinalIgnoreCase.GetHashCode(fullPath ?? string.Empty);
            if (h == int.MinValue) h = int.MinValue + 1;
            return -Math.Abs(h) - 1;
        }

        // ------------------------------------------------------------------
        // User actions (called from TabStripControl through TabPaneManager)
        // ------------------------------------------------------------------

        /// <summary>Click on a tab: activate the corresponding window.</summary>
        public void ActivateWindow(int hwnd)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab != null && tab.IsProtectedView)
                {
                    ActivateProtectedView(hwnd);
                    return;
                }

                foreach (Word.Window w in _word.Windows)
                {
                    bool match;
                    try { match = w.Hwnd == hwnd; }
                    catch (COMException) { continue; }
                    if (!match) continue;

                    // COM activation. Its failure must not cancel the Win32 raise below.
                    try
                    {
                        if (w.WindowState == Word.WdWindowState.wdWindowStateMinimize)
                            w.WindowState = Word.WdWindowState.wdWindowStateNormal;
                        w.Activate();
                    }
                    catch (COMException ex)
                    {
                        LoggingService.Warn("Activate (COM) failed for hwnd=" + hwnd + ": " + ex.Message);
                    }

                    // Win32 raise: SetForegroundWindow can fail silently during activation
                    // transitions (a mix of maximised and ordinary windows) - BringWindowToTop
                    // additionally raises the window in the z-order.
                    var h = new IntPtr(hwnd);
                    NativeMethods.SetForegroundWindow(h);
                    NativeMethods.BringWindowToTop(h);

                    LoggingService.Info("Window activated: hwnd=" + hwnd);
                    Reconcile("activate");
                    return;
                }
                LoggingService.Warn("Activate: window not found, hwnd=" + hwnd);
                Reconcile("activate-miss");
            }
            catch (Exception ex)
            {
                LoggingService.Error("ActivateWindow failed", ex);
            }
        }

        private void ActivateProtectedView(int key)
        {
            foreach (Word.ProtectedViewWindow pv in _word.ProtectedViewWindows)
            {
                try
                {
                    if (pv == null || PvKey(PvSourceFullPath(pv)) != key) continue;
                    pv.Activate();
                    LoggingService.Info("PV window activated: key=" + key);
                    return;
                }
                catch (Exception ex)
                {
                    LoggingService.Warn("PV activate failed: " + ex.Message);
                    return;
                }
            }
        }

        /// <summary>
        /// The x on a tab: close the document through Word STANDARD dialog.
        /// wdDoNotSaveChanges is never used - data loss is impossible by construction.
        /// A cancel in the dialog => Word error 4198 => the tab stays.
        /// </summary>
        public void RequestCloseWindow(int hwnd)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab != null && tab.IsProtectedView)
                {
                    CloseProtectedView(hwnd);
                    return;
                }

                foreach (Word.Window w in _word.Windows)
                {
                    try
                    {
                        if (w.Hwnd != hwnd) continue;
                        var doc = w.Document;
                        LoggingService.Info("Document closing requested: " + doc.Name);

                        bool lastWindowOfDoc = true;
                        try { lastWindowOfDoc = doc.Windows.Count <= 1; } catch (COMException) { }

                        if (lastWindowOfDoc)
                        {
                            // The last window of the document: close the document with the
                            // standard save prompt.
                            object saveOpt = Word.WdSaveOptions.wdPromptToSaveChanges;
                            doc.Close(ref saveOpt);
                        }
                        else
                        {
                            // The document has other windows: close only this window, the
                            // document stays open and no dialog is needed.
                            w.Close();
                        }
                        break;
                    }
                    catch (COMException ex)
                    {
                        if ((ex.ErrorCode & 0xFFFF) == WdErrCancelledByUser)
                            LoggingService.Info("Close cancelled by user (hwnd=" + hwnd + ")");
                        else
                            LoggingService.Warn("Close failed for hwnd=" + hwnd + ": " + ex.Message);
                        break;
                    }
                }

                RequestDeferredReconcile("after-close-request");
            }
            catch (Exception ex)
            {
                LoggingService.Error("RequestCloseWindow failed", ex);
            }
        }

        private void CloseProtectedView(int key)
        {
            foreach (Word.ProtectedViewWindow pv in _word.ProtectedViewWindows)
            {
                try
                {
                    if (pv == null || PvKey(PvSourceFullPath(pv)) != key) continue;
                    pv.Close(); // PV is read-only - data loss is impossible
                    break;
                }
                catch (Exception ex)
                {
                    LoggingService.Warn("PV close failed: " + ex.Message);
                    break;
                }
            }
            RequestDeferredReconcile("after-pv-close");
        }

        /// <summary>The "+" button: create a new empty Word document.</summary>
        public void CreateNewDocument()
        {
            try
            {
                _word.Documents.Add();
                LoggingService.Info("New document created via [+]");
                RequestReconcile("new-doc-button");
            }
            catch (Exception ex)
            {
                LoggingService.Error("CreateNewDocument failed", ex);
            }
        }

        /// <summary>
        /// Dragging a tab: move it to a new index. The order lives in the manager model,
        /// so it is the same in every strip and in the all-tabs menu; Reconcile keeps the
        /// positions of existing tabs (a stable order).
        /// Then comes the same "pinned first" invariant as in Reconcile (OrderBy is stable:
        /// a tab physically dropped inside the other group "sticks" to the border of its
        /// own group on the side it was dragged from - predictable, and with no special
        /// code to clamp the index under the cursor).
        /// </summary>
        public void ReorderTab(int hwnd, int newIndex)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab == null) return;

                var reordered = new List<DocumentTabModel>(_tabs);
                reordered.Remove(tab);
                if (newIndex < 0) newIndex = 0;
                if (newIndex > reordered.Count) newIndex = reordered.Count;
                reordered.Insert(newIndex, tab);
                reordered = reordered.OrderBy(t => t.IsPinned ? 0 : 1).ToList();

                if (DocumentTabModel.ListsEqual(_tabs, reordered)) return;
                _tabs = reordered;
                LoggingService.Info("Tab reordered: hwnd=" + hwnd + " -> index " + newIndex);
                _order.SaveOrder(_tabs);
                _paneManager.PushTabs(_tabs);
            }
            catch (Exception ex)
            {
                LoggingService.Error("ReorderTab failed", ex);
            }
        }

        /// <summary>
        /// Context menu "Close others": every document is closed through Word STANDARD
        /// dialog (wdPromptToSaveChanges); a cancel in the dialog keeps that tab and does
        /// not stop the rest from closing. Other pinned tabs are left alone - that is what
        /// pinning is for (protection from an accidental mass close); a single close (the x,
        /// the middle button) ignores pinning, because it is a deliberate action of the
        /// user on one particular tab.
        /// </summary>
        public void CloseOtherWindows(int hwnd)
        {
            try
            {
                var others = _tabs.Where(t => t.Hwnd != hwnd && !t.IsPinned).Select(t => t.Hwnd).ToList();
                LoggingService.Info("Close others requested: keep hwnd=" + hwnd + " (+pinned), closing " + others.Count);
                foreach (var h in others)
                    RequestCloseWindow(h);
            }
            catch (Exception ex)
            {
                LoggingService.Error("CloseOtherWindows failed", ex);
            }
        }

        /// <summary>
        /// Context menu "Pin tab" / "Unpin tab": pinned tabs always come first (see
        /// Reconcile/ReorderTab) and are not closed by "Close others". The model is shared,
        /// so pinning is immediately visible in every Word window.
        /// </summary>
        public void SetTabPinned(int hwnd, bool pinned)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab == null) return;

                _order.SetPinned(tab.FullPath, hwnd, pinned);
                LoggingService.Info("Tab pin " + (pinned ? "set" : "cleared") +
                    " (hwnd=" + hwnd + ", name=" + tab.Caption + ")");

                // Changed tabs are fresh clones, not a field mutated in place (see
                // DocumentTabModel.Clone): otherwise the object already handed to the
                // strips by a previous PushTabs is "replaced" by the same instance with a
                // new value, the by-value comparison in TabStripControl.UpdateTabs finds it
                // equal to itself and skips the repaint - which shows up exactly when
                // pinning does NOT move the tab (pinning the first or the only open tab,
                // for instance).
                var updated = new List<DocumentTabModel>(_tabs.Count);
                foreach (var t in _tabs)
                {
                    bool resolved = _order.IsPinned(t.FullPath, t.Hwnd);
                    if (resolved != t.IsPinned)
                    {
                        var clone = t.Clone();
                        clone.IsPinned = resolved;
                        updated.Add(clone);
                    }
                    else
                    {
                        updated.Add(t);
                    }
                }
                var reordered = updated.OrderBy(t => t.IsPinned ? 0 : 1).ToList();
                _tabs = reordered;
                _order.SaveOrder(_tabs);
                _paneManager.PushTabs(_tabs);
            }
            catch (Exception ex)
            {
                LoggingService.Error("SetTabPinned failed", ex);
            }
        }

        /// <summary>
        /// Settings window: switch the Ctrl+Tab interception on or off. Applied
        /// immediately (the hook is installed / removed right here), no Word restart.
        /// </summary>
        public void SetHotkeyEnabled(bool enabled)
        {
            try
            {
                HotkeySettings.Save(enabled);
                if (enabled)
                {
                    InstallHotkey();
                }
                else if (_keyboardHook != null)
                {
                    _keyboardHook.Dispose();
                    _keyboardHook = null;
                }
                LoggingService.Info("Ctrl+Tab hotkey " + (enabled ? "enabled" : "disabled") + " from settings");
            }
            catch (Exception ex)
            {
                LoggingService.Error("SetHotkeyEnabled failed", ex);
            }
        }

        /// <summary>Context menu "Open file location": Explorer with the file selected.</summary>
        public void OpenContainingFolder(int hwnd)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab == null || string.IsNullOrEmpty(tab.FullPath)) return;
                if (!System.IO.Path.IsPathRooted(tab.FullPath)) return; // an unsaved document

                if (System.IO.File.Exists(tab.FullPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + tab.FullPath + "\"");
                }
                else
                {
                    string dir = System.IO.Path.GetDirectoryName(tab.FullPath);
                    if (!string.IsNullOrEmpty(dir) && System.IO.Directory.Exists(dir))
                        System.Diagnostics.Process.Start("explorer.exe", "\"" + dir + "\"");
                }
                LoggingService.Info("Open folder: " + tab.FullPath);
            }
            catch (Exception ex)
            {
                LoggingService.Error("OpenContainingFolder failed", ex);
            }
        }

        /// <summary>
        /// Context menu "Tab color": set (or clear, with null) the colour.
        /// For saved documents the colour is tied to the file path and survives a Word
        /// restart; for unsaved ones it lives until the document is saved or closed.
        /// The model is shared, so the colour is immediately visible in every Word window.
        /// </summary>
        public void SetTabColor(int hwnd, System.Drawing.Color? color)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == hwnd);
                if (tab == null) return;

                _colors.SetColor(tab.FullPath, hwnd, color);
                LoggingService.Info("Tab color " +
                    (color.HasValue ? "set: " + TabColorStore.FormatRgb(color.Value) : "cleared") +
                    " (hwnd=" + hwnd + ", name=" + tab.Caption + ")");

                // Update every tab (a document may have several windows): changed ones are
                // fresh clones, not a field mutated in place (see DocumentTabModel.Clone and
                // the comment in SetTabPinned - the same technique for the same reason:
                // otherwise the repaint would sometimes be skipped).
                bool changed = false;
                var updated = new List<DocumentTabModel>(_tabs.Count);
                foreach (var t in _tabs)
                {
                    var resolved = _colors.Resolve(t.FullPath, t.Hwnd);
                    if (!Nullable.Equals(t.TabColor, resolved))
                    {
                        var clone = t.Clone();
                        clone.TabColor = resolved;
                        updated.Add(clone);
                        changed = true;
                    }
                    else
                    {
                        updated.Add(t);
                    }
                }
                if (changed)
                {
                    _tabs = updated;
                    _paneManager.PushTabs(_tabs);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("SetTabColor failed", ex);
            }
        }

        /// <summary>Ctrl+Tab / Ctrl+Shift+Tab: next (+1) / previous (-1) tab, wrapping around.</summary>
        public void ActivateAdjacent(int delta)
        {
            try
            {
                if (_tabs.Count == 0) return;
                int current = _tabs.FindIndex(t => t.IsActive);
                int next = current < 0
                    ? 0
                    : (current + delta % _tabs.Count + _tabs.Count) % _tabs.Count;
                if (next == current) return;
                ActivateWindow(_tabs[next].Hwnd);
            }
            catch (Exception ex)
            {
                LoggingService.Error("ActivateAdjacent failed", ex);
            }
        }
    }
}
