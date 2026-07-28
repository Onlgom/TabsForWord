using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// One instance of the in-window tab strip for one Word window.
    ///
    /// How it is built: a WinForms control HostSurface (WS_CHILD; created without a
    /// parent, so WinForms parks it on its own hidden parking window) holding an
    /// ordinary TabStripControl (Dock=Fill). The host HWND is then moved with
    /// SetParent into the top-level Word window and positioned over the top edge of
    /// the document area (the non-invasive mode) or with the space reserved for it
    /// (the reserve mode).
    ///
    /// Every operation is idempotent and tolerant of the HWND being destroyed at any
    /// moment: Word destroys our HWND together with its own window (WinForms learns
    /// about it properly through WM_NCDESTROY), after which the host is either
    /// recreated or removed by the manager.
    /// </summary>
    internal sealed class NativeTabHost : IDisposable
    {
        private const int LocateRetryMs = 500;   // do not search for the zone more often than every 500 ms
        private const int SizeSettleMs = 125;    // quiet time after a resize before the anchor is moved
                                                 // (was 250; halved - the ruler comes back faster)
        private const int ZoneUpgradeMs = 3000;  // how often to try upgrading a degraded zone to a proper one
        private const int ZoneUpgradeAttempts = 10; // how many such attempts to make (then the layout counts as non-standard)

        private readonly NativeTabHostManager _owner;
        private readonly IntPtr _wordTopLevel;

        /// <summary>Result of looking for the document area (see EnsureZone).</summary>
        private enum ZoneState
        {
            Ok,           // the zone is there, the strip can be laid out
            NotFound,     // no zone: the window is being built, is dying, or something is wrong
            FullPageUi    // no zone BY DESIGN: the File menu (Backstage) is open
        }

        private HostSurface _surface;
        private TabStripControl _strip;
        private WordNativeWindowLocator.ZoneChoice _zone;
        private Rectangle _lastRect = Rectangle.Empty;   // the applied rect (Word client coordinates)
        private int _lastLocateTick;
        private bool _locateFailedRecently;
        private int _locateFailCount;                    // failed zone searches in a row
        private ZoneState _lastZoneState = ZoneState.Ok; // answer of the last real search (reused while throttled)
        private bool _fullPageUiLogged;                  // the File menu episode has already been logged
        private int _zoneUpgradeAttempts;                // attempts to upgrade a degraded zone to a proper one
        private bool _failureDumped;
        private bool _disposed;

        // Space reservation (switched on with reserve=1).
        // The key to correctness: remember where WE put the top of the anchor and where
        // its natural place is. While the anchor stays where we put it, the strip lives
        // at its natural place (it does not "ride" the shifted anchor); as soon as Word
        // moves the anchor itself, its current top counts as natural again. This survives
        // a change of strip height (collapse/expand) with no downward drift and no
        // "ghosts" left on screen.
        private bool _reserveGivenUp;                 // given up for good (three pauses in a row)
        private bool _reserveApplied;                 // the anchor is currently shifted by us
        private int _shiftedAnchorTopClient;          // anchor top as we set it (client coordinates)
        private int _naturalAnchorTopClient;          // natural anchor top before the shift
        private int _reserveWindowStartTick;
        private int _reserveAdjustCount;
        private int _reserveSuspendUntilTick;         // pause after a burst of oscillations
        private bool _reserveSuspended;
        private int _reserveSuspendCount;
        private int _lastShiftLogTick;

        // Live resize with the mouse: Word re-lays out dozens of times per second.
        // While the window size is changing the anchor is left alone (the strip keeps
        // following the window) - one adjustment after SizeSettleMs of quiet.
        private Size _lastClientSize;
        private int _sizeQuietUntilTick;

        // Live background matching (bg=auto): the shade of the Word work area differs
        // between an active and an inactive window, so the colour is tracked per window.
        private Color? _autoBg;
        private int _lastBgSampleTick;
        private int _bgTrackUntilTick;
        private bool _fgKnown;
        private bool _wasForeground;
        private bool _bgLogged;

        public NativeTabHost(NativeTabHostManager owner, IntPtr wordTopLevel)
        {
            _owner = owner;
            _wordTopLevel = wordTopLevel;
        }

        public IntPtr WordWindowHandle { get { return _wordTopLevel; } }

        public IntPtr AnchorHwnd
        {
            get { var z = _zone; return z != null ? z.AnchorHwnd : IntPtr.Zero; }
        }

        public IntPtr CanvasHwnd
        {
            get { var z = _zone; return z != null ? z.CanvasHwnd : IntPtr.Zero; }
        }

        /// <summary>Host HWND; IntPtr.Zero when it is not created or already destroyed.
        /// Touching Control.Handle on a dead handle would recreate the window, so
        /// IsHandleCreated is checked first.</summary>
        public IntPtr HostHandle
        {
            get
            {
                var s = _surface;
                return s != null && s.IsHandleCreated ? s.Handle : IntPtr.Zero;
            }
        }

        public bool IsVisible { get { return !_lastRect.IsEmpty; } }

        /// <summary>The rect applied last (for tests and diagnostics).</summary>
        internal Rectangle LastAppliedRect { get { return _lastRect; } }

        // ------------------------------------------------------------------
        // Creation
        // ------------------------------------------------------------------

        /// <summary>
        /// Creates the host HWND and moves it into the Word window. Idempotent: a live
        /// host is not recreated; one killed by Word is created anew.
        /// </summary>
        public bool EnsureCreated()
        {
            if (_disposed) return false;
            try
            {
                if (_surface != null && _surface.IsHandleCreated && NativeWin32.IsWindow(_surface.Handle))
                    return true;

                DestroySurface(); // dead handle - start over

                if (!NativeWin32.IsWindow(_wordTopLevel)) return false;

                _surface = new HostSurface();
                _strip = new TabStripControl { Dock = DockStyle.Fill };
                WireStrip(_strip);
                _strip.SetCollapsed(_owner.Collapsed);
                _strip.SetActiveTabFlushBottom(true); // the document is right below - fill down to the line
                _surface.Controls.Add(_strip);

                // Force the HWND into existence (WinForms parking), then move it into Word.
                var handle = _surface.Handle;
                var prevParent = NativeWin32.SetParent(handle, _wordTopLevel);
                if (prevParent == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    LoggingService.Error("NativeTabHost: SetParent failed (err=" + err +
                        ", word=0x" + _wordTopLevel.ToInt64().ToString("X") + ")");
                    DestroySurface();
                    return false;
                }

                _lastRect = Rectangle.Empty;
                LoggingService.Info("NativeTabHost created: word=0x" + _wordTopLevel.ToInt64().ToString("X") +
                    ", host=0x" + handle.ToInt64().ToString("X"));
                return true;
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost create failed", ex);
                DestroySurface();
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Layout
        // ------------------------------------------------------------------

        /// <summary>
        /// Recomputes and applies the strip position. Cheap and idempotent: called on
        /// every Reconcile (about 1 Hz) and on every Word window move (WinEvent).
        /// With no geometry change it only re-asserts the z-order.
        ///
        /// Returns false when the native mechanics did NOT work (no surface, the document
        /// zone was not found, the geometry fell apart) - the manager counts such failures
        /// and after a run of them switches to the CTP fallback. Benign reasons for doing
        /// nothing (the window is minimised or dying) return true: those are not failures.
        /// </summary>
        public bool UpdateLayout()
        {
            if (_disposed) return true;
            // All geometry is in physical pixels: Word may keep the thread in a
            // system-aware context where coordinates are virtualised, and the strip on a
            // "non-system" monitor would end up with the wrong scale.
            var prevCtx = NativeWin32.SetThreadDpiAwarenessContextSafe(
                NativeWin32.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            try
            {
                if (!NativeWin32.IsWindow(_wordTopLevel)) return true;   // the manager will clean up
                if (NativeWin32.IsIconic(_wordTopLevel)) return true;    // minimised - geometry is meaningless
                if (!EnsureCreated()) return false;

                var zoneState = EnsureZone();
                if (zoneState != ZoneState.Ok)
                {
                    // While the File menu is open there is no document area at all, and
                    // hiding the strip is the RIGHT outcome - so it is a benign reason to
                    // do nothing (true), like a minimised window. Returning false there
                    // would feed the manager's failure counter, and with a single Word
                    // window open five such syncs in a row would drop the add-in into the
                    // CTP fallback in the middle of an ordinary trip to the File menu.
                    bool benign = zoneState == ZoneState.FullPageUi;
                    HideSurface(benign ? "full-page Word UI" : "no content zone");
                    return benign;
                }

                // Fresh anchor geometry: it cannot be cached, the Word window moves.
                NativeWin32.RECT anchorRect;
                if (!NativeWin32.GetWindowRect(_zone.AnchorHwnd, out anchorRect))
                {
                    _zone = null;
                    HideSurface("anchor died");
                    return false;
                }

                NativeWin32.RECT client;
                if (!NativeWin32.GetClientRect(_wordTopLevel, out client)) { HideSurface("client rect failed"); return false; }
                var clientSize = new Size(client.Right, client.Bottom);

                // Live resize: while the size is changing the anchor is left alone.
                if (clientSize != _lastClientSize)
                {
                    if (!_lastClientSize.IsEmpty)
                        _sizeQuietUntilTick = unchecked(Environment.TickCount + SizeSettleMs);
                    _lastClientSize = clientSize;
                }

                var anchorTop = new NativeWin32.POINT { X = anchorRect.Left, Y = anchorRect.Top };
                if (!NativeWin32.ScreenToClient(_wordTopLevel, ref anchorTop)) { HideSurface("map failed"); return false; }
                int anchorWidth = anchorRect.Right - anchorRect.Left;

                int panelH = CurrentPanelHeight();
                bool reserveActive = ReserveActiveNow();

                // Natural top of the document area. If the anchor sits exactly where we
                // shifted it, take the remembered natural place; otherwise Word moved it
                // itself and its current top is the natural one.
                int naturalTop;
                if (reserveActive && _reserveApplied && anchorTop.Y == _shiftedAnchorTopClient)
                {
                    naturalTop = _naturalAnchorTopClient;
                }
                else
                {
                    naturalTop = anchorTop.Y;
                    _reserveApplied = false; // the previous shift is no longer ours / no longer valid
                }

                var rect = NativeHostLayoutMath.ComputePanelRect(clientSize, naturalTop, panelH,
                    anchorTop.X, anchorWidth);
                // Degenerate geometry of the window itself (a tiny client while it is being
                // created or destroyed) is not a native-mechanics failure: true.
                if (rect.IsEmpty) { HideSurface("degenerate rect"); return true; }

                ApplyPanelBackground(rect);
                ApplyRect(rect);

                // The anchor is only moved after the size has been quiet: while the frame is
                // dragged Word re-lays out continuously anyway, and fighting it produces
                // jitter and bursts of oscillation.
                bool sizeSettled = unchecked(Environment.TickCount - _sizeQuietUntilTick) >= 0;
                if (reserveActive && sizeSettled)
                    EnsureAnchorPlacement(rect, naturalTop, anchorTop.Y, anchorRect);
                return true;
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost.UpdateLayout failed", ex);
                return false;
            }
            finally
            {
                if (prevCtx != IntPtr.Zero) NativeWin32.SetThreadDpiAwarenessContextSafe(prevCtx);
            }
        }

        private ZoneState EnsureZone()
        {
            var z = _zone;
            if (z != null
                && NativeWin32.IsWindow(z.AnchorHwnd)
                && NativeWin32.IsWindowVisible(z.AnchorHwnd)
                && NativeWin32.IsWindow(z.CanvasHwnd)
                && NativeWin32.GetAncestor(z.AnchorHwnd, NativeWin32.GA_ROOT) == _wordTopLevel)
            {
                // The zone is alive. If it was picked by a fallback path, periodically try
                // to upgrade it to a proper one: the Word window may still have been under
                // construction at the time of the first search (measured on a real machine:
                // the canvas temporarily covers the whole client area), and the window would
                // then live with the wrong anchor for the rest of the session.
                if (z.Degraded) TryUpgradeZone();
                _fullPageUiLogged = false;
                _lastZoneState = ZoneState.Ok;
                return ZoneState.Ok;
            }

            // A repeated search happens at most every LocateRetryMs (Locate enumerates ~a hundred windows).
            int now = Environment.TickCount;
            if (_locateFailedRecently && now - _lastLocateTick < LocateRetryMs) return _lastZoneState;
            _lastLocateTick = now;

            int dpi = NativeWin32.GetDpiForWindowSafe(_wordTopLevel, 96);
            bool fullPageUi;
            _zone = WordNativeWindowLocator.Locate(_wordTopLevel, dpi, false, out fullPageUi);
            _locateFailedRecently = _zone == null;

            if (_zone == null)
            {
                if (fullPageUi)
                {
                    // The File menu (Backstage) hides the whole document subtree. Nothing
                    // is broken: no failure counted, no dump, and one INFO line per episode
                    // so the log still explains why the strip disappeared for a while.
                    if (!_fullPageUiLogged)
                    {
                        _fullPageUiLogged = true;
                        LoggingService.Info("Full-page Word UI is open (the File menu): no document area," +
                            " the strip stays hidden until it closes: word=0x" +
                            _wordTopLevel.ToInt64().ToString("X"));
                    }
                    _locateFailCount = 0;
                    _lastZoneState = ZoneState.FullPageUi;
                    return ZoneState.FullPageUi;
                }

                // Transient failures (the window is still being built or already dying) are
                // normal; a tree dump is written only when the zone is not found steadily.
                // The threshold is 6 rather than 3: the locator deliberately answers "not
                // found" until Word has built its layout (otherwise the strip lands over the
                // ribbon), so 1-3 refusals in a row at startup are the normal picture, not a
                // reason to drop a tree dump into the log folder.
                _locateFailCount++;
                if (_locateFailCount >= 6 && !_failureDumped)
                {
                    _failureDumped = true;
                    WindowTreeDiagnostics.DumpToFile(_wordTopLevel, "locator failed to find content zone (x" +
                        _locateFailCount + ")");
                }
                _lastZoneState = ZoneState.NotFound;
                return ZoneState.NotFound;
            }

            _locateFailCount = 0;
            _fullPageUiLogged = false;
            _zoneUpgradeAttempts = 0;   // a new zone gets its own budget of upgrade attempts
            // The anchor changed - previous reserve adjustments no longer apply.
            _reserveApplied = false;
            _lastZoneState = ZoneState.Ok;
            return ZoneState.Ok;
        }

        /// <summary>
        /// Tries to replace a degraded zone with a proper one (canvas + container under
        /// the ribbon). Called at most every ZoneUpgradeMs. If there is no better zone,
        /// the current one stays in force (the strip neither blinks nor hides).
        /// The previous anchor is put back before the replacement: otherwise a window we
        /// shifted would stay shifted forever.
        /// </summary>
        private void TryUpgradeZone()
        {
            // Word finishes its layout within a fraction of a second. If no proper anchor
            // appears within ZoneUpgradeAttempts tries, the layout of this window really is
            // non-standard and there is no point searching forever.
            if (_zoneUpgradeAttempts >= ZoneUpgradeAttempts) return;
            int now = Environment.TickCount;
            if (now - _lastLocateTick < ZoneUpgradeMs) return;
            _lastLocateTick = now;
            _zoneUpgradeAttempts++;

            int dpi = NativeWin32.GetDpiForWindowSafe(_wordTopLevel, 96);
            var better = WordNativeWindowLocator.Locate(_wordTopLevel, dpi, true);
            if (better == null || better.Degraded) return;
            if (better.AnchorHwnd == _zone.AnchorHwnd) return;

            TryRestoreAnchor("zone upgrade");
            _zone = better;
            _reserveApplied = false;
            LoggingService.Info("Zone upgraded to full anchor: word=0x" +
                _wordTopLevel.ToInt64().ToString("X") + ", anchor=0x" +
                better.AnchorHwnd.ToInt64().ToString("X") + " \"" + better.AnchorClass + "\"");
        }

        private int CurrentPanelHeight()
        {
            // Every window draws in the DPI space it lives in. WinForms inside Word creates
            // its windows system-aware (verified with GetWindowDpiAwarenessContext on a real
            // Word), and on a monitor with a "non-system" scale Windows itself scales their
            // image and coordinates. Hence: the drawing scale of the strip = the DPI of ITS
            // OWN window (the space its pixels live in before virtualisation), while the
            // physical height of the strip inside the Word window is a conversion from that
            // space into the per-monitor DPI of the Word window. The formula is universal:
            // if the strip window ever becomes per-monitor aware, both DPIs coincide and
            // the conversion becomes an identity.
            int windowDpi = NativeWin32.GetDpiForWindowSafe(_wordTopLevel, 96);
            var strip = _strip;
            if (strip != null)
            {
                try
                {
                    int stripDpi = strip.IsHandleCreated
                        ? NativeWin32.GetDpiForWindowSafe(strip.Handle, windowDpi)
                        : windowDpi;
                    strip.SetDpiOverride(stripDpi);
                    return NativeHostLayoutMath.ScaleHeight(strip.DesiredContentHeight, stripDpi, windowDpi);
                }
                catch { /* the control is in a transient state - the fallback below computes it */ }
            }
            return NativeHostLayoutMath.Scale96(37, windowDpi);
        }

        private void ApplyRect(Rectangle rect)
        {
            var handle = HostHandle;
            if (handle == IntPtr.Zero) return;

            // We compare against the ACTUAL window geometry rather than the _lastRect cache:
            // when the Word window moved to a monitor with a different scale, WinForms
            // resized the host window behind our back (WM_DPICHANGED_AFTERPARENT) while the
            // cache still matched the desired rect - and the strip got stuck at the wrong
            // size. Those messages are now swallowed (HostSurface/TabStripControl), and
            // comparing against reality is the second line of defence: at about 1 Hz the
            // strip repairs itself, whoever moved it.
            var actual = Rectangle.Empty;
            NativeWin32.RECT wr;
            if (NativeWin32.GetWindowRect(handle, out wr))
            {
                var tl = new NativeWin32.POINT { X = wr.Left, Y = wr.Top };
                if (NativeWin32.ScreenToClient(_wordTopLevel, ref tl))
                    actual = new Rectangle(tl.X, tl.Y, wr.Right - wr.Left, wr.Bottom - wr.Top);
            }

            bool changed = rect != actual;
            uint flags = NativeWin32.SWP_NOACTIVATE | NativeWin32.SWP_SHOWWINDOW;
            if (!changed) flags |= NativeWin32.SWP_NOMOVE | NativeWin32.SWP_NOSIZE; // z-order / show only

            // The strip goes below the ribbon shadow window (DropShadow): the Word shadow
            // falls onto it, so the strip does not look glued onto the ribbon. No shadow - just to the top.
            IntPtr insertAfter = NativeWin32.HWND_TOP;
            var z = _zone;
            if (z != null && z.ShadowHwnd != IntPtr.Zero
                && NativeWin32.IsWindow(z.ShadowHwnd) && NativeWin32.IsWindowVisible(z.ShadowHwnd))
            {
                insertAfter = z.ShadowHwnd;
            }

            NativeWin32.SetWindowPos(handle, insertAfter,
                rect.X, rect.Y, rect.Width, rect.Height, flags);

            if (_lastRect.IsEmpty && changed)
                LoggingService.Info("NativeTabHost repositioned (shown): word=0x" +
                    _wordTopLevel.ToInt64().ToString("X") + ", rect=" + rect.X + "," + rect.Y +
                    "," + rect.Width + "x" + rect.Height);

            _lastRect = rect;
        }

        private void HideSurface(string reason)
        {
            try
            {
                var handle = HostHandle;
                if (handle != IntPtr.Zero && NativeWin32.IsWindow(handle) && !_lastRect.IsEmpty)
                {
                    NativeWin32.ShowWindow(handle, NativeWin32.SW_HIDE);
                    LoggingService.Info("NativeTabHost hidden (" + reason + "): word=0x" +
                        _wordTopLevel.ToInt64().ToString("X"));
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost hide failed", ex);
            }
            finally
            {
                _lastRect = Rectangle.Empty;
            }
        }

        // ------------------------------------------------------------------
        // Space reservation (reserve=1, the default since v1.6.0)
        // ------------------------------------------------------------------

        /// <summary>Reservation is active (not given up and not paused after oscillations).</summary>
        private bool ReserveActiveNow()
        {
            if (!TabHostSettings.ReserveSpace || _reserveGivenUp) return false;
            if (_reserveSuspended)
            {
                if (unchecked(Environment.TickCount - _reserveSuspendUntilTick) < 0) return false;
                _reserveSuspended = false; // the pause is over - try again
            }
            return true;
        }

        /// <summary>
        /// Puts the anchor top exactly under the bottom of the strip (both ways: down
        /// after expanding, up after collapsing). The anchor bottom is fixed and the
        /// calculation always starts from the current actual geometry - nothing accumulates.
        /// A burst of oscillation (a fight with the Word layout) means a 30 s pause;
        /// three bursts in a row mean the window goes to overlay for good.
        /// </summary>
        private void EnsureAnchorPlacement(Rectangle panelRect, int naturalTop,
            int anchorTopClientY, NativeWin32.RECT anchorScreenRect)
        {
            try
            {
                int desiredTop = panelRect.Bottom;
                if (anchorTopClientY == desiredTop)
                {
                    // Already as it should be - just record the bookkeeping.
                    _reserveApplied = true;
                    _shiftedAnchorTopClient = desiredTop;
                    _naturalAnchorTopClient = naturalTop;
                    return;
                }

                int delta = desiredTop - anchorTopClientY; // >0 = down, <0 = up (the strip got shorter)
                int newHeight = anchorScreenRect.Height - delta;
                if (newHeight < 60) return; // there is almost no window left - leave it alone

                // Guard against a layout loop. The threshold is forgiving: manual clicks on
                // collapse/expand and a live resize produce several legitimate adjustments in a row.
                int now = Environment.TickCount;
                if (now - _reserveWindowStartTick > 2000)
                {
                    _reserveWindowStartTick = now;
                    _reserveAdjustCount = 0;
                }
                if (++_reserveAdjustCount > 8)
                {
                    _reserveSuspendCount++;
                    TryRestoreAnchor("oscillation");
                    if (_reserveSuspendCount >= 3)
                    {
                        _reserveGivenUp = true;
                        LoggingService.Warn("Reserve mode disabled for word=0x" +
                            _wordTopLevel.ToInt64().ToString("X") + " (persistent layout oscillation)");
                    }
                    else
                    {
                        _reserveSuspended = true;
                        _reserveSuspendUntilTick = unchecked(now + 30000);
                        LoggingService.Warn("Reserve mode suspended for 30s (word=0x" +
                            _wordTopLevel.ToInt64().ToString("X") + ", layout oscillation)");
                    }
                    return;
                }

                // Coordinates for SetWindowPos are in the client space of the anchor PARENT.
                var parent = NativeWin32.GetAncestor(_zone.AnchorHwnd, NativeWin32.GA_PARENT);
                if (parent == IntPtr.Zero || !NativeWin32.IsWindow(parent)) return;
                var topLeft = new NativeWin32.POINT { X = anchorScreenRect.Left, Y = anchorScreenRect.Top };
                if (!NativeWin32.ScreenToClient(parent, ref topLeft)) return;

                NativeWin32.SetWindowPos(_zone.AnchorHwnd, IntPtr.Zero,
                    topLeft.X, topLeft.Y + delta, anchorScreenRect.Width, newHeight,
                    NativeWin32.SWP_NOZORDER | NativeWin32.SWP_NOACTIVATE);
                _reserveApplied = true;
                _shiftedAnchorTopClient = desiredTop;
                _naturalAnchorTopClient = naturalTop;

                if (now - _lastShiftLogTick > 500)
                {
                    _lastShiftLogTick = now;
                    LoggingService.Info("Reserve: anchor moved by " + delta + "px (word=0x" +
                        _wordTopLevel.ToInt64().ToString("X") + ")");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("EnsureAnchorPlacement failed", ex);
            }
        }

        /// <summary>
        /// Puts the anchor back to its natural place (the bottom is fixed). Called when
        /// reservation is paused or given up and when the host is destroyed, so that the
        /// Word layout is never left in a shifted state.
        /// </summary>
        private void TryRestoreAnchor(string reason)
        {
            try
            {
                if (!_reserveApplied) return;
                _reserveApplied = false;

                var z = _zone;
                if (z == null || !NativeWin32.IsWindow(z.AnchorHwnd)) return;
                if (!NativeWin32.IsWindow(_wordTopLevel)) return;

                NativeWin32.RECT r;
                if (!NativeWin32.GetWindowRect(z.AnchorHwnd, out r)) return;
                var probe = new NativeWin32.POINT { X = r.Left, Y = r.Top };
                if (!NativeWin32.ScreenToClient(_wordTopLevel, ref probe)) return;
                if (probe.Y != _shiftedAnchorTopClient) return; // Word has already moved it itself

                int delta = _naturalAnchorTopClient - probe.Y; // upwards (the bottom is fixed)
                var parent = NativeWin32.GetAncestor(z.AnchorHwnd, NativeWin32.GA_PARENT);
                if (parent == IntPtr.Zero || !NativeWin32.IsWindow(parent)) return;
                var topLeft = new NativeWin32.POINT { X = r.Left, Y = r.Top };
                if (!NativeWin32.ScreenToClient(parent, ref topLeft)) return;

                NativeWin32.SetWindowPos(z.AnchorHwnd, IntPtr.Zero,
                    topLeft.X, topLeft.Y + delta, r.Width, r.Height - delta,
                    NativeWin32.SWP_NOZORDER | NativeWin32.SWP_NOACTIVATE);
                LoggingService.Info("Reserve: anchor restored (" + reason + ", word=0x" +
                    _wordTopLevel.ToInt64().ToString("X") + ")");
            }
            catch (Exception ex)
            {
                LoggingService.Error("TryRestoreAnchor failed", ex);
            }
        }

        // ------------------------------------------------------------------
        // Data and state
        // ------------------------------------------------------------------

        public void PushTabs(System.Collections.Generic.List<DocumentTabModel> tabs)
        {
            try
            {
                var strip = _strip;
                if (strip != null && !strip.IsDisposed) strip.UpdateTabs(tabs);
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost.PushTabs failed", ex);
            }
        }

        public void SetCollapsed(bool collapsed)
        {
            try
            {
                var strip = _strip;
                if (strip != null && !strip.IsDisposed) strip.SetCollapsed(collapsed);
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost.SetCollapsed failed", ex);
            }
        }

        /// <summary>Language switch: rebuild the captions of this window's strip.</summary>
        public void RefreshLocalizedUi()
        {
            try
            {
                var strip = _strip;
                if (strip != null && !strip.IsDisposed) strip.RefreshLocalizedUi();
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost.RefreshLocalizedUi failed", ex);
            }
        }

        /// <summary>Tab size: a new multiplier and an immediate re-layout.</summary>
        public void SetUserScale(float scale)
        {
            try
            {
                var strip = _strip;
                if (strip != null && !strip.IsDisposed) strip.SetUserScale(scale);
                UpdateLayout();
            }
            catch (Exception ex)
            {
                LoggingService.Error("NativeTabHost.SetUserScale failed", ex);
            }
        }

        /// <summary>Is the input focus currently on the host or one of its children?</summary>
        public bool OwnsFocus(IntPtr focusHwnd)
        {
            if (focusHwnd == IntPtr.Zero) return false;
            var handle = HostHandle;
            if (handle == IntPtr.Zero) return false;
            return focusHwnd == handle || NativeWin32.IsChild(handle, focusHwnd);
        }

        private void WireStrip(TabStripControl strip)
        {
            strip.TabActivateRequested += _owner.HandleTabActivate;
            strip.TabCloseRequested += _owner.HandleTabClose;
            strip.NewDocumentRequested += _owner.HandleNewDocument;
            strip.CloseOthersRequested += _owner.HandleCloseOthers;
            strip.OpenFolderRequested += _owner.HandleOpenFolder;
            strip.ReorderRequested += _owner.HandleReorder;
            strip.CollapseToggleRequested += _owner.HandleCollapseToggle;
            strip.TabColorChangeRequested += _owner.HandleTabColorChange;
            strip.TabSizeChangeRequested += _owner.HandleTabSizeChange;
            strip.TabPinChangeRequested += _owner.HandleTabPinChange;
            strip.HotkeyToggleRequested += _owner.HandleHotkeyToggle;
            strip.LanguageChangeRequested += _owner.HandleLanguageChange;
        }

        private void UnwireStrip(TabStripControl strip)
        {
            strip.TabActivateRequested -= _owner.HandleTabActivate;
            strip.TabCloseRequested -= _owner.HandleTabClose;
            strip.NewDocumentRequested -= _owner.HandleNewDocument;
            strip.CloseOthersRequested -= _owner.HandleCloseOthers;
            strip.OpenFolderRequested -= _owner.HandleOpenFolder;
            strip.ReorderRequested -= _owner.HandleReorder;
            strip.CollapseToggleRequested -= _owner.HandleCollapseToggle;
            strip.TabColorChangeRequested -= _owner.HandleTabColorChange;
            strip.TabSizeChangeRequested -= _owner.HandleTabSizeChange;
            strip.TabPinChangeRequested -= _owner.HandleTabPinChange;
            strip.HotkeyToggleRequested -= _owner.HandleHotkeyToggle;
            strip.LanguageChangeRequested -= _owner.HandleLanguageChange;
        }

        /// <summary>
        /// Strip background per the bg= setting: spec = the specification 2b colour,
        /// #RRGGBB = a fixed one, auto = live matching to the work area of THIS particular
        /// Word window (the shade differs between an active and an inactive window).
        /// </summary>
        private void ApplyPanelBackground(Rectangle panelRect)
        {
            try
            {
                var setting = TabHostSettings.Background;
                Color? bg;
                if (setting == "spec")
                {
                    bg = null;
                }
                else
                {
                    var fixedColor = TabHostSettings.ParseColor(setting);
                    bg = fixedColor.HasValue ? fixedColor : AutoSampleBackground(panelRect);
                }

                var strip = _strip;
                if (strip != null && !strip.IsDisposed) strip.SetPanelBackgroundOverride(bg);
                var surface = _surface;
                if (surface != null && !surface.IsDisposed)
                {
                    var effective = bg.HasValue ? bg.Value : TabTheme.PanelBg;
                    if (surface.BackColor != effective) surface.BackColor = effective;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("ApplyPanelBackground failed", ex);
            }
        }

        /// <summary>
        /// Live background matching: after the activation changes, about 1.5 s of frequent
        /// probes (Word repaints its work area asynchronously), then rare checks.
        /// Returns null until the first successful measurement (the specification colour is used).
        /// </summary>
        private Color? AutoSampleBackground(Rectangle panelRect)
        {
            int now = Environment.TickCount;

            bool foreground = NativeWin32.GetForegroundWindow() == _wordTopLevel;
            if (!_fgKnown || foreground != _wasForeground)
            {
                _fgKnown = true;
                _wasForeground = foreground;
                _bgTrackUntilTick = unchecked(now + 1500);
            }

            bool tracking = unchecked(now - _bgTrackUntilTick) < 0;
            int interval = tracking ? 150 : 3000;
            if (!_autoBg.HasValue && interval > 1000) interval = 1000;

            if (unchecked(now - _lastBgSampleTick) >= interval)
            {
                _lastBgSampleTick = now;
                Color sampled;
                if (TrySampleWorkspaceScreen(panelRect, out sampled) && _autoBg != sampled)
                {
                    _autoBg = sampled;
                    if (!_bgLogged)
                    {
                        _bgLogged = true;
                        LoggingService.Info("Panel background matched to Word workspace: #" +
                            sampled.R.ToString("X2") + sampled.G.ToString("X2") + sampled.B.ToString("X2") +
                            " (word=0x" + _wordTopLevel.ToInt64().ToString("X") + ")");
                    }
                }
            }
            return _autoBg;
        }

        /// <summary>
        /// Reads the colour of the Word work area FROM THE SCREEN at points just below the
        /// strip (on the left, past the vertical ruler). Word draws its canvas with
        /// Direct2D, so GetPixel from its DC does not see the content - we read the
        /// composited screen frame instead. Every point is checked for belonging to our
        /// own window (WindowFromPoint): an area covered by a foreign window is not
        /// sampled, and the colour simply stays as it was.
        /// White is rejected - that is the page (web layout / a narrow window), not the background.
        /// </summary>
        private bool TrySampleWorkspaceScreen(Rectangle panelRect, out Color color)
        {
            color = Color.Empty;
            try
            {
                int dpi = NativeWin32.GetDpiForWindowSafe(_wordTopLevel, 96);
                int x0 = NativeHostLayoutMath.Scale96(45, dpi);
                int y0 = panelRect.Bottom + NativeHostLayoutMath.Scale96(45, dpi);
                var probes = new[]
                {
                    new Point(x0, y0),
                    new Point(x0 + NativeHostLayoutMath.Scale96(9, dpi), y0 + NativeHostLayoutMath.Scale96(10, dpi)),
                    new Point(x0 + NativeHostLayoutMath.Scale96(17, dpi), y0 + NativeHostLayoutMath.Scale96(4, dpi))
                };

                var samples = new System.Collections.Generic.List<Color>(probes.Length);
                IntPtr screenDc = NativeWin32.GetDC(IntPtr.Zero);
                if (screenDc == IntPtr.Zero) return false;
                try
                {
                    foreach (var probe in probes)
                    {
                        var pt = new NativeWin32.POINT { X = probe.X, Y = probe.Y };
                        if (!NativeWin32.ClientToScreen(_wordTopLevel, ref pt)) continue;

                        // is the point visible and does it belong to our Word window?
                        var at = NativeWin32.WindowFromPoint(pt);
                        if (at == IntPtr.Zero ||
                            NativeWin32.GetAncestor(at, NativeWin32.GA_ROOT) != _wordTopLevel) continue;

                        uint raw = NativeWin32.GetPixel(screenDc, pt.X, pt.Y);
                        if (raw == NativeWin32.CLR_INVALID) continue;
                        samples.Add(FromColorRef(raw));
                    }
                }
                finally
                {
                    NativeWin32.ReleaseDC(IntPtr.Zero, screenDc);
                }

                if (samples.Count < 2) return false;
                for (int i = 1; i < samples.Count; i++)
                {
                    if (Math.Abs(samples[i].R - samples[0].R) > 3 ||
                        Math.Abs(samples[i].G - samples[0].G) > 3 ||
                        Math.Abs(samples[i].B - samples[0].B) > 3) return false; // not uniform - not the background
                }

                var c = samples[0];
                if (c.R >= 250 && c.G >= 250 && c.B >= 250) return false; // white - the page, not the background

                color = c;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Color FromColorRef(uint colorRef) // COLORREF: 0x00BBGGRR
        {
            return Color.FromArgb(
                (int)(colorRef & 0xFF),
                (int)((colorRef >> 8) & 0xFF),
                (int)((colorRef >> 16) & 0xFF));
        }

        private void DestroySurface()
        {
            // Never leave the Word layout in a shifted state.
            TryRestoreAnchor("host destroy");

            var strip = _strip;
            var surface = _surface;
            _strip = null;
            _surface = null;
            _zone = null;
            _lastRect = Rectangle.Empty;
            _reserveApplied = false;

            if (strip != null)
            {
                try { UnwireStrip(strip); } catch { }
                try { strip.Dispose(); } catch { }
            }
            if (surface != null)
            {
                // Dispose destroys the HWND (if Word has not destroyed it already).
                try { surface.Dispose(); } catch { }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            DestroySurface();
            LoggingService.Info("NativeTabHost destroyed: word=0x" + _wordTopLevel.ToInt64().ToString("X"));
        }

        /// <summary>
        /// The WinForms surface of the host: WS_CHILD + CLIPCHILDREN/CLIPSIBLINGS,
        /// background in the strip colour (no flashes while the layout is recomputed).
        /// </summary>
        private sealed class HostSurface : Control
        {
            private const int WS_CLIPCHILDREN = 0x02000000;
            private const int WS_CLIPSIBLINGS = 0x04000000;

            public HostSurface()
            {
                BackColor = TabTheme.PanelBg;
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    var cp = base.CreateParams;
                    cp.Style |= WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
                    return cp;
                }
            }

            private const int WM_DPICHANGED_BEFOREPARENT = 0x02E2;
            private const int WM_DPICHANGED_AFTERPARENT = 0x02E3;

            /// <summary>
            /// A managed exception from a window procedure would bring Word down
            /// (0xE0434352) - swallow and log, this is the last line of defence.
            /// DPI change messages are swallowed: the scale is driven by NativeTabHost
            /// (the DPI of the Word window), while the WinForms auto-response resized the
            /// host window at the wrong moment when it moved between monitors.
            /// </summary>
            protected override void WndProc(ref Message m)
            {
                try
                {
                    if (m.Msg == WM_DPICHANGED_BEFOREPARENT || m.Msg == WM_DPICHANGED_AFTERPARENT)
                    {
                        m.Result = IntPtr.Zero;
                        return;
                    }
                    base.WndProc(ref m);
                }
                catch (Exception ex)
                {
                    LoggingService.Error("HostSurface WndProc failed (msg=0x" + m.Msg.ToString("X4") + ")", ex);
                }
            }
        }
    }
}
