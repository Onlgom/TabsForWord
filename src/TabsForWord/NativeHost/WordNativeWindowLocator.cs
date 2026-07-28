using System;
using System.Collections.Generic;
using System.Drawing;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// Locates the document area inside the top-level Word window (OpusApp).
    ///
    /// The tab strip has to be pinned to the top edge of the document area - right
    /// below the ribbon. The locator finds that edge from several signals rather than
    /// trusting a single window class name:
    ///
    ///  1. The main signal is the document canvas window of class "_WwG" (it has
    ///     existed reliably from Word 97 through Word M365). From the canvas we walk
    ///     up the ancestor chain and take the topmost container that starts BELOW the
    ///     top of the client area (that is, under the ribbon): usually "_WwB" (the
    ///     document frame with the rulers).
    ///  2. The fallback is the topmost visible child window that lies below the top
    ///     of the client area and is wide and tall enough to be the document area.
    ///
    /// Sometimes there is legitimately no document area at all: while the File menu
    /// (Backstage) is open Word hides its whole document subtree - see IsFullPageUiActive.
    ///
    /// Every choice carries a textual reason for the log.
    /// The class is split into pure logic (FindContentZone, covered by tests) and the
    /// Win32 snapshot (Snapshot), which needs live windows.
    /// </summary>
    internal static class WordNativeWindowLocator
    {
        /// <summary>Class of the Word document canvas (checked first, but not the only signal).</summary>
        internal const string CanvasClassName = "_WwG";

        /// <summary>
        /// Class of the Word full-page UI host: the File menu (Backstage) and everything
        /// opened from it - Print, Save As, Account, Options.
        /// </summary>
        internal const string FullPageUiClassName = "FullpageUIHost";

        /// <summary>How much of the client area a full-page UI window has to cover, in percent.</summary>
        internal const int FullPageUiCoverPercent = 90;

        /// <summary>Threshold for "starts below the client top", in px at 96 DPI.</summary>
        internal const int BelowTopThreshold96 = 8;

        /// <summary>Minimum height that must remain below the content (96 DPI).</summary>
        internal const int MinRemainingHeight96 = 80;

        /// <summary>Minimum share of the client width for a fallback candidate.</summary>
        internal const double MinFallbackWidthShare = 0.4;

        /// <summary>Snapshot of one window for pure analysis (no live HWNDs).</summary>
        internal sealed class WindowSnapshot
        {
            public IntPtr Hwnd;
            public IntPtr Parent;
            public string ClassName = string.Empty;
            public Rectangle ScreenRect;   // GetWindowRect
            public bool Visible;
        }

        /// <summary>Search result: the strip has an anchor (the content edge) and a canvas.</summary>
        internal sealed class ZoneChoice
        {
            public IntPtr AnchorHwnd;      // the window whose top = the top of the document area
            public string AnchorClass = string.Empty;
            public Rectangle AnchorRect;
            public IntPtr CanvasHwnd;      // the document canvas (used to return focus)
            public string CanvasClass = string.Empty;
            public IntPtr ShadowHwnd;      // the ribbon shadow window (DropShadow) - the strip goes below it
            public string Reason = string.Empty;

            /// <summary>
            /// The zone was chosen by a fallback path (anchor = canvas, or the generic
            /// fallback) rather than "canvas + container under the ribbon". Such a choice
            /// can be right (a non-standard Word layout), but it can also mean the window
            /// is still being built: the host periodically re-searches the zone so it can
            /// upgrade to a proper anchor once one appears.
            /// </summary>
            public bool Degraded;
        }

        // ------------------------------------------------------------------
        // Win32 snapshot
        // ------------------------------------------------------------------

        /// <summary>Snapshot of every descendant of a window. Tolerant of windows dying.</summary>
        internal static List<WindowSnapshot> Snapshot(IntPtr root)
        {
            var result = new List<WindowSnapshot>();
            try
            {
                NativeWin32.EnumChildWindows(root, delegate(IntPtr h, IntPtr l)
                {
                    try
                    {
                        var s = new WindowSnapshot { Hwnd = h };
                        s.Parent = NativeWin32.GetAncestor(h, NativeWin32.GA_PARENT);
                        s.ClassName = NativeWin32.GetClassNameSafe(h);
                        NativeWin32.RECT r;
                        if (NativeWin32.GetWindowRect(h, out r))
                            s.ScreenRect = new Rectangle(r.Left, r.Top, r.Width, r.Height);
                        s.Visible = NativeWin32.IsWindowVisible(h);
                        result.Add(s);
                    }
                    catch
                    {
                        // the window died during enumeration - skip it
                    }
                    return result.Count < 4096;
                }, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                LoggingService.Warn("Locator snapshot failed: " + ex.Message);
            }
            return result;
        }

        /// <summary>Client area of a window in screen coordinates.</summary>
        internal static bool TryGetClientScreenRect(IntPtr hwnd, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            try
            {
                NativeWin32.RECT client;
                if (!NativeWin32.GetClientRect(hwnd, out client)) return false;
                var origin = new NativeWin32.POINT { X = 0, Y = 0 };
                if (!NativeWin32.ClientToScreen(hwnd, ref origin)) return false;
                rect = new Rectangle(origin.X, origin.Y, client.Right, client.Bottom);
                return rect.Width > 0 && rect.Height > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>The full cycle: snapshot + analysis + a log line with the reason.</summary>
        internal static ZoneChoice Locate(IntPtr root, int dpi)
        {
            bool fullPageUi;
            return Locate(root, dpi, false, out fullPageUi);
        }

        /// <summary>
        /// The same, but able to stay quiet: background attempts to upgrade a degraded
        /// zone repeat every few seconds, and their result is only interesting when it is
        /// better than the current one (the caller logs that).
        /// </summary>
        internal static ZoneChoice Locate(IntPtr root, int dpi, bool quiet)
        {
            bool fullPageUi;
            return Locate(root, dpi, quiet, out fullPageUi);
        }

        /// <summary>
        /// The same, and it also says WHY there is no zone: fullPageUi = the document area
        /// does not exist right now because Word has put up its full-page UI (the File
        /// menu). That is a normal state, not a failure, so the caller neither counts it
        /// as one nor writes a window-tree dump - and it says nothing in the log at this
        /// level either (the host logs it once per episode).
        /// </summary>
        internal static ZoneChoice Locate(IntPtr root, int dpi, bool quiet, out bool fullPageUi)
        {
            fullPageUi = false;

            Rectangle clientScreen;
            if (!TryGetClientScreenRect(root, out clientScreen))
            {
                if (!quiet)
                    LoggingService.Warn("Locator: GetClientRect failed for root=0x" + root.ToInt64().ToString("X"));
                return null;
            }

            var nodes = Snapshot(root);
            var choice = FindContentZone(root, clientScreen, nodes, dpi);
            if (choice == null) fullPageUi = IsFullPageUiActive(clientScreen, nodes);
            if (quiet) return choice;
            if (choice != null)
            {
                LoggingService.Info("Candidate document host selected: root=0x" + root.ToInt64().ToString("X") +
                    ", anchor=0x" + choice.AnchorHwnd.ToInt64().ToString("X") +
                    " \"" + choice.AnchorClass + "\"" +
                    ", canvas=0x" + choice.CanvasHwnd.ToInt64().ToString("X") +
                    " \"" + choice.CanvasClass + "\"" +
                    ", rect=" + choice.AnchorRect.X + "," + choice.AnchorRect.Y + "," +
                    choice.AnchorRect.Width + "x" + choice.AnchorRect.Height +
                    ", reason=" + choice.Reason);
            }
            else if (!fullPageUi)
            {
                LoggingService.Warn("Locator: no content zone found for root=0x" + root.ToInt64().ToString("X") +
                    " (nodes=" + nodes.Count + ")");
            }
            return choice;
        }

        /// <summary>
        /// Word's full-page UI is up: the File menu (Backstage) and everything opened from
        /// it - Print, Save As, Account, Options. Measured on a live Word (field report,
        /// machine C, 2026-07-28): a visible window of class "FullpageUIHost" covers the
        /// whole client area, while the ENTIRE document subtree - _WwF, _WwB, _WwG, every
        /// MsoCommandBarDock and every MsoWorkPane - becomes hidden.
        ///
        /// There is no document area at such a moment, so "zone not found" is the correct
        /// answer and hiding the strip is the correct outcome (otherwise it would float
        /// over the File screen). What must NOT happen is treating it as a failure: it
        /// used to bump the failure counters, drop a window-tree dump into the log folder
        /// - which is our signal of a REAL breakage - and, with a single Word window open,
        /// five such syncs in a row would have switched the add-in to the CTP fallback.
        /// </summary>
        internal static bool IsFullPageUiActive(Rectangle clientScreen, IList<WindowSnapshot> nodes)
        {
            if (nodes == null) return false;
            if (clientScreen.Width <= 0 || clientScreen.Height <= 0) return false;

            long clientArea = Area(clientScreen);
            foreach (var n in nodes)
            {
                if (n == null || !n.Visible) continue;
                if (!string.Equals(n.ClassName, FullPageUiClassName, StringComparison.OrdinalIgnoreCase)) continue;
                // Only the part inside the client area counts: the window is a child of
                // OpusApp, so anything outside is not covering the document anyway.
                var covered = Rectangle.Intersect(n.ScreenRect, clientScreen);
                if (covered.Width <= 0 || covered.Height <= 0) continue;
                if (Area(covered) * 100 >= clientArea * FullPageUiCoverPercent) return true;
            }
            return false;
        }

        // ------------------------------------------------------------------
        // Pure selection logic (covered by unit tests)
        // ------------------------------------------------------------------

        /// <summary>
        /// The ribbon shadow window: class DropShadow, visible, thin, sitting right on top
        /// of the content zone. The strip is placed just below it in the z-order so that
        /// the Word shadow falls onto the strip (without it the strip looks glued onto the ribbon).
        /// </summary>
        internal static IntPtr FindShadowHwnd(IList<WindowSnapshot> nodes, int contentTopScreenY, int dpi)
        {
            if (nodes == null) return IntPtr.Zero;
            int maxH = NativeHostLayoutMath.Scale96(12, dpi) + 2;
            int slack = NativeHostLayoutMath.Scale96(3, dpi);
            foreach (var n in nodes)
            {
                if (n == null || !n.Visible) continue;
                if (!string.Equals(n.ClassName, "DropShadow", StringComparison.OrdinalIgnoreCase)) continue;
                var r = n.ScreenRect;
                if (r.Height <= 0 || r.Height > maxH) continue;
                if (Math.Abs(r.Top - contentTopScreenY) > slack) continue;
                return n.Hwnd;
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Chooses the document area from a window-tree snapshot.
        /// clientScreen is the client area of the root window in screen coordinates.
        /// </summary>
        internal static ZoneChoice FindContentZone(IntPtr root, Rectangle clientScreen,
            IList<WindowSnapshot> nodes, int dpi)
        {
            if (nodes == null || nodes.Count == 0) return null;
            if (clientScreen.Width <= 0 || clientScreen.Height <= 0) return null;

            int belowTop = clientScreen.Top + NativeHostLayoutMath.Scale96(BelowTopThreshold96, dpi);
            int minRemaining = NativeHostLayoutMath.Scale96(MinRemainingHeight96, dpi);

            var byHwnd = new Dictionary<IntPtr, WindowSnapshot>();
            foreach (var n in nodes)
            {
                if (n != null && n.Hwnd != IntPtr.Zero && !byHwnd.ContainsKey(n.Hwnd))
                    byHwnd[n.Hwnd] = n;
            }

            // --- 1. The main signal: the "_WwG" canvas ---
            WindowSnapshot canvas = null;
            foreach (var n in nodes)
            {
                if (n == null || !n.Visible) continue;
                if (!string.Equals(n.ClassName, CanvasClassName, StringComparison.OrdinalIgnoreCase)) continue;
                if (n.ScreenRect.Width <= 0 || n.ScreenRect.Height <= 0) continue;
                if (canvas == null || Area(n.ScreenRect) > Area(canvas.ScreenRect)) canvas = n;
            }

            if (canvas != null)
            {
                // Ancestor chain of the canvas: from the direct child of the root down to the canvas.
                var chain = AncestorChain(canvas, root, byHwnd);

                // The topmost container starting below the client top (= under the ribbon).
                // The chain ends with the canvas itself: if no container qualified, the
                // canvas becomes the anchor - a workable choice, but marked Degraded (the
                // host will later try to upgrade it).
                foreach (var candidate in chain)
                {
                    if (!candidate.Visible) continue;
                    if (candidate.ScreenRect.Top < belowTop) continue; // reaches the client top - contains the ribbon
                    if (!NativeHostLayoutMath.IsPlausibleContentTop(candidate.ScreenRect.Top,
                        clientScreen.Top, clientScreen.Bottom, minRemaining)) continue;

                    bool anchorIsCanvas = candidate.Hwnd == canvas.Hwnd;
                    return new ZoneChoice
                    {
                        AnchorHwnd = candidate.Hwnd,
                        AnchorClass = candidate.ClassName,
                        AnchorRect = candidate.ScreenRect,
                        CanvasHwnd = canvas.Hwnd,
                        CanvasClass = canvas.ClassName,
                        ShadowHwnd = FindShadowHwnd(nodes, candidate.ScreenRect.Top, dpi),
                        Reason = anchorIsCanvas
                            ? "canvas=" + CanvasClassName + " (no container qualified, anchor = canvas)"
                            : "canvas=" + CanvasClassName + ", anchor=" + candidate.ClassName +
                              " (topmost container under the ribbon)",
                        Degraded = anchorIsCanvas
                    };
                }

                // There is a canvas, but neither it nor its containers start below the top
                // of the client area. This is not "a different Word" but a WINDOW STILL
                // BEING BUILT: while Word lays out its interface the canvas temporarily
                // occupies the whole client area (measured on a real machine: _WwG =
                // 0,0,2534x1528 instead of 525,267,2035x1228). The generic fallback below
                // is meant for the _WwG class disappearing in some future Word and here it
                // would only pick a random window - so we honestly report "not found":
                // the search repeats after LocateRetryMs, once the layout exists.
                // Without this the strip landed at 0,0 - over the title bar and the ribbon.
                return null;
            }

            // --- 2. Fallback: the topmost suitable window below the client top ---
            WindowSnapshot best = null;
            int minWidth = (int)(clientScreen.Width * MinFallbackWidthShare);
            int minHeight = NativeHostLayoutMath.Scale96(100, dpi);
            foreach (var n in nodes)
            {
                if (n == null || !n.Visible) continue;
                var r = n.ScreenRect;
                if (r.Width < minWidth || r.Height < minHeight) continue;
                if (r.Top < belowTop) continue;
                if (!NativeHostLayoutMath.IsPlausibleContentTop(r.Top,
                    clientScreen.Top, clientScreen.Bottom, minRemaining)) continue;

                if (best == null
                    || r.Top < best.ScreenRect.Top
                    || (r.Top == best.ScreenRect.Top && Area(r) > Area(best.ScreenRect)))
                {
                    best = n;
                }
            }

            if (best != null)
            {
                return new ZoneChoice
                {
                    AnchorHwnd = best.Hwnd,
                    AnchorClass = best.ClassName,
                    AnchorRect = best.ScreenRect,
                    CanvasHwnd = best.Hwnd,
                    CanvasClass = best.ClassName,
                    ShadowHwnd = FindShadowHwnd(nodes, best.ScreenRect.Top, dpi),
                    Reason = "fallback: canvas " + CanvasClassName + " not found, took the topmost" +
                             " suitable window \"" + best.ClassName + "\"",
                    Degraded = true
                };
            }

            return null;
        }

        /// <summary>
        /// Ancestor chain of the canvas, from the direct child of root down to the canvas
        /// itself (inclusive). If no ancestors are in the snapshot, returns [canvas].
        /// </summary>
        private static List<WindowSnapshot> AncestorChain(WindowSnapshot canvas, IntPtr root,
            Dictionary<IntPtr, WindowSnapshot> byHwnd)
        {
            var chain = new List<WindowSnapshot> { canvas };
            var guard = new HashSet<IntPtr> { canvas.Hwnd };
            var current = canvas;
            while (current.Parent != IntPtr.Zero && current.Parent != root)
            {
                WindowSnapshot parent;
                if (!byHwnd.TryGetValue(current.Parent, out parent)) break;
                if (!guard.Add(parent.Hwnd)) break; // a cycle - stop
                chain.Add(parent);
                current = parent;
            }
            chain.Reverse(); // top down: [direct child of root, ..., canvas]
            return chain;
        }

        private static long Area(Rectangle r)
        {
            return (long)Math.Max(0, r.Width) * Math.Max(0, r.Height);
        }
    }
}
