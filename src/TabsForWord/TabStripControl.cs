using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TabsForWord
{
    /// <summary>
    /// The tab strip (a WinForms UserControl) per the "Tab specification 2b".
    /// The same control serves both hosts: the in-window Win32 host of the main mode
    /// (NativeTabHost) and the Custom Task Pane of the classic one. The CTP path is why
    /// the class is COM-visible and registered as a COM object - CreateCTP takes a ProgId.
    ///
    /// Drawing is entirely custom paint: the active tab with a single flared outline
    /// (R8 flowing into the bottom border of the strip), rounded inactive tabs, amber
    /// unsaved indicators, even shrinking on overflow -> a compact mode -> scrolling
    /// with the arrows, tab dragging, the all-tabs menu and collapsing.
    /// Every size scales with DPI (the specification is given in px at 100%).
    /// </summary>
    [ComVisible(true)]
    [Guid("D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8")]
    [ProgId("TabsForWord.TabStripControl")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TabStripControl : UserControl
    {
        // ------------------------------------------------------------------
        // Specification 2b metrics (px at 100% scale)
        // ------------------------------------------------------------------
        private const int PanelH = 37;         // 6 padding + 30 tabs + 1 border
        private const int TabTop = 6;
        private const int TabH = 30;
        private const int ActiveTabH = 31;     // overlaps the strip border by 1 px
        private const int PadLeft = 14;
        private const int PadRight = 12;
        private const int BtnSize = 26;        // the "+", all-tabs, collapse and scroll buttons
        private const int CloseBox = 18;
        private const int DotSize = 6;
        private const int PinIconSize = 12;    // the pin glyph on a pinned tab
        private const int SepH = 16;
        private const int GapNormal = 8;
        private const int GapCompact = 5;
        private const int TabPad = 12;
        private const int TabPadCompact = 9;
        private const int ActivePadRight = 7;
        private const int ContentGap = 7;      // between the text, the dot and the close button
        private const int ContentGapCompact = 5;
        private const int CornerR = 7;         // radius of the top tab corners
        private const int FlareR = 8;          // radius of the active-tab flare
        private const int TextMaxW = 170;      // the name is ellipsised beyond this
        private const int ActiveMinW = 180;
        private const int InactiveMinW = 80;
        private const int CollapsedH = 16;     // collapsed strip: a sliver with the expand arrow
        private const int ScrollStep = 60;
        private const int DragThreshold = 4;

        private List<DocumentTabModel> _tabs = new List<DocumentTabModel>();
        private int _scrollOffset;
        private bool _collapsed;
        private float _scale = 1f;
        private int _dpiOverride;          // 0 = use the control DeviceDpi
        private float _userScale = TabSizeSettings.Current;  // tab size (a multiplier on top of DPI)
        private Color? _panelBgOverride;   // in-window mode: strip background in the Word background colour
        private bool _activeFlushBottom;   // in-window mode: the active fill stops at the strip line

        // Fonts (recreated when the DPI changes)
        private Font _fontTab;          // Segoe UI 12.5 px
        private Font _fontTabActive;    // Segoe UI Semibold 12.5 px
        private Font _fontTabCompact;   // Segoe UI 12 px
        private readonly Dictionary<int, Font> _glyphFonts = new Dictionary<int, Font>();

        // ------------------------------------------------------------------
        // Layout: the result of laying out for the current width
        // ------------------------------------------------------------------
        private sealed class TabLayout
        {
            public DocumentTabModel Tab;
            public Rectangle Bounds;    // the tab body without the flares
        }

        private sealed class StripLayout
        {
            public List<TabLayout> Tabs = new List<TabLayout>();
            public Rectangle Plus, MenuBtn, CollapseBtn, Separator, ScrollLeft, ScrollRight;
            public Rectangle Viewport;  // the tab zone (for clipping and scrolling)
            public bool Overflow;       // scrolling mode
            public bool Compact;        // compact metrics for inactive tabs
            public int MaxScroll;
            public int? PinBoundaryX;   // border between pinned and ordinary tabs (null = one group or empty)
        }

        private enum HitKind { None, Tab, TabClose, Plus, Menu, Collapse, ScrollLeft, ScrollRight, Empty }

        private struct Hit
        {
            public HitKind Kind;
            public int Hwnd;
            public bool Same(Hit o) { return Kind == o.Kind && Hwnd == o.Hwnd; }
        }

        private Hit _hover;
        private Hit _pressed;
        private readonly ToolTip _toolTip = new ToolTip();
        private string _lastTooltip;

        // Dragging
        private bool _mouseDownLeft;
        private Point _mouseDownPt;
        private Hit _mouseDownHit;
        private bool _dragging;
        private int _dragHwnd;
        private int _dragGrabOffsetX;
        private int _dragMouseX;
        private List<int> _dragOrder;   // hwnd order for the duration of the drag

        // Keyboard focus
        private int _focusIndex = -1;

        public event Action<int> TabActivateRequested;
        public event Action<int> TabCloseRequested;
        public event Action NewDocumentRequested;
        public event Action<int> CloseOthersRequested;
        public event Action<int> OpenFolderRequested;
        public event Action<int, int> ReorderRequested;   // (hwnd, new index)
        public event Action CollapseToggleRequested;
        public event Action<int, Color?> TabColorChangeRequested; // (hwnd, colour; null = reset)
        public event Action<float> TabSizeChangeRequested;        // tab size multiplier
        public event Action<int, bool> TabPinChangeRequested;     // (hwnd, pin/unpin)
        public event Action<bool> HotkeyToggleRequested;          // settings window: Ctrl+Tab on/off
        public event Action<UiLang?> LanguageChangeRequested;      // settings window: language (null = "same as Word")

        public TabStripControl()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.OptimizedDoubleBuffer
                   | ControlStyles.UserPaint
                   | ControlStyles.ResizeRedraw
                   | ControlStyles.Selectable, true);
            TabStop = true;
            BackColor = TabTheme.PanelBg;
            _toolTip.InitialDelay = 400;
            _toolTip.ReshowDelay = 200;
            UpdateScale();
        }

        /// <summary>The content height the strip asks its host for (the in-window host or the CTP).</summary>
        public int DesiredContentHeight
        {
            get { return S(_collapsed ? CollapsedH : PanelH); }
        }

        public bool Collapsed
        {
            get { return _collapsed; }
        }

        /// <summary>
        /// Overrides the strip background (used by the in-window host: the background takes
        /// the colour of the Word work area so the tabs look as if they lie directly on it).
        /// null restores the specification 2b colour. Has no effect in CTP mode.
        /// </summary>
        public void SetPanelBackgroundOverride(Color? color)
        {
            if (_panelBgOverride == color) return;
            _panelBgOverride = color;
            BackColor = color.HasValue ? color.Value : TabTheme.PanelBg;
            Invalidate();
        }

        private Color PanelBgEffective
        {
            get { return _panelBgOverride.HasValue ? _panelBgOverride.Value : TabTheme.PanelBg; }
        }

        /// <summary>
        /// In CTP mode the fill of the active tab flows through the bottom border of the
        /// strip and merges with the white area below it (specification 2b).
        /// In the in-window mode the Word document is right below, so the fill has to stop
        /// at the line, otherwise a 1 px white lip shows underneath.
        /// </summary>
        public void SetActiveTabFlushBottom(bool value)
        {
            if (_activeFlushBottom == value) return;
            _activeFlushBottom = value;
            Invalidate();
        }

        /// <summary>Collapse or expand the strip (the CTP height is changed by TabPaneManager).</summary>
        public void SetCollapsed(bool collapsed)
        {
            if (_collapsed == collapsed) return;
            _collapsed = collapsed;
            _hover = new Hit { Kind = HitKind.None };
            CancelDrag();
            Invalidate();
        }

        /// <summary>Updates the tab list. No change means no repaint (the UI does not flicker).</summary>
        public void UpdateTabs(IList<DocumentTabModel> tabs)
        {
            try
            {
                var copy = new List<DocumentTabModel>(tabs);
                if (DocumentTabModel.ListsEqual(_tabs, copy)) return;

                // The set of tabs changed during a drag - cancel the drag.
                if (_dragging)
                {
                    var oldSet = new HashSet<int>(_tabs.Select(t => t.Hwnd));
                    var newSet = new HashSet<int>(copy.Select(t => t.Hwnd));
                    if (!oldSet.SetEquals(newSet)) CancelDrag();
                }

                _tabs = copy;
                if (_focusIndex >= _tabs.Count) _focusIndex = _tabs.Count - 1;
                EnsureActiveVisible();
                Invalidate();
            }
            catch (Exception ex)
            {
                LoggingService.Error("UpdateTabs failed", ex);
            }
        }

        // ------------------------------------------------------------------
        // Scale and fonts
        // ------------------------------------------------------------------

        private int S(float v) { return (int)Math.Round(v * _scale); }
        private float SF(float v) { return v * _scale; }
        private float PenW { get { return Math.Max(1f, (float)Math.Round(_scale)); } }

        /// <summary>
        /// An explicit DPI from the strip owner (both the in-window host and the CTP pass
        /// the DPI of the WORD WINDOW - GetDpiForWindow). That is the only reliable source:
        /// the DeviceDpi of a WinForms control inside a foreign per-monitor Word window is
        /// not in sync with the monitor and drifts on multi-DPI setups (verified on a real
        /// Word: the control on a 150% monitor thought it was 96, and after moving to a
        /// 100% monitor thought it was 144). 0 goes back to DeviceDpi.
        /// </summary>
        public void SetDpiOverride(int dpi)
        {
            if (dpi != 0 && (dpi < 48 || dpi > 768)) return; // ignore junk values
            if (_dpiOverride == dpi) return;
            _dpiOverride = dpi;
            LoggingService.Info("Tab strip DPI override: " + (dpi == 0 ? "off" : dpi + " (" +
                (int)Math.Round(dpi * 100.0 / 96) + "%)"));
            try { UpdateScale(); }
            catch (Exception ex) { LoggingService.Error("SetDpiOverride failed", ex); }
        }

        /// <summary>The DPI currently in effect for the strip (for conversions in the host).</summary>
        public int EffectiveDpi
        {
            get
            {
                if (_dpiOverride > 0) return _dpiOverride;
                try { return DeviceDpi; } catch { return 96; }
            }
        }

        /// <summary>
        /// The user tab size: a multiplier on top of the DPI scale.
        /// Saving it and pushing it to every window is the strip owner job.
        /// </summary>
        public void SetUserScale(float scale)
        {
            float v = TabSizeSettings.Clamp(scale);
            if (Math.Abs(v - _userScale) < 0.001f) return;
            _userScale = v;
            try { UpdateScale(); }
            catch (Exception ex) { LoggingService.Error("SetUserScale failed", ex); }
        }

        public float UserScale
        {
            get { return _userScale; }
        }

        private void UpdateScale()
        {
            float s = 1f;
            try { s = EffectiveDpi / 96f * _userScale; } catch { }
            if (s < 0.5f || s > 8f) s = 1f;
            _scale = s;

            DisposeFonts();
            _fontTab = new Font("Segoe UI", SF(12.5f), FontStyle.Regular, GraphicsUnit.Pixel);
            _fontTabCompact = new Font("Segoe UI", SF(12f), FontStyle.Regular, GraphicsUnit.Pixel);
            _fontTabActive = CreateSemibold(12.5f);
            Font = _fontTab;
            Invalidate();
        }

        private Font CreateSemibold(float px)
        {
            try
            {
                var f = new Font("Segoe UI Semibold", SF(px), FontStyle.Regular, GraphicsUnit.Pixel);
                if (string.Equals(f.Name, "Segoe UI Semibold", StringComparison.OrdinalIgnoreCase)) return f;
                f.Dispose();
            }
            catch { }
            return new Font("Segoe UI", SF(px), FontStyle.Bold, GraphicsUnit.Pixel);
        }

        private Font GlyphFont(int px)
        {
            Font f;
            int key = S(px);
            if (_glyphFonts.TryGetValue(key, out f)) return f;
            f = new Font("Segoe UI", key, FontStyle.Regular, GraphicsUnit.Pixel);
            _glyphFonts[key] = f;
            return f;
        }

        private void DisposeFonts()
        {
            if (_fontTab != null) { _fontTab.Dispose(); _fontTab = null; }
            if (_fontTabActive != null) { _fontTabActive.Dispose(); _fontTabActive = null; }
            if (_fontTabCompact != null) { _fontTabCompact.Dispose(); _fontTabCompact = null; }
            foreach (var f in _glyphFonts.Values) f.Dispose();
            _glyphFonts.Clear();
        }

        protected override void OnDpiChangedAfterParent(EventArgs e)
        {
            base.OnDpiChangedAfterParent(e);
            try { UpdateScale(); } catch (Exception ex) { LoggingService.Error("DPI change failed", ex); }
        }

        // ------------------------------------------------------------------
        // Measurement and layout
        // ------------------------------------------------------------------

        private const TextFormatFlags MeasureFlags =
            TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding;

        private const TextFormatFlags DrawFlags = MeasureFlags |
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis;

        private Font TabFont(DocumentTabModel t, bool compact)
        {
            if (t.IsActive) return _fontTabActive;
            return compact ? _fontTabCompact : _fontTab;
        }

        private int MeasureCaption(Graphics g, DocumentTabModel t, bool compact)
        {
            int w = TextRenderer.MeasureText(g, t.Caption ?? string.Empty, TabFont(t, compact),
                new Size(int.MaxValue, int.MaxValue), MeasureFlags).Width;
            return Math.Min(w, S(TextMaxW));
        }

        /// <summary>Natural tab width (before any shrinking).</summary>
        private int NaturalWidth(Graphics g, DocumentTabModel t, bool compact)
        {
            int text = MeasureCaption(g, t, compact);
            if (t.IsActive)
            {
                // 12 + [pin 12 + 7] + text + 7 + [dot 6 + 7] + close 18 + 7
                int pinActive = t.IsPinned ? S(PinIconSize) + S(ContentGap) : 0;
                return S(TabPad) + pinActive + text + S(ContentGap)
                     + (t.Saved ? 0 : S(DotSize) + S(ContentGap))
                     + S(CloseBox) + S(ActivePadRight);
            }
            int pad = compact ? S(TabPadCompact) : S(TabPad);
            int gap = compact ? S(ContentGapCompact) : S(ContentGap);
            int pin = t.IsPinned ? S(PinIconSize) + gap : 0;
            return pad + pin + text + (t.Saved ? 0 : gap + S(DotSize)) + pad;
        }

        private StripLayout ComputeLayout(Graphics g, IList<DocumentTabModel> tabs)
        {
            var l = new StripLayout();
            int btn = S(BtnSize);
            int btnY = S(TabTop);

        // The right-hand block: ... [separator] 4 [all tabs] 2 [collapse] 12
            int x = Width - S(PadRight);
            l.CollapseBtn = new Rectangle(x - btn, btnY, btn, btn);
            x -= btn + S(2);
            l.MenuBtn = new Rectangle(x - btn, btnY, btn, btn);
            x -= btn + S(4);
            l.Separator = new Rectangle(x - S(1), btnY + (btn - S(SepH)) / 2, Math.Max(1, S(1)), S(SepH));
            x -= S(1) + S(GapNormal);

            int zoneLeft = S(PadLeft);
            int zoneRight = Math.Max(zoneLeft + btn, x);
            l.Viewport = new Rectangle(zoneLeft, 0, zoneRight - zoneLeft, S(PanelH));

            int n = tabs.Count;
            if (n == 0)
            {
                l.Plus = new Rectangle(zoneLeft, btnY, btn, btn);
                return l;
            }

            // 1) Natural widths with the ordinary metrics
            var natural = new int[n];
            long total = 0;
            for (int i = 0; i < n; i++)
            {
                natural[i] = NaturalWidth(g, tabs[i], false);
                total += natural[i];
            }
            int avail = zoneRight - zoneLeft;
            // (n-1) gaps between tabs + the gap before "+" + "+" itself
            long fixedNormal = (long)S(GapNormal) * n + btn;

            if (total + fixedNormal <= avail)
            {
                PlaceTabs(l, tabs, natural, zoneLeft, S(GapNormal), false);
                return l;
            }

            // 2) Compact metrics plus even shrinking
            l.Compact = true;
            var naturalC = new int[n];
            var floor = new int[n];
            int maxNat = 0;
            for (int i = 0; i < n; i++)
            {
                naturalC[i] = NaturalWidth(g, tabs[i], !tabs[i].IsActive);
                floor[i] = Math.Min(naturalC[i], tabs[i].IsActive ? S(ActiveMinW) : S(InactiveMinW));
                if (naturalC[i] > maxNat) maxNat = naturalC[i];
            }
            int gapC = S(GapCompact);
            long fixedCompact = (long)gapC * n + btn;
            long minTotal = 0;
            for (int i = 0; i < n; i++) minTotal += floor[i];

            if (minTotal + fixedCompact <= avail)
            {
                // Binary search for a common width cap: the widest tabs shrink.
                int lo = 0, hi = maxNat;
                while (lo < hi)
                {
                    int mid = (lo + hi + 1) / 2;
                    long sum = 0;
                    for (int i = 0; i < n; i++) sum += ClampWidth(naturalC[i], floor[i], mid);
                    if (sum + fixedCompact <= avail) lo = mid; else hi = mid - 1;
                }
                var widths = new int[n];
                long used = 0;
                for (int i = 0; i < n; i++) { widths[i] = ClampWidth(naturalC[i], floor[i], lo); used += widths[i]; }
                // The leftover pixels are handed to the shrunk tabs from left to right.
                long spare = avail - fixedCompact - used;
                for (int i = 0; i < n && spare > 0; i++)
                {
                    if (widths[i] < naturalC[i]) { widths[i]++; spare--; }
                }
                PlaceTabs(l, tabs, widths, zoneLeft, gapC, false);
                return l;
            }

            // 3) It does not fit even at the minimum - switch to scrolling
            l.Overflow = true;
            l.ScrollLeft = new Rectangle(zoneLeft, btnY, btn, btn);
            l.Plus = new Rectangle(zoneRight - btn, btnY, btn, btn);
            l.ScrollRight = new Rectangle(l.Plus.X - gapC - btn, btnY, btn, btn);
            int vpLeft = l.ScrollLeft.Right + gapC;
            int vpRight = l.ScrollRight.X - gapC;
            l.Viewport = new Rectangle(vpLeft, 0, Math.Max(S(InactiveMinW), vpRight - vpLeft), S(PanelH));

            long totalTabs = 0;
            for (int i = 0; i < n; i++) totalTabs += floor[i];
            totalTabs += (long)gapC * (n - 1);
            l.MaxScroll = (int)Math.Max(0, totalTabs - l.Viewport.Width);
            if (_scrollOffset > l.MaxScroll) _scrollOffset = l.MaxScroll;
            if (_scrollOffset < 0) _scrollOffset = 0;

            int tx = l.Viewport.X - _scrollOffset;
            for (int i = 0; i < n; i++)
            {
                int h = tabs[i].IsActive ? S(ActiveTabH) : S(TabH);
                l.Tabs.Add(new TabLayout { Tab = tabs[i], Bounds = new Rectangle(tx, S(TabTop), floor[i], h) });
                tx += floor[i] + gapC;
            }
            ComputePinBoundary(l);
            return l;
        }

        private static int ClampWidth(int natural, int floor, int cap)
        {
            return Math.Min(natural, Math.Max(floor, cap));
        }

        private void PlaceTabs(StripLayout l, IList<DocumentTabModel> tabs, int[] widths, int startX, int gap, bool overflow)
        {
            int x = startX;
            for (int i = 0; i < tabs.Count; i++)
            {
                int h = tabs[i].IsActive ? S(ActiveTabH) : S(TabH);
                l.Tabs.Add(new TabLayout { Tab = tabs[i], Bounds = new Rectangle(x, S(TabTop), widths[i], h) });
                x += widths[i] + gap;
            }
            l.Plus = new Rectangle(x, S(TabTop), S(BtnSize), S(BtnSize));
            ComputePinBoundary(l);
        }

        /// <summary>
        /// The border between the pinned group and the ordinary tabs (for the thin
        /// separator between them) - where a pinned tab is followed by an unpinned one in
        /// the displayed order. That order guarantees there is at most one such border
        /// (see DocumentWindowManager: pinned tabs always form the first group).
        /// </summary>
        private static void ComputePinBoundary(StripLayout l)
        {
            for (int i = 1; i < l.Tabs.Count; i++)
            {
                if (l.Tabs[i - 1].Tab.IsPinned && !l.Tabs[i].Tab.IsPinned)
                {
                    l.PinBoundaryX = (l.Tabs[i - 1].Bounds.Right + l.Tabs[i].Bounds.Left) / 2;
                    return;
                }
            }
        }

        /// <summary>
        /// Rectangles of the tab content: the pin of a pinned tab, the text, the close
        /// button and the unsaved dot. showClose: always for the active tab; for an
        /// inactive one on hover (except in the compact mode).
        /// </summary>
        private void GetTabContent(Graphics g, DocumentTabModel t, Rectangle b, bool compact, bool showClose,
                                   out Rectangle textRect, out Rectangle closeRect, out Rectangle dotRect,
                                   out Rectangle pinRect)
        {
            closeRect = Rectangle.Empty;
            dotRect = Rectangle.Empty;
            pinRect = Rectangle.Empty;
            bool tabCompact = compact && !t.IsActive;
            int gap = tabCompact ? S(ContentGapCompact) : S(ContentGap);
            int padL = tabCompact ? S(TabPadCompact) : S(TabPad);
            int x = b.X + padL;
            int right;

            if (t.IsPinned)
            {
                int p = S(PinIconSize);
                pinRect = new Rectangle(x, b.Y + (b.Height - p) / 2, p, p);
                x = pinRect.Right + gap;
            }

            if (showClose)
            {
                int cs = S(CloseBox);
                closeRect = new Rectangle(b.Right - S(ActivePadRight) - cs,
                                          b.Y + (b.Height - cs) / 2, cs, cs);
                right = closeRect.X - gap;
                if (!t.Saved)
                {
                    int d = S(DotSize);
                    // the dot sits left of the close button (specification section 4)
                    dotRect = new Rectangle(right - d, b.Y + (b.Height - d) / 2, d, d);
                    right = dotRect.X - gap;
                }
                int availW = Math.Max(0, right - x);
                textRect = new Rectangle(x, b.Y, Math.Min(availW, S(TextMaxW)), b.Height);
            }
            else
            {
                right = b.Right - padL;
                int reserve = t.Saved ? 0 : gap + S(DotSize);
                int availW = Math.Max(0, right - x - reserve);
                int textW = Math.Min(Math.Min(MeasureCaption(g, t, compact), availW), S(TextMaxW));
                textRect = new Rectangle(x, b.Y, textW, b.Height);
                if (!t.Saved)
                {
                    int d = S(DotSize);
                    // the dot sits 7 px after the text (specification section 4)
                    dotRect = new Rectangle(x + textW + gap, b.Y + (b.Height - d) / 2, d, d);
                }
            }
        }

        /// <summary>
        /// Test-only: [visible zone, left arrow, right arrow] in scrolling mode;
        /// an empty array means there is no overflow.
        /// </summary>
        internal Rectangle[] OverflowGeometryForTests()
        {
            using (var g = CreateGraphics())
            {
                var lay = ComputeLayout(g, _tabs);
                return lay.Overflow
                    ? new[] { lay.Viewport, lay.ScrollLeft, lay.ScrollRight }
                    : new Rectangle[0];
            }
        }

        private void EnsureActiveVisible()
        {
            try
            {
                if (!IsHandleCreated || _collapsed) return;
                using (var g = CreateGraphics())
                {
                    var lay = ComputeLayout(g, _tabs);
                    if (!lay.Overflow) { _scrollOffset = 0; return; }
                    for (int i = 0; i < lay.Tabs.Count; i++)
                    {
                        if (!lay.Tabs[i].Tab.IsActive) continue;
                        var b = lay.Tabs[i].Bounds;
                        if (b.Left < lay.Viewport.Left) _scrollOffset -= lay.Viewport.Left - b.Left;
                        else if (b.Right > lay.Viewport.Right) _scrollOffset += b.Right - lay.Viewport.Right;
                        if (_scrollOffset < 0) _scrollOffset = 0;
                        if (_scrollOffset > lay.MaxScroll) _scrollOffset = lay.MaxScroll;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("EnsureActiveVisible failed", ex);
            }
        }

        /// <summary>Tabs in displayed order (during a drag, in the temporary order).</summary>
        private List<DocumentTabModel> DisplayTabs()
        {
            if (!_dragging || _dragOrder == null) return _tabs;
            var byHwnd = new Dictionary<int, DocumentTabModel>();
            foreach (var t in _tabs) byHwnd[t.Hwnd] = t;
            var result = new List<DocumentTabModel>(_tabs.Count);
            foreach (var h in _dragOrder)
            {
                DocumentTabModel t;
                if (byHwnd.TryGetValue(h, out t)) result.Add(t);
            }
            foreach (var t in _tabs)
            {
                if (!_dragOrder.Contains(t.Hwnd)) result.Add(t);
            }
            return result;
        }

        // ------------------------------------------------------------------
        // Outline geometry
        // ------------------------------------------------------------------

        private static void AddArcC(GraphicsPath p, float cx, float cy, float r, float start, float sweep)
        {
            p.AddArc(cx - r, cy - r, 2f * r, 2f * r, start, sweep);
        }

        /// <summary>Closed outline of an inactive tab (only the top corners are rounded).</summary>
        private static GraphicsPath RoundedTopPath(RectangleF b, float r)
        {
            var p = new GraphicsPath();
            p.StartFigure();
            p.AddLine(b.Left, b.Bottom, b.Left, b.Top + r);
            AddArcC(p, b.Left + r, b.Top + r, r, 180f, 90f);
            p.AddLine(b.Left + r, b.Top, b.Right - r, b.Top);
            AddArcC(p, b.Right - r, b.Top + r, r, 270f, 90f);
            p.AddLine(b.Right, b.Top + r, b.Right, b.Bottom);
            p.CloseFigure();
            return p;
        }

        /// <summary>Open border stroke of an inactive tab: the sides and the top, no bottom.</summary>
        private static GraphicsPath RoundedTopStroke(RectangleF b, float r, float half)
        {
            float l = b.Left + half, rt = b.Right - half, t = b.Top + half;
            float rr = Math.Max(1f, r - half);
            var p = new GraphicsPath();
            p.StartFigure();
            p.AddLine(l, b.Bottom, l, t + rr);
            AddArcC(p, l + rr, t + rr, rr, 180f, 90f);
            p.AddLine(l + rr, t, rt - rr, t);
            AddArcC(p, rt - rr, t + rr, rr, 270f, 90f);
            p.AddLine(rt, t + rr, rt, b.Bottom);
            return p;
        }

        /// <summary>
        /// Fill of the active tab: the body plus the R8 flares descending into the bottom
        /// border of the strip on both sides (specification section 5). bx0/bx1 are the
        /// edges of the tab body. With _activeFlushBottom the fill bottom is raised by the
        /// thickness of the strip line.
        private GraphicsPath ActiveFillPath(float bx0, float bx1)
        {
            float top = SF(TabTop), bottom = SF(PanelH) - (_activeFlushBottom ? PenW : 0f);
            float r7 = SF(CornerR), r8 = SF(FlareR);
            var p = new GraphicsPath();
            p.StartFigure();
            p.AddLine(bx0 - r8, bottom, bx0 - r8, bottom - 0.01f);
            AddArcC(p, bx0 - r8, bottom - r8, r8, 90f, -90f);
            p.AddLine(bx0, bottom - r8, bx0, top + r7);
            AddArcC(p, bx0 + r7, top + r7, r7, 180f, 90f);
            p.AddLine(bx0 + r7, top, bx1 - r7, top);
            AddArcC(p, bx1 - r7, top + r7, r7, 270f, 90f);
            p.AddLine(bx1, top + r7, bx1, bottom - r8);
            AddArcC(p, bx1 + r8, bottom - r8, r8, 180f, -90f);
            p.CloseFigure();
            return p;
        }

        /// <summary>
        /// Open outline of the active tab - one continuous path from the bottom border of
        /// the strip on the left to the same border on the right (no bottom side). The
        /// half-stroke inset matches the reference SVG (section 5).
        /// </summary>
        private GraphicsPath ActiveStrokePath(float bx0, float bx1)
        {
            float half = PenW / 2f;
            float top = SF(TabTop) + half;
            float bottom = SF(PanelH) - half;           // middle of the 1 px strip border
            float r7 = Math.Max(1f, SF(CornerR) - half);
            float r8 = SF(FlareR);
            float l = bx0 + half, rt = bx1 - half;
            var p = new GraphicsPath();
            p.StartFigure();
            // the left flare: from the border line up into the side of the tab
            AddArcC(p, bx0 - r8 + half, bottom - r8, r8, 90f, -90f);
            p.AddLine(l, bottom - r8, l, top + r7);
            AddArcC(p, l + r7, top + r7, r7, 180f, 90f);
            p.AddLine(l + r7, top, rt - r7, top);
            AddArcC(p, rt - r7, top + r7, r7, 270f, 90f);
            p.AddLine(rt, top + r7, rt, bottom - r8);
            // the right flare: back down into the border line
            AddArcC(p, bx1 + r8 - half, bottom - r8, r8, 180f, -90f);
            return p;
        }

        // ------------------------------------------------------------------
        // Paint
        // ------------------------------------------------------------------

        protected override void OnPaint(PaintEventArgs e)
        {
            try
            {
                base.OnPaint(e);
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.None;

                if (_collapsed)
                {
                    DrawCollapsed(g);
                    return;
                }

                int h = S(PanelH);
                using (var bg = new SolidBrush(PanelBgEffective))
                    g.FillRectangle(bg, 0, 0, Width, h);
                if (ClientSize.Height > h)
                {
                    using (var below = new SolidBrush(TabTheme.BelowPanelBg))
                        g.FillRectangle(below, 0, h, Width, ClientSize.Height - h);
                }

                var disp = DisplayTabs();
                var lay = ComputeLayout(g, disp);

                // Bottom border of the strip (the active tab paints over its own stretch)
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(TabTheme.PanelBorder, PenW))
                    g.DrawLine(pen, 0, h - PenW / 2f, Width, h - PenW / 2f);
                g.SmoothingMode = SmoothingMode.None;

                // Tabs: the inactive ones first, the active one on top (its flares reach into the gaps).
                // In scrolling mode the tabs are drawn into a SEPARATE BUFFER the size of the
                // visible zone and blitted in: Graphics.Clip is unreliable - TextRenderer (GDI)
                // ignores the GDI+ clip, and tab captions
                // would "escape" the zone onto the arrows and the buttons (verified live:
                // fills were clipped, text was not). A buffer clips everything for certain.
                if (lay.Overflow)
                {
                    var vp = new Rectangle(lay.Viewport.X, 0, lay.Viewport.Width, S(PanelH));
                    if (vp.Width > 0 && vp.Height > 0)
                    {
                        // The buffer is OPAQUE (32bppRgb) and pre-filled with the strip
                        // background plus the segment of the bottom line: on a transparent
                        // ARGB surface GDI text is drawn without anti-aliasing - "fat and
                        // jagged". Blitting an opaque buffer is seamless: its background is
                        // the same as the strip around it.
                        using (var buf = new Bitmap(vp.Width, vp.Height, PixelFormat.Format32bppRgb))
                        {
                            using (var bg2 = Graphics.FromImage(buf))
                            {
                                using (var bgBrush = new SolidBrush(PanelBgEffective))
                                    bg2.FillRectangle(bgBrush, 0, 0, vp.Width, vp.Height);
                                bg2.SmoothingMode = SmoothingMode.AntiAlias;
                                using (var pen = new Pen(TabTheme.PanelBorder, PenW))
                                    bg2.DrawLine(pen, 0, S(PanelH) - PenW / 2f, vp.Width, S(PanelH) - PenW / 2f);
                                bg2.SmoothingMode = SmoothingMode.None;

                                // NOT TranslateTransform: TextRenderer (GDI) ignores a Graphics
                                // transform exactly as it ignores the clip - the tab rectangles
                                // themselves are shifted (dx), so text and fill cannot drift
                                // apart. Captions are NOT drawn into the buffer: GDI text on a
                                // bitmap comes out unsmoothed ("fat and jagged") - they go in a
                                // second pass straight onto the screen.
                                DrawTabsRow(bg2, lay, -vp.X, false);
                                if (lay.PinBoundaryX.HasValue)
                                    DrawPinBoundary(bg2, lay.PinBoundaryX.Value - vp.X);
                            }
                            g.DrawImage(buf, vp.Location);
                        }

                        // Captions go onto the screen DC (live ClearType); they cannot escape
                        // the zone because the text rectangle is trimmed to it, and DrawText
                        // respects the bounds of its rectangle by itself.
                        DrawTabsRowCaptions(g, lay, vp);
                    }
                }
                else
                {
                    DrawTabsRow(g, lay, 0, true);
                    if (lay.PinBoundaryX.HasValue) DrawPinBoundary(g, lay.PinBoundaryX.Value);
                }

                // Scroll buttons
                if (lay.Overflow)
                {
                    DrawIconButton(g, lay.ScrollLeft, "‹", 15, _scrollOffset > 0,
                        HitIs(_hover, HitKind.ScrollLeft), HitIs(_pressed, HitKind.ScrollLeft));
                    DrawIconButton(g, lay.ScrollRight, "›", 15, _scrollOffset < lay.MaxScroll,
                        HitIs(_hover, HitKind.ScrollRight), HitIs(_pressed, HitKind.ScrollRight));
                }

                // "+", separator, all-tabs, collapse
                DrawIconButton(g, lay.Plus, "+", 15, true, HitIs(_hover, HitKind.Plus), HitIs(_pressed, HitKind.Plus));
                using (var sep = new SolidBrush(TabTheme.Separator))
                    g.FillRectangle(sep, lay.Separator);
                DrawIconButton(g, lay.MenuBtn, "▾", 10, true, HitIs(_hover, HitKind.Menu), HitIs(_pressed, HitKind.Menu));
                DrawIconButton(g, lay.CollapseBtn, "⌄", 12, true, HitIs(_hover, HitKind.Collapse), HitIs(_pressed, HitKind.Collapse));

                // Keyboard focus frame (section 9); nothing is drawn outside the visible zone
                if (Focused && _focusIndex >= 0 && _focusIndex < lay.Tabs.Count)
                {
                    var fb = lay.Tabs[_focusIndex].Bounds;
                    if (!lay.Overflow ||
                        (fb.Left >= lay.Viewport.Left - S(2) && fb.Right <= lay.Viewport.Right + S(2)))
                        DrawFocusFrame(g, fb);
                }

                // The dragged tab goes on top of everything
                if (_dragging)
                    DrawDraggedTab(g, lay);
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnPaint failed", ex);
            }
        }

        private static bool HitIs(Hit h, HitKind kind) { return h.Kind == kind; }

        /// <summary>
        /// The row of tabs (without the buttons): the inactive ones, then the active one on top.
        /// dx is the horizontal shift of the rectangles (for drawing into the scroll-zone
        /// buffer); skips outside the zone are computed from the original bounds.
        /// </summary>
        private void DrawTabsRow(Graphics g, StripLayout lay, int dx, bool withText)
        {
            TabLayout active = null;
            foreach (var tl in lay.Tabs)
            {
                if (_dragging && tl.Tab.Hwnd == _dragHwnd) continue; // drawn on top, floating
                if (tl.Tab.IsActive) { active = tl; continue; }
                if (tl.Bounds.Right < lay.Viewport.Left - S(FlareR) ||
                    tl.Bounds.Left > lay.Viewport.Right + S(FlareR)) continue;
                var b = tl.Bounds;
                b.Offset(dx, 0);
                DrawInactiveTab(g, tl.Tab, b, lay.Compact, withText);
            }
            if (active != null &&
                active.Bounds.Right >= lay.Viewport.Left - S(FlareR) &&
                active.Bounds.Left <= lay.Viewport.Right + S(FlareR))
            {
                var b = active.Bounds;
                b.Offset(dx, 0);
                DrawActiveTab(g, active.Tab, b, withText);
            }
        }

        /// <summary>Thin separator between the pinned group and the ordinary tabs.</summary>
        private void DrawPinBoundary(Graphics g, int x)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(TabTheme.Separator, PenW))
                g.DrawLine(pen, x, S(TabTop) + S(4), x, S(TabTop) + S(TabH) - S(4));
            g.SmoothingMode = SmoothingMode.None;
        }

        /// <summary>Caption colour of a tab (one rule for every drawing pass).</summary>
        private static Color CaptionColor(DocumentTabModel t)
        {
            if (t.IsProtectedView) return TabTheme.ProtectedViewText;
            return t.IsActive ? TabTheme.ActiveText : TabTheme.InactiveText;
        }

        /// <summary>
        /// The second pass in scrolling mode: captions only, straight onto the screen DC.
        /// The text rectangle is trimmed to the visible zone - DrawText will not go past
        /// it, and the caption of a partly visible tab gets an ellipsis.
        /// </summary>
        private void DrawTabsRowCaptions(Graphics g, StripLayout lay, Rectangle clip)
        {
            foreach (var tl in lay.Tabs)
            {
                if (_dragging && tl.Tab.Hwnd == _dragHwnd) continue;
                if (tl.Bounds.Right < clip.Left || tl.Bounds.Left > clip.Right) continue;
                var t = tl.Tab;
                bool hovered = (_hover.Kind == HitKind.Tab || _hover.Kind == HitKind.TabClose) && _hover.Hwnd == t.Hwnd;
                bool showClose = t.IsActive || (hovered && !lay.Compact && !_dragging);
                DrawTabCaption(g, t, tl.Bounds, lay.Compact, showClose, CaptionColor(t), clip);
            }
        }

        private void DrawTabCaption(Graphics g, DocumentTabModel t, Rectangle b, bool compact,
                                    bool showClose, Color color, Rectangle clip)
        {
            Rectangle textRect, closeRect, dotRect, pinRect;
            GetTabContent(g, t, b, compact, showClose, out textRect, out closeRect, out dotRect, out pinRect);
            textRect.Intersect(clip);
            if (textRect.Width > 0)
                TextRenderer.DrawText(g, t.Caption ?? string.Empty, TabFont(t, compact), textRect, color, DrawFlags);
        }

        private void DrawCollapsed(Graphics g)
        {
            int h = S(CollapsedH);
            using (var bg = new SolidBrush(PanelBgEffective))
                g.FillRectangle(bg, 0, 0, Width, h);
            if (ClientSize.Height > h)
            {
                using (var below = new SolidBrush(TabTheme.BelowPanelBg))
                    g.FillRectangle(below, 0, h, Width, ClientSize.Height - h);
            }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(TabTheme.PanelBorder, PenW))
                g.DrawLine(pen, 0, h - PenW / 2f, Width, h - PenW / 2f);
            g.SmoothingMode = SmoothingMode.None;

            var btn = CollapsedExpandRect();
            DrawIconButton(g, btn, "⌃", 11, true, HitIs(_hover, HitKind.Collapse), HitIs(_pressed, HitKind.Collapse));
        }

        private Rectangle CollapsedExpandRect()
        {
            int h = S(CollapsedH);
            int bw = S(BtnSize), bh = Math.Max(S(12), h - S(2));
            return new Rectangle(Width - S(PadRight) - bw, Math.Max(0, (h - S(1) - bh) / 2), bw, bh);
        }

        private void DrawInactiveTab(Graphics g, DocumentTabModel t, Rectangle b, bool compact, bool withText)
        {
            bool hovered = (_hover.Kind == HitKind.Tab || _hover.Kind == HitKind.TabClose) && _hover.Hwnd == t.Hwnd;
            bool showClose = hovered && !compact && !_dragging;

            var rb = new RectangleF(b.X, b.Y, b.Width, b.Height);
            float r = SF(CornerR);

            // User colour: a pastel fill and a tinted border.
            Color fillColor = hovered ? TabTheme.InactiveHoverBg : TabTheme.InactiveBg;
            Color borderColor = TabTheme.InactiveBorder;
            if (t.TabColor.HasValue)
            {
                fillColor = hovered
                    ? TabTheme.TintInactiveHoverBg(t.TabColor.Value)
                    : TabTheme.TintInactiveBg(t.TabColor.Value);
                borderColor = TabTheme.TintBorder(t.TabColor.Value);
            }

            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = RoundedTopPath(rb, r))
            {
                using (var fill = new SolidBrush(fillColor))
                    g.FillPath(fill, path);

                // Unsaved edge: 3 px across the full width on top of the upper border (section 4)
                if (!t.Saved)
                {
                    var oldClip = g.Clip;
                    g.SetClip(path, CombineMode.Intersect);
                    using (var edge = new SolidBrush(TabTheme.UnsavedEdge))
                        g.FillRectangle(edge, b.X, b.Y, b.Width, S(3));
                    g.Clip = oldClip;
                }

                using (var stroke = RoundedTopStroke(rb, r, PenW / 2f))
                using (var pen = new Pen(borderColor, PenW))
                    g.DrawPath(pen, stroke);
            }
            g.SmoothingMode = SmoothingMode.None;

            DrawTabContent(g, t, b, compact, showClose,
                t.IsProtectedView ? TabTheme.ProtectedViewText : TabTheme.InactiveText, withText);
        }

        /// <summary>Colours of the active tab, taking the user colour into account.</summary>
        private static void ActiveTabColors(DocumentTabModel t, out Color fill, out Color stroke)
        {
            if (t.TabColor.HasValue)
            {
                fill = TabTheme.TintActiveBg(t.TabColor.Value);
                stroke = TabTheme.TintBorder(t.TabColor.Value);
            }
            else
            {
                fill = TabTheme.ActiveBg;
                stroke = TabTheme.PanelBorder;
            }
        }

        private void DrawActiveTab(Graphics g, DocumentTabModel t, Rectangle b, bool withText)
        {
            Color fillColor, strokeColor;
            ActiveTabColors(t, out fillColor, out strokeColor);

            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var fill = ActiveFillPath(b.X, b.Right))
            using (var brush = new SolidBrush(fillColor))
                g.FillPath(brush, fill);
            using (var stroke = ActiveStrokePath(b.X, b.Right))
            using (var pen = new Pen(strokeColor, PenW))
                g.DrawPath(pen, stroke);

            // In-window mode: the anti-aliased edge of the fill mixes white into the row of
            // the bottom line - the segment is redrawn on top so that the line stays
            // continuous and equally dense under the tab and next to it.
            if (_activeFlushBottom)
            {
                float y = SF(PanelH) - PenW / 2f;
                using (var pen = new Pen(TabTheme.PanelBorder, PenW))
                    g.DrawLine(pen, b.X - SF(FlareR) - PenW, y, b.Right + SF(FlareR) + PenW, y);
            }
            g.SmoothingMode = SmoothingMode.None;

            DrawTabContent(g, t, b, false, true,
                t.IsProtectedView ? TabTheme.ProtectedViewText : TabTheme.ActiveText, withText);
        }

        private void DrawTabContent(Graphics g, DocumentTabModel t, Rectangle b, bool compact, bool showClose,
                                    Color textColor, bool includeText)
        {
            Rectangle textRect, closeRect, dotRect, pinRect;
            GetTabContent(g, t, b, compact, showClose, out textRect, out closeRect, out dotRect, out pinRect);

            if (includeText && textRect.Width > 0)
                TextRenderer.DrawText(g, t.Caption ?? string.Empty, TabFont(t, compact), textRect, textColor, DrawFlags);

            // The pin of a pinned tab, like the dot below, is drawn in the FIRST pass
            // (independent of includeText): inside the overflow buffer it is GDI+ graphics
            // (FillPath) rather than TextRenderer, so clipping and shifting do not affect it.
            if (t.IsPinned && pinRect.Width > 0)
                Glyphs.DrawPin(g, pinRect, TabTheme.PinGlyph);

            if (!t.Saved && dotRect.Width > 0)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var dot = new SolidBrush(TabTheme.UnsavedDot))
                    g.FillEllipse(dot, dotRect);
                g.SmoothingMode = SmoothingMode.None;
            }

            if (showClose && closeRect.Width > 0)
                DrawCloseBox(g, closeRect, _hover.Kind == HitKind.TabClose && _hover.Hwnd == t.Hwnd);
        }

        private void DrawCloseBox(Graphics g, Rectangle r, bool hovered)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (hovered)
            {
                using (var path = Glyphs.RoundedRect(r, SF(4)))
                using (var bg = new SolidBrush(TabTheme.CloseHoverBg))
                    g.FillPath(bg, path);
            }
            float inset = SF(5.5f);
            using (var pen = new Pen(hovered ? TabTheme.CloseHoverGlyph : TabTheme.CloseGlyph, Math.Max(1.2f, SF(1.2f))))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                g.DrawLine(pen, r.Left + inset, r.Top + inset, r.Right - inset, r.Bottom - inset);
                g.DrawLine(pen, r.Right - inset, r.Top + inset, r.Left + inset, r.Bottom - inset);
            }
            g.SmoothingMode = SmoothingMode.None;
        }

        private void DrawIconButton(Graphics g, Rectangle r, string glyph, int glyphPx,
                                    bool enabled, bool hovered, bool pressed)
        {
            if (r.Width <= 0) return;
            if (enabled && (hovered || pressed))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Glyphs.RoundedRect(r, SF(5)))
                using (var bg = new SolidBrush(pressed ? TabTheme.ButtonPressedBg : TabTheme.ButtonHoverBg))
                    g.FillPath(bg, path);
                g.SmoothingMode = SmoothingMode.None;
            }

            var color = enabled ? TabTheme.ButtonGlyph : TabTheme.ButtonGlyphDisabled;

            // The all-tabs and collapse arrows are drawn with GDI+ lines rather than a font:
            // the Segoe UI glyphs sit below the middle of the line (the chevrons are
            // modifier characters), which made the buttons look shifted down and small.
            // A vector arrow sits geometrically in the centre of the button and is bigger.
            switch (glyph)
            {
                case "▾":
                    DrawMenuArrow(g, r, color);
                    return;
                case "⌄":
                    DrawChevron(g, r, color, true);
                    return;
                case "⌃":
                    DrawChevron(g, r, color, false);
                    return;
                case "‹":
                    DrawChevronSide(g, r, color, true);
                    return;
                case "›":
                    DrawChevronSide(g, r, color, false);
                    return;
            }

            TextRenderer.DrawText(g, glyph, GlyphFont(glyphPx), r,
                color,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        /// <summary>The all-tabs button: a filled triangle in the centre of the button.</summary>
        private void DrawMenuArrow(Graphics g, Rectangle r, Color color)
        {
            float cx = r.X + r.Width / 2f;
            float cy = r.Y + r.Height / 2f;
            float halfW = SF(5.5f);
            float halfH = SF(3.5f);
            var pts = new[]
            {
                new PointF(cx - halfW, cy - halfH),
                new PointF(cx + halfW, cy - halfH),
                new PointF(cx, cy + halfH)
            };
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var brush = new SolidBrush(color))
                g.FillPolygon(brush, pts);
            g.SmoothingMode = SmoothingMode.None;
        }

        /// <summary>
        /// Scroll arrows: a chevron in the centre of the button. Vector-drawn and thicker
        /// than the font glyphs - the previous ones were nearly invisible.
        /// </summary>
        private void DrawChevronSide(Graphics g, Rectangle r, Color color, bool left)
        {
            float cx = r.X + r.Width / 2f;
            float cy = r.Y + r.Height / 2f;
            float halfW = SF(3f);
            float halfH = SF(5.5f);
            PointF[] pts = left
                ? new[]
                {
                    new PointF(cx + halfW, cy - halfH),
                    new PointF(cx - halfW, cy),
                    new PointF(cx + halfW, cy + halfH)
                }
                : new[]
                {
                    new PointF(cx - halfW, cy - halfH),
                    new PointF(cx + halfW, cy),
                    new PointF(cx - halfW, cy + halfH)
                };
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(color, Math.Max(1.8f, SF(1.8f))))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                g.DrawLines(pen, pts);
            }
            g.SmoothingMode = SmoothingMode.None;
        }

        /// <summary>The collapse / expand buttons: a chevron in the centre of the button.</summary>
        private void DrawChevron(Graphics g, Rectangle r, Color color, bool down)
        {
            float cx = r.X + r.Width / 2f;
            float cy = r.Y + r.Height / 2f;
            float halfW = SF(5.5f);
            float halfH = SF(3f);
            PointF[] pts = down
                ? new[]
                {
                    new PointF(cx - halfW, cy - halfH),
                    new PointF(cx, cy + halfH),
                    new PointF(cx + halfW, cy - halfH)
                }
                : new[]
                {
                    new PointF(cx - halfW, cy + halfH),
                    new PointF(cx, cy - halfH),
                    new PointF(cx + halfW, cy + halfH)
                };
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(color, Math.Max(1.6f, SF(1.6f))))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                g.DrawLines(pen, pts);
            }
            g.SmoothingMode = SmoothingMode.None;
        }

        private void DrawFocusFrame(Graphics g, Rectangle tabBounds)
        {
            var r = new RectangleF(tabBounds.X + S(2), tabBounds.Y + S(2),
                                   tabBounds.Width - S(4), tabBounds.Height - S(4) - (tabBounds.Height > S(TabH) ? S(1) : 0));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = Glyphs.RoundedRect(r, SF(5)))
            using (var pen = new Pen(TabTheme.FocusFrame, PenW))
                g.DrawPath(pen, path);
            g.SmoothingMode = SmoothingMode.None;
        }

        private void DrawDraggedTab(Graphics g, StripLayout lay)
        {
            TabLayout slot = null;
            foreach (var tl in lay.Tabs)
            {
                if (tl.Tab.Hwnd == _dragHwnd) { slot = tl; break; }
            }
            if (slot == null) return;

            int flare = S(FlareR);
            int w = slot.Bounds.Width;
            int minX = lay.Viewport.Left + flare;
            int maxX = Math.Max(minX, lay.Viewport.Right - w - flare);
            int x = Math.Max(minX, Math.Min(maxX, _dragMouseX - _dragGrabOffsetX));
            var bounds = new Rectangle(x, S(TabTop), w, S(ActiveTabH));

            // Shadow: 0 4px 12px rgba(20,26,36,.2) - approximated with three layers
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int[] grow = { S(6), S(3), S(1) };
            int[] alpha = { 10, 18, 28 };
            for (int i = 0; i < grow.Length; i++)
            {
                var sr = new RectangleF(bounds.X - grow[i], bounds.Y + S(4) - grow[i] / 2f,
                                        bounds.Width + grow[i] * 2, bounds.Height + grow[i]);
                using (var path = Glyphs.RoundedRect(sr, SF(CornerR) + grow[i]))
                using (var sb = new SolidBrush(Color.FromArgb(alpha[i], TabTheme.DragShadow)))
                    g.FillPath(sb, path);
            }
            g.SmoothingMode = SmoothingMode.None;

            // The tab drawn "as if active" at 0.9 opacity (section 6) - through an offscreen
            // bitmap. No TranslateTransform (TextRenderer ignores the transform - the caption
            // simply was not drawn): coordinates inside the bitmap are explicit, the bitmap
            // covers the full strip height, and the blit goes to (x-flare-1, 0).
            int bmpW = w + flare * 2 + 2;
            int bmpH = S(PanelH) + 2;
            using (var bmp = new Bitmap(bmpW, bmpH, PixelFormat.Format32bppArgb))
            {
                var local = new Rectangle(flare + 1, bounds.Y, bounds.Width, bounds.Height);
                using (var bg = Graphics.FromImage(bmp))
                {
                    Color fillColor, strokeColor;
                    ActiveTabColors(slot.Tab, out fillColor, out strokeColor);
                    bg.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var fill = ActiveFillPath(local.X, local.Right))
                    using (var brush = new SolidBrush(fillColor))
                        bg.FillPath(brush, fill);
                    using (var stroke = ActiveStrokePath(local.X, local.Right))
                    using (var pen = new Pen(strokeColor, PenW))
                        bg.DrawPath(pen, stroke);
                    bg.SmoothingMode = SmoothingMode.None;
                    // the caption does not go into the bitmap (GDI text there is unsmoothed) - it goes on screen below
                    DrawTabContent(bg, slot.Tab, local, false, true, TabTheme.ActiveText, false);
                }
                var cm = new ColorMatrix { Matrix33 = 0.9f };
                using (var ia = new ImageAttributes())
                {
                    ia.SetColorMatrix(cm);
                    g.DrawImage(bmp,
                        new Rectangle(bounds.X - flare - 1, 0, bmpW, bmpH),
                        0, 0, bmpW, bmpH, GraphicsUnit.Pixel, ia);
                }
            }

            // Caption of the floating tab - on the screen DC (crisp ClearType)
            DrawTabCaption(g, slot.Tab, bounds, false, true, TabTheme.ActiveText, ClientRectangle);
        }

        // ------------------------------------------------------------------
        // Hit testing
        // ------------------------------------------------------------------

        private Hit HitTest(Point p)
        {
            var none = new Hit { Kind = HitKind.None };
            try
            {
                if (_collapsed)
                {
                    if (CollapsedExpandRect().Contains(p)) return new Hit { Kind = HitKind.Collapse };
                    return p.Y <= S(CollapsedH) ? new Hit { Kind = HitKind.Empty } : none;
                }
                if (p.Y > S(PanelH)) return none;

                using (var g = CreateGraphics())
                {
                    var lay = ComputeLayout(g, DisplayTabs());

                    if (lay.CollapseBtn.Contains(p)) return new Hit { Kind = HitKind.Collapse };
                    if (lay.MenuBtn.Contains(p)) return new Hit { Kind = HitKind.Menu };
                    if (lay.Plus.Contains(p)) return new Hit { Kind = HitKind.Plus };
                    if (lay.Overflow)
                    {
                        if (lay.ScrollLeft.Contains(p)) return new Hit { Kind = HitKind.ScrollLeft };
                        if (lay.ScrollRight.Contains(p)) return new Hit { Kind = HitKind.ScrollRight };
                    }

                    if (!lay.Overflow || lay.Viewport.Contains(p))
                    {
                        // The active tab catches the click first (it is drawn on top)
                        for (int pass = 0; pass < 2; pass++)
                        {
                            foreach (var tl in lay.Tabs)
                            {
                                bool activePass = pass == 0;
                                if (tl.Tab.IsActive != activePass) continue;
                                if (!tl.Bounds.Contains(p)) continue;

                                bool hovered = true;
                                bool showClose = tl.Tab.IsActive || (hovered && !lay.Compact);
                                if (showClose)
                                {
                                    Rectangle textRect, closeRect, dotRect, pinRect;
                                    GetTabContent(g, tl.Tab, tl.Bounds, lay.Compact, true,
                                                  out textRect, out closeRect, out dotRect, out pinRect);
                                    if (closeRect.Contains(p))
                                        return new Hit { Kind = HitKind.TabClose, Hwnd = tl.Tab.Hwnd };
                                }
                                return new Hit { Kind = HitKind.Tab, Hwnd = tl.Tab.Hwnd };
                            }
                        }
                    }
                }
                return new Hit { Kind = HitKind.Empty };
            }
            catch (Exception ex)
            {
                LoggingService.Error("HitTest failed", ex);
                return none;
            }
        }

        // ------------------------------------------------------------------
        // Mouse
        // ------------------------------------------------------------------

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            try
            {
                if (_mouseDownLeft && _mouseDownHit.Kind == HitKind.Tab)
                {
                    if (!_dragging && Math.Abs(e.X - _mouseDownPt.X) > S(DragThreshold))
                        StartDrag(e);
                    if (_dragging)
                    {
                        UpdateDrag(e);
                        return;
                    }
                }

                var hit = HitTest(e.Location);
                if (!hit.Same(_hover))
                {
                    _hover = hit;
                    UpdateTooltip(hit);
                    Invalidate();
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnMouseMove failed", ex);
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hover = new Hit { Kind = HitKind.None };
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            try
            {
                var hit = HitTest(e.Location);
                if (e.Button == MouseButtons.Left)
                {
                    _mouseDownLeft = true;
                    _mouseDownPt = e.Location;
                    _mouseDownHit = hit;
                    _pressed = hit;
                    if (hit.Kind == HitKind.Tab)
                    {
                        int idx = _tabs.FindIndex(t => t.Hwnd == hit.Hwnd);
                        if (idx >= 0) _focusIndex = idx;
                    }
                    Invalidate();
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnMouseDown failed", ex);
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            try
            {
                var hit = HitTest(e.Location);

                if (e.Button == MouseButtons.Left)
                {
                    _mouseDownLeft = false;
                    _pressed = new Hit { Kind = HitKind.None };

                    if (_dragging)
                    {
                        CommitDrag();
                        Invalidate();
                        return;
                    }

                    if (hit.Same(_mouseDownHit))
                        PerformLeftClick(hit);
                    Invalidate();
                    return;
                }

                if (e.Button == MouseButtons.Middle && hit.Kind == HitKind.Tab)
                {
                    // Middle button closes (section 8)
                    var close = TabCloseRequested;
                    if (close != null) close(hit.Hwnd);
                    return;
                }

                if (e.Button == MouseButtons.Right &&
                    (hit.Kind == HitKind.Tab || hit.Kind == HitKind.TabClose))
                {
                    var tab = _tabs.FirstOrDefault(t => t.Hwnd == hit.Hwnd);
                    if (tab != null) ShowTabContextMenu(tab, e.Location);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnMouseUp failed", ex);
            }
        }

        private void PerformLeftClick(Hit hit)
        {
            switch (hit.Kind)
            {
                case HitKind.Tab:
                    var act = TabActivateRequested;
                    if (act != null) act(hit.Hwnd);
                    break;
                case HitKind.TabClose:
                    var close = TabCloseRequested;
                    if (close != null) close(hit.Hwnd);
                    break;
                case HitKind.Plus:
                    var nw = NewDocumentRequested;
                    if (nw != null) nw();
                    break;
                case HitKind.Menu:
                    ShowTabsMenu();
                    break;
                case HitKind.Collapse:
                    var col = CollapseToggleRequested;
                    if (col != null) col();
                    break;
                case HitKind.ScrollLeft:
                    _scrollOffset -= S(ScrollStep);
                    Invalidate();
                    break;
                case HitKind.ScrollRight:
                    _scrollOffset += S(ScrollStep);
                    Invalidate();
                    break;
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            try
            {
                if (e.Button != MouseButtons.Left) return;
                var hit = HitTest(e.Location);
                // Double click on empty strip space creates a new document (section 8)
                if (hit.Kind == HitKind.Empty && !_collapsed)
                {
                    var nw = NewDocumentRequested;
                    if (nw != null) nw();
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnMouseDoubleClick failed", ex);
            }
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            try
            {
                if (_collapsed) return;
                _scrollOffset -= Math.Sign(e.Delta) * S(ScrollStep);
                Invalidate();
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnMouseWheel failed", ex);
            }
        }

        // ------------------------------------------------------------------
        // Dragging
        // ------------------------------------------------------------------

        private void StartDrag(MouseEventArgs e)
        {
            try
            {
                _dragging = true;
                _dragHwnd = _mouseDownHit.Hwnd;
                _dragOrder = new List<int>();
                foreach (var t in _tabs) _dragOrder.Add(t.Hwnd);
                _dragMouseX = e.X;

                using (var g = CreateGraphics())
                {
                    var lay = ComputeLayout(g, _tabs);
                    _dragGrabOffsetX = S(20);
                    foreach (var tl in lay.Tabs)
                    {
                        if (tl.Tab.Hwnd == _dragHwnd)
                        {
                            _dragGrabOffsetX = Math.Max(0, Math.Min(tl.Bounds.Width, _mouseDownPt.X - tl.Bounds.X));
                            break;
                        }
                    }
                }
                _hover = new Hit { Kind = HitKind.None };
                Invalidate();
            }
            catch (Exception ex)
            {
                LoggingService.Error("StartDrag failed", ex);
                CancelDrag();
            }
        }

        private void UpdateDrag(MouseEventArgs e)
        {
            try
            {
                _dragMouseX = e.X;
                using (var g = CreateGraphics())
                {
                    var disp = DisplayTabs();
                    var lay = ComputeLayout(g, disp);
                    int cur = _dragOrder.IndexOf(_dragHwnd);
                    if (cur < 0 || cur >= lay.Tabs.Count) { Invalidate(); return; }

                    float cx = _dragMouseX - _dragGrabOffsetX + lay.Tabs[cur].Bounds.Width / 2f;
                    int target = 0;
                    foreach (var tl in lay.Tabs)
                    {
                        if (tl.Tab.Hwnd == _dragHwnd) continue;
                        if (tl.Bounds.X + tl.Bounds.Width / 2f < cx) target++;
                    }
                    if (target != cur)
                    {
                        _dragOrder.RemoveAt(cur);
                        _dragOrder.Insert(target, _dragHwnd);
                        ApplyPinOrderInvariant(_dragOrder);
                    }
                }
                Invalidate();
            }
            catch (Exception ex)
            {
                LoggingService.Error("UpdateDrag failed", ex);
            }
        }

        /// <summary>
        /// The same "pinned first" invariant as in DocumentWindowManager
        /// (Reconcile/ReorderTab) - applied to the live drag preview so that the order
        /// does not "jump" on commit relative to what the user saw while dragging.
        /// </summary>
        private void ApplyPinOrderInvariant(List<int> order)
        {
            var pinned = new HashSet<int>();
            foreach (var t in _tabs) { if (t.IsPinned) pinned.Add(t.Hwnd); }

            var result = new List<int>(order.Count);
            foreach (var h in order) { if (pinned.Contains(h)) result.Add(h); }
            foreach (var h in order) { if (!pinned.Contains(h)) result.Add(h); }

            order.Clear();
            order.AddRange(result);
        }

        private void CommitDrag()
        {
            int hwnd = _dragHwnd;
            int from = _tabs.FindIndex(t => t.Hwnd == hwnd);
            int to = _dragOrder != null ? _dragOrder.IndexOf(hwnd) : -1;
            CancelDrag();
            if (from >= 0 && to >= 0 && from != to)
            {
                var reorder = ReorderRequested;
                if (reorder != null) reorder(hwnd, to);
            }
        }

        private void CancelDrag()
        {
            _dragging = false;
            _dragOrder = null;
            _mouseDownLeft = false;
            _pressed = new Hit { Kind = HitKind.None };
        }

        // ------------------------------------------------------------------
        // Context menu and the all-tabs menu
        // ------------------------------------------------------------------

        // ------------------------------------------------------------------
        // The tab context menu is a classic Win32 menu (TrackPopupMenu).
        //
        // History (important): ContextMenuStrip behaved fatally inside Word.
        // 1. Disposing it from its own Closed event brought WINWORD down
        //    (0xE0434352): internal ToolStrip code keeps running after Closed, and
        //    there is nobody to catch an exception from a window procedure - Word
        //    has no WinForms message loop.
        // 2. Even without the Dispose, the first menu shown in a session would not
        //    close on a click outside: the ToolStrip close mechanics rely on WinForms
        //    message filters, which are not active inside the Word loop.
        // A classic menu is modal and driven by user32 itself: it always closes on any
        // click outside and creates no WinForms windows at all.
        // ------------------------------------------------------------------

        private ContextMenu _tabMenu;
        private MenuItem _menuPin;
        private MenuItem _menuClose;
        private MenuItem _menuCloseOthers;
        private MenuItem _menuOpenFolder;
        private MenuItem _menuColorRoot;
        private MenuItem[] _menuColorItems;   // palette presets (owner-draw: a swatch plus the name)
        private MenuItem _menuColorCustom;
        private MenuItem _menuColorReset;
        private MenuItem _menuSizeRoot;
        private MenuItem[] _menuSizeItems;    // tab sizes (RadioCheck)
        private int _menuTargetHwnd;

        // Tab size: multipliers on top of the automatic DPI scale.
        // The captions come from Strings.SizeNames (same order).
        private static readonly float[] SizeValues = { 0.85f, 1.0f, 1.15f, 1.3f };

        // Palette of ready tab colours (saturated originals; the tab itself is painted
        // with their pastel derivative, see TabTheme.Tint*).
        // The captions come from Strings.ColorNames (same order).
        private static readonly Color[] PaletteColors =
        {
            Color.FromArgb(0xD9, 0x30, 0x25), Color.FromArgb(0xE8, 0x71, 0x0A),
            Color.FromArgb(0xF9, 0xAB, 0x00), Color.FromArgb(0x1E, 0x8E, 0x3E),
            Color.FromArgb(0x0F, 0x9D, 0xA8), Color.FromArgb(0x1A, 0x73, 0xE8),
            Color.FromArgb(0x93, 0x34, 0xE6), Color.FromArgb(0xD0, 0x18, 0x84)
        };

        /// <summary>
        /// Re-reads every localized caption after the user switches the language.
        /// The context menu is built once and cached, so it has to be dropped;
        /// tooltips and painted text are produced on demand and only need a repaint.
        /// </summary>
        public void RefreshLocalizedUi()
        {
            try
            {
                if (_tabMenu != null) { _tabMenu.Dispose(); _tabMenu = null; }
                _menuPin = _menuClose = _menuCloseOthers = _menuOpenFolder = null;
                _menuColorRoot = _menuColorCustom = _menuColorReset = _menuSizeRoot = null;
                _menuColorItems = null;
                _menuSizeItems = null;
                Invalidate();
            }
            catch (Exception ex)
            {
                LoggingService.Error("RefreshLocalizedUi failed", ex);
            }
        }

        private void EnsureTabMenu()
        {
            if (_tabMenu != null) return;

            _menuPin = new MenuItem(Strings.MenuPin, OnMenuPinClick);
            _menuClose = new MenuItem(Strings.MenuClose, OnMenuCloseClick);
            _menuCloseOthers = new MenuItem(Strings.MenuCloseOthers, OnMenuCloseOthersClick);
            _menuOpenFolder = new MenuItem(Strings.MenuOpenFolder, OnMenuOpenFolderClick);

            var paletteNames = Strings.ColorNames;
            _menuColorItems = new MenuItem[PaletteColors.Length];
            var colorSub = new List<MenuItem>();
            for (int i = 0; i < PaletteColors.Length; i++)
            {
                var item = new MenuItem(paletteNames[i], OnMenuColorPresetClick);
                item.OwnerDraw = true;
                item.MeasureItem += OnColorItemMeasure;
                item.DrawItem += OnColorItemDraw;
                _menuColorItems[i] = item;
                colorSub.Add(item);
            }
            colorSub.Add(new MenuItem("-"));
            _menuColorCustom = new MenuItem(Strings.MenuColorCustom, OnMenuColorCustomClick);
            colorSub.Add(_menuColorCustom);
            _menuColorReset = new MenuItem(Strings.MenuColorReset, OnMenuColorResetClick);
            colorSub.Add(_menuColorReset);
            _menuColorRoot = new MenuItem(Strings.MenuTabColor);
            _menuColorRoot.MenuItems.AddRange(colorSub.ToArray());

            var sizeNames = Strings.SizeNames;
            _menuSizeItems = new MenuItem[SizeValues.Length];
            var sizeSub = new List<MenuItem>();
            for (int i = 0; i < SizeValues.Length; i++)
            {
                var item = new MenuItem(sizeNames[i], OnMenuSizePresetClick);
                item.RadioCheck = true;
                _menuSizeItems[i] = item;
                sizeSub.Add(item);
            }
            _menuSizeRoot = new MenuItem(Strings.MenuTabSize);
            _menuSizeRoot.MenuItems.AddRange(sizeSub.ToArray());

            _tabMenu = new ContextMenu(new[]
            {
                _menuPin,
                new MenuItem("-"),
                _menuClose,
                _menuCloseOthers,
                new MenuItem("-"),
                _menuColorRoot,
                _menuSizeRoot,
                new MenuItem("-"),
                _menuOpenFolder
            });
        }

        private void OnMenuPinClick(object sender, EventArgs e)
        {
            try
            {
                var tab = _tabs.FirstOrDefault(t => t.Hwnd == _menuTargetHwnd);
                bool nowPinned = tab == null || !tab.IsPinned;
                var h = TabPinChangeRequested;
                if (h != null) h(_menuTargetHwnd, nowPinned);
            }
            catch (Exception ex) { LoggingService.Error("Menu Pin failed", ex); }
        }

        private void OnMenuCloseClick(object sender, EventArgs e)
        {
            try
            {
                var h = TabCloseRequested;
                if (h != null) h(_menuTargetHwnd);
            }
            catch (Exception ex) { LoggingService.Error("Menu Close failed", ex); }
        }

        private void OnMenuCloseOthersClick(object sender, EventArgs e)
        {
            try
            {
                var h = CloseOthersRequested;
                if (h != null) h(_menuTargetHwnd);
            }
            catch (Exception ex) { LoggingService.Error("Menu CloseOthers failed", ex); }
        }

        private void OnMenuOpenFolderClick(object sender, EventArgs e)
        {
            try
            {
                var h = OpenFolderRequested;
                if (h != null) h(_menuTargetHwnd);
            }
            catch (Exception ex) { LoggingService.Error("Menu OpenFolder failed", ex); }
        }

        // ------------------------------------------------------------------
        // Tab colour
        // ------------------------------------------------------------------

        private void RaiseTabColorChange(Color? color)
        {
            var h = TabColorChangeRequested;
            if (h != null) h(_menuTargetHwnd, color);
        }

        private Color? MenuTargetColor()
        {
            var t = _tabs.Find(x => x.Hwnd == _menuTargetHwnd);
            return t != null ? t.TabColor : null;
        }

        private void OnMenuColorPresetClick(object sender, EventArgs e)
        {
            try
            {
                int i = Array.IndexOf(_menuColorItems, sender as MenuItem);
                if (i >= 0) RaiseTabColorChange(PaletteColors[i]);
            }
            catch (Exception ex) { LoggingService.Error("Menu color preset failed", ex); }
        }

        private void OnMenuColorCustomClick(object sender, EventArgs e)
        {
            try
            {
                // The standard Windows colour dialog (ChooseColor) - native and modal,
                // it needs no WinForms message loop.
                using (var dlg = new ColorDialog())
                {
                    dlg.FullOpen = true;
                    dlg.AnyColor = true;
                    var current = MenuTargetColor();
                    if (current.HasValue) dlg.Color = current.Value;
                    // Our palette goes into the "custom colours" cells (COLORREF = BGR)
                    var custom = new int[PaletteColors.Length];
                    for (int i = 0; i < PaletteColors.Length; i++)
                        custom[i] = PaletteColors[i].R | (PaletteColors[i].G << 8) | (PaletteColors[i].B << 16);
                    dlg.CustomColors = custom;

                    if (dlg.ShowDialog(this) == DialogResult.OK)
                        RaiseTabColorChange(Color.FromArgb(dlg.Color.R, dlg.Color.G, dlg.Color.B));
                }
            }
            catch (Exception ex) { LoggingService.Error("Menu color custom failed", ex); }
        }

        private void OnMenuColorResetClick(object sender, EventArgs e)
        {
            try { RaiseTabColorChange(null); }
            catch (Exception ex) { LoggingService.Error("Menu color reset failed", ex); }
        }

        private void OnMenuSizePresetClick(object sender, EventArgs e)
        {
            try
            {
                int i = Array.IndexOf(_menuSizeItems, sender as MenuItem);
                if (i < 0) return;
                var h = TabSizeChangeRequested;
                if (h != null) h(SizeValues[i]);
            }
            catch (Exception ex) { LoggingService.Error("Menu size preset failed", ex); }
        }

        private int MenuSwatchSize()
        {
            return S(14);
        }

        private void OnColorItemMeasure(object sender, MeasureItemEventArgs e)
        {
            try
            {
                var item = (MenuItem)sender;
                var font = SystemInformation.MenuFont;
                var sz = TextRenderer.MeasureText(e.Graphics, item.Text, font, new Size(int.MaxValue, int.MaxValue), MeasureFlags);
                int swatch = MenuSwatchSize();
                e.ItemHeight = Math.Max(sz.Height, swatch) + S(8);
                e.ItemWidth = S(6) + swatch + S(8) + sz.Width + S(16);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Color item measure failed", ex);
                e.ItemHeight = 24;
                e.ItemWidth = 140;
            }
        }

        private void OnColorItemDraw(object sender, DrawItemEventArgs e)
        {
            try
            {
                var item = (MenuItem)sender;
                int i = Array.IndexOf(_menuColorItems, item);
                if (i < 0) return;
                var g = e.Graphics;
                bool selected = (e.State & DrawItemState.Selected) != 0;
                bool isChecked = (e.State & DrawItemState.Checked) != 0;

                using (var bg = new SolidBrush(selected ? SystemColors.Highlight : SystemColors.Menu))
                    g.FillRectangle(bg, e.Bounds);

                int swatch = MenuSwatchSize();
                var sw = new Rectangle(e.Bounds.X + S(6),
                                       e.Bounds.Y + (e.Bounds.Height - swatch) / 2, swatch, swatch);
                var old = g.SmoothingMode;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Glyphs.RoundedRect(new RectangleF(sw.X, sw.Y, sw.Width, sw.Height), SF(3)))
                {
                    using (var fill = new SolidBrush(PaletteColors[i]))
                        g.FillPath(fill, path);
                    using (var pen = new Pen(TabTheme.Blend(PaletteColors[i], Color.Black, 0.25f), 1f))
                        g.DrawPath(pen, path);
                }
                if (isChecked)
                {
                    // Current tab colour - a contrasting ring around the swatch
                    var ring = new RectangleF(sw.X - 2.5f, sw.Y - 2.5f, sw.Width + 5f, sw.Height + 5f);
                    using (var path = Glyphs.RoundedRect(ring, SF(4.5f)))
                    using (var pen = new Pen(selected ? SystemColors.HighlightText : SystemColors.MenuText, 1.5f))
                        g.DrawPath(pen, path);
                }
                g.SmoothingMode = old;

                var textRect = new Rectangle(sw.Right + S(8), e.Bounds.Y,
                                             e.Bounds.Right - sw.Right - S(8), e.Bounds.Height);
                TextRenderer.DrawText(g, item.Text, SystemInformation.MenuFont, textRect,
                    selected ? SystemColors.HighlightText : SystemColors.MenuText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Color item draw failed", ex);
            }
        }

        private static bool SameRgb(Color a, Color b)
        {
            return a.R == b.R && a.G == b.G && a.B == b.B;
        }

        private void ShowTabContextMenu(DocumentTabModel tab, Point location)
        {
            try
            {
                EnsureTabMenu();
                _menuTargetHwnd = tab.Hwnd;
                _menuPin.Text = tab.IsPinned ? Strings.MenuUnpin : Strings.MenuPin;
                // "Close others" leaves other pinned tabs alone - if there is nothing
                // left to close besides the current tab, the item is disabled.
                _menuCloseOthers.Enabled = _tabs.Any(t => t.Hwnd != tab.Hwnd && !t.IsPinned);

                bool hasPath = false;
                try { hasPath = !string.IsNullOrEmpty(tab.FullPath) && System.IO.Path.IsPathRooted(tab.FullPath); }
                catch { }
                _menuOpenFolder.Enabled = hasPath;

                // Tick the current colour in the palette; "Reset" is enabled only with a colour
                for (int i = 0; i < _menuColorItems.Length; i++)
                    _menuColorItems[i].Checked = tab.TabColor.HasValue && SameRgb(tab.TabColor.Value, PaletteColors[i]);
                _menuColorReset.Enabled = tab.TabColor.HasValue;

                // Tick the current size (the dot is only put on the presets)
                for (int i = 0; i < _menuSizeItems.Length; i++)
                    _menuSizeItems[i].Checked = Math.Abs(_userScale - SizeValues[i]) < 0.01f;

                _tabMenu.Show(this, location);
            }
            catch (Exception ex)
            {
                LoggingService.Error("ShowTabContextMenu failed", ex);
            }
        }

        private void ShowTabsMenu()
        {
            try
            {
                using (var g = CreateGraphics())
                {
                    var lay = ComputeLayout(g, DisplayTabs());
                    var anchor = PointToScreen(new Point(lay.MenuBtn.Right, S(PanelH) + S(2)));
                    var popup = new TabListPopup(_tabs, _scale, delegate(int hwnd)
                    {
                        var h = TabActivateRequested;
                        if (h != null) h(hwnd);
                    }, delegate { ShowSettingsForm(anchor); });
                    popup.ShowBelow(anchor);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("ShowTabsMenu failed", ex);
            }
        }

        /// <summary>
        /// The settings window (the "Settings..." button in the footer of the all-tabs
        /// menu): tab size, Ctrl+Tab interception and language. It opens on the same
        /// anchor as the all-tabs menu - the same family of strip popups.
        /// </summary>
        private void ShowSettingsForm(Point anchor)
        {
            try
            {
                var form = new SettingsForm(_userScale, HotkeySettings.CtrlTabEnabled, _scale,
                    scale =>
                    {
                        var h = TabSizeChangeRequested;
                        if (h != null) h(scale);
                    },
                    enabled =>
                    {
                        var h = HotkeyToggleRequested;
                        if (h != null) h(enabled);
                    },
                    lang =>
                    {
                        var h = LanguageChangeRequested;
                        if (h != null) h(lang);
                    });
                form.ShowBelow(anchor);
            }
            catch (Exception ex)
            {
                LoggingService.Error("ShowSettingsForm failed", ex);
            }
        }

        // ------------------------------------------------------------------
        // Keyboard (sections 8 and 9)
        // ------------------------------------------------------------------

        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData & Keys.KeyCode)
            {
                case Keys.Left:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.Enter:
                    return true;
            }
            return base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            try
            {
                if (_tabs.Count == 0 || _collapsed) return;
                switch (e.KeyCode)
                {
                    case Keys.Left:
                        _focusIndex = _focusIndex <= 0 ? _tabs.Count - 1 : _focusIndex - 1;
                        e.Handled = true;
                        Invalidate();
                        break;
                    case Keys.Right:
                        _focusIndex = (_focusIndex + 1) % _tabs.Count;
                        e.Handled = true;
                        Invalidate();
                        break;
                    case Keys.Home:
                        _focusIndex = 0;
                        e.Handled = true;
                        Invalidate();
                        break;
                    case Keys.End:
                        _focusIndex = _tabs.Count - 1;
                        e.Handled = true;
                        Invalidate();
                        break;
                    case Keys.Enter:
                    case Keys.Space:
                        if (_focusIndex >= 0 && _focusIndex < _tabs.Count)
                        {
                            var act = TabActivateRequested;
                            if (act != null) act(_tabs[_focusIndex].Hwnd);
                            e.Handled = true;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnKeyDown failed", ex);
            }
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            if (_focusIndex < 0)
                _focusIndex = _tabs.FindIndex(t => t.IsActive);
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            Invalidate();
        }

        // ------------------------------------------------------------------
        // Tooltip
        // ------------------------------------------------------------------

        private void UpdateTooltip(Hit hit)
        {
            string text = null;
            switch (hit.Kind)
            {
                case HitKind.Tab:
                    var t = _tabs.Find(x => x.Hwnd == hit.Hwnd);
                    if (t != null)
                    {
                        // Full file name and path (section 8)
                        text = t.Caption + "\r\n" + t.FullPath;
                        if (t.IsProtectedView) text += "\r\n" + Strings.TipProtectedView;
                        else if (!t.Saved) text += "\r\n" + Strings.TipUnsaved;
                    }
                    break;
                case HitKind.TabClose:
                    text = Strings.TipCloseDocument;
                    break;
                case HitKind.Plus:
                    text = Strings.TipNewDocument;
                    break;
                case HitKind.Menu:
                    text = Strings.TipAllTabs;
                    break;
                case HitKind.Collapse:
                    text = _collapsed ? Strings.TipExpandBar : Strings.TipCollapseBar;
                    break;
            }
            if (text != _lastTooltip)
            {
                _lastTooltip = text;
                _toolTip.SetToolTip(this, text);
            }
        }

        /// <summary>
        /// A safety net for the whole control: a managed exception escaping a window
        /// procedure into the Word message loop brings WINWORD down (0xE0434352).
        /// The handlers are already in try/catch - this is the last line of defence for
        /// WinForms.
        /// </summary>
        private const int WM_DPICHANGED_BEFOREPARENT = 0x02E2;
        private const int WM_DPICHANGED_AFTERPARENT = 0x02E3;
        private const int WM_MOUSEHWHEEL = 0x020E;

        protected override void WndProc(ref Message m)
        {
            try
            {
                // The scale is driven by the strip owner (SetDpiOverride from the DPI of the
                // Word window) - the WinForms auto-response to a DPI change resized the
                // control and drove DeviceDpi at the wrong moments on multi-monitor setups.
                if (m.Msg == WM_DPICHANGED_BEFOREPARENT || m.Msg == WM_DPICHANGED_AFTERPARENT)
                {
                    m.Result = IntPtr.Zero;
                    return;
                }
                // Horizontal scrolling with a touchpad (two fingers sideways): Windows sends
                // WM_MOUSEHWHEEL to the window under the cursor; delta arrives in small
                // steps, and the scroll is proportional to it (so it feels smooth).
                if (m.Msg == WM_MOUSEHWHEEL)
                {
                    if (!_collapsed)
                    {
                        int delta = unchecked((short)((long)m.WParam >> 16));
                        _scrollOffset += (int)Math.Round(delta * (float)S(ScrollStep) / 120f);
                        Invalidate();
                    }
                    m.Result = IntPtr.Zero;
                    return;
                }
                base.WndProc(ref m);
            }
            catch (Exception ex)
            {
                LoggingService.Error("TabStripControl WndProc failed (msg=0x" + m.Msg.ToString("X4") + ")", ex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _toolTip.Dispose(); } catch { }
                try { if (_tabMenu != null) { _tabMenu.Dispose(); _tabMenu = null; } } catch { }
                try { DisposeFonts(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
