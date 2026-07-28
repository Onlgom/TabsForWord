using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TabsForWord
{
    /// <summary>
    /// The all-tabs menu (the small arrow button) per specification 2b, sections 4 and 7:
    /// a 300 px window with radius 8, a 30 px search box, 26 px rows (radius 5), the
    /// active document marked with a tick on #EEF4FC, unsaved ones with an amber dot.
    /// Row order = tab order. Closes on Esc and on losing focus.
    /// </summary>
    internal sealed class TabListPopup : Form
    {
        private const int MenuW = 300;
        private const int Pad = 6;
        private const int SearchH = 30;
        private const int RowH = 26;
        private const int MaxVisibleRows = 12;

        private readonly List<DocumentTabModel> _all;
        private readonly Action<int> _onActivate;
        private readonly Action _onSettings;
        private readonly float _scale;

        private readonly TextBox _search;
        private readonly Panel _searchHost;
        private readonly ListBox _list;
        private readonly Label _empty;
        private readonly Panel _footer;
        private bool _footerHover;
        private readonly Font _fontText;
        private readonly Font _fontCheck;
        private List<DocumentTabModel> _filtered;

        public TabListPopup(IList<DocumentTabModel> tabs, float scale, Action<int> onActivate, Action onSettings)
        {
            _all = new List<DocumentTabModel>(tabs);
            _scale = scale;
            _onActivate = onActivate;
            _onSettings = onSettings;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = TabTheme.MenuBg;
            DoubleBuffered = true;

            _fontText = new Font("Segoe UI", SF(12.5f), FontStyle.Regular, GraphicsUnit.Pixel);
            _fontCheck = new Font("Segoe UI", SF(12f), FontStyle.Regular, GraphicsUnit.Pixel);

            Width = S(MenuW);

            _searchHost = new Panel
            {
                Left = S(Pad),
                Top = S(Pad),
                Width = Width - S(Pad) * 2,
                Height = S(SearchH),
                BackColor = TabTheme.MenuBg
            };
            _searchHost.Paint += PaintSearchBorder;

            _search = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font = _fontText,
                BackColor = TabTheme.MenuBg,
                ForeColor = TabTheme.MenuText
            };
            _search.Left = S(8);
            _search.Width = _searchHost.Width - S(16);
            _search.Top = Math.Max(0, (_searchHost.Height - _search.Height) / 2);
            _search.TextChanged += delegate { ApplyFilter(); };
            _searchHost.Controls.Add(_search);
            Controls.Add(_searchHost);

            _list = new ListBox
            {
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = S(RowH),
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                BackColor = TabTheme.MenuBg,
                Left = S(Pad),
                Top = _searchHost.Bottom + S(Pad),
                Width = Width - S(Pad) * 2
            };
            _list.DrawItem += DrawRow;
            _list.MouseMove += delegate(object s, MouseEventArgs e)
            {
                int i = _list.IndexFromPoint(e.Location);
                if (i >= 0 && i != _list.SelectedIndex) _list.SelectedIndex = i;
            };
            _list.MouseUp += delegate(object s, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Left) return;
                int i = _list.IndexFromPoint(e.Location);
                if (i >= 0) ActivateRow(i);
            };
            Controls.Add(_list);

            _empty = new Label
            {
                Text = Strings.PopupNoMatches,
                ForeColor = TabTheme.ProtectedViewText,
                Font = _fontText,
                TextAlign = ContentAlignment.MiddleCenter,
                Left = S(Pad),
                Width = Width - S(Pad) * 2,
                Visible = false
            };
            Controls.Add(_empty);

            _footer = new Panel
            {
                Left = S(Pad),
                Width = Width - S(Pad) * 2,
                Height = S(RowH),
                BackColor = TabTheme.MenuBg,
                Cursor = Cursors.Hand
            };
            _footer.Paint += PaintFooter;
            _footer.MouseEnter += delegate { _footerHover = true; _footer.Invalidate(); };
            _footer.MouseLeave += delegate { _footerHover = false; _footer.Invalidate(); };
            _footer.MouseUp += delegate(object s, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Left) return;
                try { Close(); } catch { }
                try { if (_onSettings != null) _onSettings(); }
                catch (Exception ex) { LoggingService.Error("Popup settings failed", ex); }
            };
            Controls.Add(_footer);

            ApplyFilter();
        }

        private void PaintFooter(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var bg = new SolidBrush(_footerHover ? TabTheme.MenuHoverRow : TabTheme.MenuBg))
                g.FillRectangle(bg, _footer.ClientRectangle);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(TabTheme.MenuBorder, 1f))
                g.DrawLine(pen, 0, 0.5f, _footer.Width, 0.5f);
            g.SmoothingMode = SmoothingMode.None;
            var textRect = new Rectangle(S(6), 0, _footer.Width - S(12), _footer.Height);
            TextRenderer.DrawText(g, Strings.PopupSettings, _fontText, textRect, TabTheme.MenuText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        private int S(float v) { return (int)Math.Round(v * _scale); }
        private float SF(float v) { return v * _scale; }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= NativeMethods.CS_DROPSHADOW;
                cp.ExStyle |= NativeMethods.WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            NativeMethods.TryRoundCorners(Handle);
            NativeMethods.TrySetCueBanner(_search.Handle, Strings.PopupSearchHint);
        }

        /// <summary>Shows the menu below the strip, right edge at the anchor point.</summary>
        public void ShowBelow(Point anchorRightBottom)
        {
            var wa = Screen.FromPoint(anchorRightBottom).WorkingArea;
            int x = anchorRightBottom.X - Width;
            int y = anchorRightBottom.Y;
            if (x < wa.Left) x = wa.Left;
            if (x + Width > wa.Right) x = wa.Right - Width;
            if (y + Height > wa.Bottom) y = Math.Max(wa.Top, wa.Bottom - Height);
            Location = new Point(x, y);
            Show();
            Activate();
            _search.Focus();
        }

        private void ApplyFilter()
        {
            string q = (_search.Text ?? string.Empty).Trim();
            _filtered = new List<DocumentTabModel>();
            foreach (var t in _all)
            {
                if (q.Length == 0 ||
                    (t.Caption ?? string.Empty).IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (t.FullPath ?? string.Empty).IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _filtered.Add(t);
                }
            }

            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var t in _filtered) _list.Items.Add(t.Caption ?? string.Empty);
            _list.EndUpdate();

            int rows = Math.Max(1, Math.Min(_filtered.Count, MaxVisibleRows));
            _list.Height = rows * S(RowH);
            _list.Visible = _filtered.Count > 0;
            _empty.Top = _list.Top;
            _empty.Height = S(RowH);
            _empty.Visible = _filtered.Count == 0;

            _footer.Top = _list.Top + rows * S(RowH) + S(Pad);
            Height = _footer.Top + S(RowH) + S(Pad);

            // Select the active document (tick), otherwise the first row
            int sel = 0;
            for (int i = 0; i < _filtered.Count; i++)
            {
                if (_filtered[i].IsActive) { sel = i; break; }
            }
            if (_filtered.Count > 0) _list.SelectedIndex = sel;
        }

        private void ActivateRow(int index)
        {
            if (index < 0 || index >= _filtered.Count) return;
            int hwnd = _filtered[index].Hwnd;
            try { Close(); } catch { }
            try
            {
                if (_onActivate != null) _onActivate(hwnd);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Popup activate failed", ex);
            }
        }

        private void DrawRow(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _filtered.Count) return;
            var t = _filtered[e.Index];
            var g = e.Graphics;
            bool selected = (e.State & DrawItemState.Selected) != 0;

            using (var bg = new SolidBrush(TabTheme.MenuBg))
                g.FillRectangle(bg, e.Bounds);

            Color rowBg = selected ? (t.IsActive ? TabTheme.MenuActiveRow : TabTheme.MenuHoverRow)
                        : (t.IsActive ? TabTheme.MenuActiveRow : TabTheme.MenuBg);
            if (rowBg != TabTheme.MenuBg)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new RectangleF(e.Bounds.X, e.Bounds.Y + 0.5f, e.Bounds.Width, e.Bounds.Height - 1f);
                using (var path = Glyphs.RoundedRect(r, SF(5)))
                using (var b = new SolidBrush(rowBg))
                    g.FillPath(b, path);
                g.SmoothingMode = SmoothingMode.None;
            }

            // A tick for the active document; for an inactive pinned one, a pin in the
            // same slot (active+pinned at once is rare, the tick wins - being pinned is
            // visible anyway from the position in the list and in the tab strip).
            if (t.IsActive)
            {
                var checkRect = new Rectangle(e.Bounds.X + S(6), e.Bounds.Y, S(18), e.Bounds.Height);
                TextRenderer.DrawText(g, "✓", _fontCheck, checkRect, TabTheme.ActiveText,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
            else if (t.IsPinned)
            {
                var pinBox = new RectangleF(e.Bounds.X + S(7), e.Bounds.Y + (e.Bounds.Height - S(11)) / 2f, S(11), S(11));
                Glyphs.DrawPin(g, pinBox, TabTheme.PinGlyph);
            }

            int textLeft = e.Bounds.X + S(28);
            int textRight = e.Bounds.Right - (t.Saved ? S(10) : S(24));
            var textRect = new Rectangle(textLeft, e.Bounds.Y, Math.Max(0, textRight - textLeft), e.Bounds.Height);
            TextRenderer.DrawText(g, t.Caption ?? string.Empty, _fontText, textRect,
                t.IsProtectedView ? TabTheme.ProtectedViewText : TabTheme.MenuText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis |
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);

            // Unsaved-changes dot
            if (!t.Saved)
            {
                int d = S(6);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var dot = new SolidBrush(TabTheme.UnsavedDot))
                    g.FillEllipse(dot, e.Bounds.Right - S(16), e.Bounds.Y + (e.Bounds.Height - d) / 2, d, d);
                g.SmoothingMode = SmoothingMode.None;
            }
        }

        private void PaintSearchBorder(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, _searchHost.Width - 1f, _searchHost.Height - 1f);
            using (var path = Glyphs.RoundedRect(r, SF(5)))
            using (var pen = new Pen(TabTheme.MenuSearchBorder, 1f))
                g.DrawPath(pen, path);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            // 1 px menu border #D5D8DD (the corners are rounded by DWM)
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (var path = Glyphs.RoundedRect(r, SF(8)))
            using (var pen = new Pen(TabTheme.MenuBorder, 1f))
                g.DrawPath(pen, path);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Escape:
                    Close();
                    return true;
                case Keys.Enter:
                    ActivateRow(_list.SelectedIndex);
                    return true;
                case Keys.Down:
                    if (_list.Items.Count > 0)
                        _list.SelectedIndex = Math.Min(_list.SelectedIndex + 1, _list.Items.Count - 1);
                    return true;
                case Keys.Up:
                    if (_list.Items.Count > 0)
                        _list.SelectedIndex = Math.Max(_list.SelectedIndex - 1, 0);
                    return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            try { Close(); } catch { }
        }

        /// <summary>
        /// A managed exception escaping a window procedure into Word message loop
        /// brings WINWORD down (0xE0434352) - always swallow it and log.
        /// </summary>
        protected override void WndProc(ref Message m)
        {
            try
            {
                base.WndProc(ref m);
            }
            catch (Exception ex)
            {
                LoggingService.Error("TabListPopup WndProc failed (msg=0x" + m.Msg.ToString("X4") + ")", ex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            // A modeless form disposes itself after Close(); the fonts are released
            // here rather than from OnFormClosed (releasing from the form own close
            // event is dangerous: the closing code still touches form objects).
            if (disposing)
            {
                try { _fontText.Dispose(); } catch { }
                try { _fontCheck.Dispose(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
