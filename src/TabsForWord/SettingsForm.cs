using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TabsForWord
{
    /// <summary>
    /// Settings window: tab size, Ctrl+Tab interception and interface language.
    /// Opened by the "Settings..." button in the footer of the all-tabs menu, on the
    /// same anchor. Styled like TabListPopup (the same family of strip popups):
    /// no border, rounded corners (DWM), the TabTheme palette, closes on Esc or on
    /// losing focus. Changes apply instantly through the callbacks passed in (the
    /// same event chain as the context-menu items), which is why there are no
    /// separate OK/Cancel buttons.
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private const int FormW = 300;
        private const int Pad = 14;
        private const int RowH = 28;

        private readonly float _scale;
        private readonly Action<float> _onSizeChange;
        private readonly Action<bool> _onHotkeyChange;
        private readonly Action<UiLang?> _onLanguageChange;

        private readonly Font _fontHeader;
        private readonly Font _fontSection;
        private readonly Font _fontText;
        private readonly RadioButton[] _sizeRadios;
        private readonly CheckBox _hotkeyCheck;
        private readonly RadioButton[] _langRadios;
        private bool _suppressSizeEvents;   // do not fire the callback while setting the initial Checked
        private bool _suppressLangEvents;

        private static readonly float[] SizeValues = { 0.85f, 1.0f, 1.15f, 1.3f };

        // Order of the language buttons: null = "same as Word" (auto-detect), then the languages.
        private static readonly UiLang?[] LangValues = { null, UiLang.En, UiLang.Ru };

        public SettingsForm(float currentScale, bool hotkeyEnabled, float uiScale,
            Action<float> onSizeChange, Action<bool> onHotkeyChange,
            Action<UiLang?> onLanguageChange)
        {
            _scale = uiScale;
            _onSizeChange = onSizeChange;
            _onHotkeyChange = onHotkeyChange;
            _onLanguageChange = onLanguageChange;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = TabTheme.MenuBg;
            DoubleBuffered = true;
            Width = S(FormW);

            _fontHeader = new Font("Segoe UI", SF(13f), FontStyle.Bold, GraphicsUnit.Pixel);
            _fontSection = new Font("Segoe UI", SF(11f), FontStyle.Bold, GraphicsUnit.Pixel);
            _fontText = new Font("Segoe UI", SF(12.5f), FontStyle.Regular, GraphicsUnit.Pixel);

            int y = S(Pad) + S(28); // room for the header (drawn in OnPaint)

            Controls.Add(MakeSectionLabel(Strings.SettingsSectionSize, ref y));

            var sizeNames = Strings.SizeNames;
            _suppressSizeEvents = true;
            _sizeRadios = new RadioButton[SizeValues.Length];
            for (int i = 0; i < SizeValues.Length; i++)
            {
                var rb = new RadioButton
                {
                    Text = sizeNames[i],
                    Left = S(Pad),
                    Top = y,
                    Width = Width - S(Pad) * 2,
                    Height = S(RowH),
                    Font = _fontText,
                    ForeColor = TabTheme.MenuText,
                    BackColor = TabTheme.MenuBg,
                    Checked = Math.Abs(currentScale - SizeValues[i]) < 0.01f,
                    Tag = i
                };
                rb.CheckedChanged += OnSizeRadioChanged;
                Controls.Add(rb);
                _sizeRadios[i] = rb;
                y += S(RowH);
            }
            _suppressSizeEvents = false;

            y += S(8);
            Controls.Add(MakeSectionLabel(Strings.SettingsSectionKeyboard, ref y));

            _hotkeyCheck = new CheckBox
            {
                Text = Strings.SettingsCtrlTab,
                Left = S(Pad),
                Top = y,
                Width = Width - S(Pad) * 2,
                Height = S(RowH) + S(12),
                Font = _fontText,
                ForeColor = TabTheme.MenuText,
                BackColor = TabTheme.MenuBg,
                Checked = hotkeyEnabled,
                AutoSize = false
            };
            _hotkeyCheck.CheckedChanged += delegate
            {
                try { if (_onHotkeyChange != null) _onHotkeyChange(_hotkeyCheck.Checked); }
                catch (Exception ex) { LoggingService.Error("Settings hotkey toggle failed", ex); }
            };
            Controls.Add(_hotkeyCheck);
            y += _hotkeyCheck.Height + S(8);

            Controls.Add(MakeSectionLabel(Strings.SettingsSectionLanguage, ref y));

            // Language names stay in their own language (someone who switched Word to a
            // language they cannot read must still find the way back).
            var langNames = new[] { Strings.SettingsLangAuto, Strings.LangNameEn, Strings.LangNameRu };
            var currentLang = UiLanguage.IsExplicit ? (UiLang?)UiLanguage.Current : null;

            _suppressLangEvents = true;
            _langRadios = new RadioButton[LangValues.Length];
            for (int i = 0; i < LangValues.Length; i++)
            {
                var rb = new RadioButton
                {
                    Text = langNames[i],
                    Left = S(Pad),
                    Top = y,
                    Width = Width - S(Pad) * 2,
                    Height = S(RowH),
                    Font = _fontText,
                    ForeColor = TabTheme.MenuText,
                    BackColor = TabTheme.MenuBg,
                    Checked = LangValues[i].HasValue == currentLang.HasValue
                              && (!currentLang.HasValue || LangValues[i].Value == currentLang.Value),
                    Tag = i
                };
                rb.CheckedChanged += OnLangRadioChanged;
                Controls.Add(rb);
                _langRadios[i] = rb;
                y += S(RowH);
            }
            _suppressLangEvents = false;

            var hint = new Label
            {
                Text = Strings.SettingsLangRestartHint,
                Left = S(Pad),
                Top = y,
                Width = Width - S(Pad) * 2,
                Height = S(18),
                Font = _fontSection,
                ForeColor = TabTheme.ProtectedViewText,
                BackColor = TabTheme.MenuBg
            };
            Controls.Add(hint);
            y += hint.Height + S(Pad);

            Height = y;
        }

        private Label MakeSectionLabel(string text, ref int y)
        {
            var lbl = new Label
            {
                Text = text.ToUpperInvariant(),
                Left = S(Pad),
                Top = y,
                Width = Width - S(Pad) * 2,
                Height = S(18),
                Font = _fontSection,
                ForeColor = TabTheme.ProtectedViewText,
                BackColor = TabTheme.MenuBg
            };
            y += S(24);
            return lbl;
        }

        private void OnLangRadioChanged(object sender, EventArgs e)
        {
            if (_suppressLangEvents) return;
            var rb = (RadioButton)sender;
            if (!rb.Checked) return;
            try
            {
                int i = (int)rb.Tag;
                if (_onLanguageChange != null) _onLanguageChange(LangValues[i]);
                // There is nothing to repaint the settings window with - it was built with
                // the previous captions; close it, the next call opens it in the new language.
                Close();
            }
            catch (Exception ex) { LoggingService.Error("Settings language change failed", ex); }
        }

        private void OnSizeRadioChanged(object sender, EventArgs e)
        {
            if (_suppressSizeEvents) return;
            var rb = (RadioButton)sender;
            if (!rb.Checked) return;
            try
            {
                int i = (int)rb.Tag;
                if (_onSizeChange != null) _onSizeChange(SizeValues[i]);
            }
            catch (Exception ex) { LoggingService.Error("Settings size change failed", ex); }
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
        }

        /// <summary>Shows the window below the strip, right edge at the anchor (like TabListPopup.ShowBelow).</summary>
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
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (var path = Glyphs.RoundedRect(r, SF(8)))
            using (var pen = new Pen(TabTheme.MenuBorder, 1f))
                g.DrawPath(pen, path);
            g.SmoothingMode = SmoothingMode.None;

            var headerRect = new Rectangle(S(Pad), S(10), Width - S(Pad) * 2, S(20));
            TextRenderer.DrawText(g, Strings.SettingsTitle, _fontHeader, headerRect, TabTheme.MenuText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { Close(); return true; }
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
                LoggingService.Error("SettingsForm WndProc failed (msg=0x" + m.Msg.ToString("X4") + ")", ex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _fontHeader.Dispose(); } catch { }
                try { _fontSection.Dispose(); } catch { }
                try { _fontText.Dispose(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
