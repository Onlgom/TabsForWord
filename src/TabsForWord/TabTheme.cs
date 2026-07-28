using System;
using System.Drawing;

namespace TabsForWord
{
    /// <summary>
    /// Tab strip palette from the "Tab specification 2b" (July 2026), section 3.
    /// Every colour is a constant: the strip mirrors the light Word theme and does
    /// not depend on system colours (the specification fixes exact values).
    /// </summary>
    internal static class TabTheme
    {
        // Strip
        public static readonly Color PanelBg = Color.FromArgb(0xF3, 0xF4, 0xF5);
        public static readonly Color PanelBorder = Color.FromArgb(0xCF, 0xD0, 0xD2); // bottom border and active-tab outline
        public static readonly Color BelowPanelBg = Color.FromArgb(0xFD, 0xFD, 0xFD); // background of the area under the strip

        // Active tab
        public static readonly Color ActiveBg = Color.FromArgb(0xFD, 0xFD, 0xFD);
        public static readonly Color ActiveText = Color.FromArgb(0x18, 0x5A, 0xBD);

        // Inactive tab
        public static readonly Color InactiveBg = Color.FromArgb(0xEB, 0xED, 0xEF);
        public static readonly Color InactiveBorder = Color.FromArgb(0xD9, 0xDB, 0xDE);
        public static readonly Color InactiveText = Color.FromArgb(0x5D, 0x60, 0x65);
        public static readonly Color InactiveHoverBg = Color.FromArgb(0xE3, 0xE5, 0xE8);

        // Unsaved-changes indicators
        public static readonly Color UnsavedEdge = Color.FromArgb(0xDF, 0xA2, 0x4B); // 3 px edge
        public static readonly Color UnsavedDot = Color.FromArgb(0xC0, 0x7A, 0x1D);  // dot, 6 px across

        // Close button
        public static readonly Color CloseGlyph = Color.FromArgb(0x85, 0x88, 0x8D);
        public static readonly Color CloseHoverBg = Color.FromArgb(0xE8, 0xE8, 0xE8);
        public static readonly Color CloseHoverGlyph = Color.FromArgb(0x33, 0x33, 0x33);

        // The "+", all-tabs, collapse and scroll buttons
        public static readonly Color ButtonGlyph = Color.FromArgb(0x5F, 0x63, 0x6A);
        public static readonly Color ButtonGlyphDisabled = Color.FromArgb(0xB0, 0xB4, 0xBA);
        public static readonly Color ButtonHoverBg = Color.FromArgb(0xE9, 0xEA, 0xEC);
        public static readonly Color ButtonPressedBg = Color.FromArgb(0xDF, 0xE3, 0xE8);

        // Separator before the right-hand button block
        public static readonly Color Separator = Color.FromArgb(0xD6, 0xD9, 0xDD);

        // Protected View - muted text (not in the specification; kept from the first version)
        public static readonly Color ProtectedViewText = Color.FromArgb(0x8A, 0x8F, 0x96);

        // Pinned-tab glyph (the pin) - the same muted tone as ProtectedViewText
        public static readonly Color PinGlyph = Color.FromArgb(0x8A, 0x8F, 0x96);

        // All-tabs menu
        public static readonly Color MenuBg = Color.White;
        public static readonly Color MenuBorder = Color.FromArgb(0xD5, 0xD8, 0xDD);
        public static readonly Color MenuActiveRow = Color.FromArgb(0xEE, 0xF4, 0xFC);
        public static readonly Color MenuHoverRow = Color.FromArgb(0xF3, 0xF4, 0xF6);
        public static readonly Color MenuText = Color.FromArgb(0x26, 0x29, 0x2E);
        public static readonly Color MenuSearchBorder = Color.FromArgb(0xD5, 0xD8, 0xDD);

        // Keyboard focus
        public static readonly Color FocusFrame = Color.FromArgb(0x18, 0x5A, 0xBD);

        // Shadow of a dragged tab: rgba(20,26,36,.2)
        public static readonly Color DragShadow = Color.FromArgb(51, 20, 26, 36);

        // ------------------------------------------------------------------
        // User tab colour: pastel derivatives.
        // The tones are deliberately light (>=60% white) so that the amber edge and
        // the unsaved dot stay visible on top of any colour.
        // ------------------------------------------------------------------

        /// <summary>Blends a with b: weightB is the share of b (0..1).</summary>
        public static Color Blend(Color a, Color b, float weightB)
        {
            if (weightB < 0f) weightB = 0f;
            if (weightB > 1f) weightB = 1f;
            return Color.FromArgb(
                (int)Math.Round(a.R + (b.R - a.R) * weightB),
                (int)Math.Round(a.G + (b.G - a.G) * weightB),
                (int)Math.Round(a.B + (b.B - a.B) * weightB));
        }

        /// <summary>Fill of an active coloured tab (the lightest one - text on top stays readable).</summary>
        public static Color TintActiveBg(Color c) { return Blend(c, Color.White, 0.82f); }

        /// <summary>Fill of an inactive coloured tab.</summary>
        public static Color TintInactiveBg(Color c) { return Blend(c, Color.White, 0.72f); }

        /// <summary>Fill of an inactive coloured tab on hover (slightly richer).</summary>
        public static Color TintInactiveHoverBg(Color c) { return Blend(c, Color.White, 0.62f); }

        /// <summary>Border of a coloured tab.</summary>
        public static Color TintBorder(Color c) { return Blend(c, Color.White, 0.45f); }
    }
}
