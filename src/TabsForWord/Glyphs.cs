using System.Drawing;
using System.Drawing.Drawing2D;

namespace TabsForWord
{
    /// <summary>
    /// Vector drawing shared by the tab strip (TabStripControl), the all-tabs menu
    /// (TabListPopup) and the settings window (SettingsForm), so that they never
    /// drift apart.
    /// </summary>
    internal static class Glyphs
    {
        /// <summary>
        /// A rounded rectangle as a closed path. The caller owns the path and disposes
        /// it (every call site wraps this in a using).
        /// </summary>
        public static GraphicsPath RoundedRect(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = radius * 2f;
            p.StartFigure();
            p.AddArc(r.Left, r.Top, d, d, 180f, 90f);
            p.AddArc(r.Right - d, r.Top, d, d, 270f, 90f);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0f, 90f);
            p.AddArc(r.Left, r.Bottom - d, d, d, 90f, 90f);
            p.CloseFigure();
            return p;
        }

        /// <summary>
        /// Pinned-tab glyph: a pin (round head plus a tip pointing down).
        /// Everything is proportional to box, so the same code serves the strip
        /// (larger) and a menu row (smaller) without separate magic numbers.
        /// </summary>
        public static void DrawPin(Graphics g, RectangleF box, Color color)
        {
            float cx = box.X + box.Width / 2f;
            float headR = box.Width * 0.28f;
            float headCy = box.Y + box.Height * 0.36f;
            float tipY = box.Bottom;

            var old = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = new GraphicsPath())
            {
                path.AddEllipse(cx - headR, headCy - headR, headR * 2f, headR * 2f);
                path.AddPolygon(new[]
                {
                    new PointF(cx - headR * 0.75f, headCy + headR * 0.55f),
                    new PointF(cx + headR * 0.75f, headCy + headR * 0.55f),
                    new PointF(cx, tipY)
                });
                using (var brush = new SolidBrush(color))
                    g.FillPath(brush, path);
            }
            g.SmoothingMode = old;
        }
    }
}
