using System;
using System.Drawing;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// Pure geometry of the in-window host: no Win32 calls, fully covered by unit
    /// tests. Every method works in the single coordinate space named in its
    /// contract.
    /// </summary>
    internal static class NativeHostLayoutMath
    {
        /// <summary>
        /// The strip rectangle in the CLIENT coordinates of the Word window.
        /// The strip is pinned to the top edge of the document area (contentTopClientY)
        /// and spans the HORIZONTAL RANGE of the document area rather than the whole
        /// window width: Word side panes (Navigation, Editor, Transcribe - all of them
        /// MsoCommandBarDock) start at the natural top of the document area, and a
        /// full-width strip would cover their upper part. The document area (_WwF)
        /// narrows when a side pane opens, and the strip follows it (verified with a
        /// window-tree dump of a real Word).
        /// Returns Rectangle.Empty when the strip cannot sensibly be placed.
        /// </summary>
        internal static Rectangle ComputePanelRect(Size parentClient, int contentTopClientY, int panelHeight,
            int contentLeftClientX, int contentWidth)
        {
            if (parentClient.Width <= 0 || parentClient.Height <= 0) return Rectangle.Empty;
            if (panelHeight <= 0) return Rectangle.Empty;

            // The strip must not spill outside the client area.
            int maxTop = parentClient.Height - panelHeight;
            if (maxTop < 0) return Rectangle.Empty; // the window is shorter than the strip

            int y = contentTopClientY;
            if (y < 0) y = 0;
            if (y > maxTop) y = maxTop;

            int x = contentLeftClientX;
            if (x < 0) x = 0;
            if (x >= parentClient.Width) return Rectangle.Empty;
            int w = contentWidth;
            if (w > parentClient.Width - x) w = parentClient.Width - x;
            if (w <= 0) return Rectangle.Empty;

            return new Rectangle(x, y, w, panelHeight);
        }

        /// <summary>The simple case: the document area spans the whole window width.</summary>
        internal static Rectangle ComputePanelRect(Size parentClient, int contentTopClientY, int panelHeight)
        {
            return ComputePanelRect(parentClient, contentTopClientY, panelHeight, 0, parentClient.Width);
        }

        /// <summary>
        /// Converts a height measured at one DPI into another DPI.
        /// Used while the DPI of the WinForms control has not caught up with the DPI of
        /// the Word window yet (right after the window is moved to another monitor, say).
        /// </summary>
        internal static int ScaleHeight(int height, int fromDpi, int toDpi)
        {
            if (height <= 0) return 0;
            if (fromDpi <= 0 || toDpi <= 0 || fromDpi == toDpi) return height;
            return (int)Math.Round((double)height * toDpi / fromDpi);
        }

        /// <summary>Scales a threshold or size from the 96-DPI base into the actual DPI.</summary>
        internal static int Scale96(int value96, int dpi)
        {
            if (dpi <= 0 || dpi == 96) return value96;
            return (int)Math.Round((double)value96 * dpi / 96.0);
        }

        /// <summary>
        /// The top of the content zone (in screen coordinates) is plausible when it lies
        /// inside the client area of the window and still leaves some room for both the
        /// strip and the document.
        /// </summary>
        internal static bool IsPlausibleContentTop(int contentTopScreenY, int clientTopScreenY,
            int clientBottomScreenY, int minRemainingHeight)
        {
            if (contentTopScreenY < clientTopScreenY) return false;
            if (contentTopScreenY > clientBottomScreenY - minRemainingHeight) return false;
            return true;
        }
    }
}
