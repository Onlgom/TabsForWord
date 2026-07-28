using System;
using System.Collections.Generic;

namespace TabsForWord
{
    /// <summary>
    /// A snapshot of one tab (one Word window or one Protected View window).
    /// A plain POCO with no COM references: RCW objects are never cached beyond a single Reconcile.
    /// </summary>
    public sealed class DocumentTabModel
    {
        public int Hwnd;
        public string Caption;        // tab text (document name; with 2+ windows "Name:N")
        public string FullPath;       // full path, or the name for unsaved documents
        public bool Saved;            // false => the unsaved indicator is drawn
        public bool IsActive;
        public bool IsProtectedView;
        public System.Drawing.Color? TabColor;  // user colour of the tab (null = the ordinary look)
        public bool IsPinned;         // pinned (always among the first tabs, see DocumentWindowManager)

        /// <summary>
        /// A copy with the same fields. Used wherever one field has to change on some
        /// of the tabs (SetTabColor/SetTabPinned): mutating a field IN PLACE on an
        /// object already handed to the strips by a previous PushTabs is not allowed
        /// - TabStripControl.UpdateTabs compares the new list with its previous one
        /// by value (SameAs), and if it is the very same (already mutated) object the
        /// comparison finds it "equal to itself" and skips the repaint. A fresh clone
        /// carrying the new value lets SameAs see the difference properly.
        /// </summary>
        public DocumentTabModel Clone()
        {
            return new DocumentTabModel
            {
                Hwnd = Hwnd,
                Caption = Caption,
                FullPath = FullPath,
                Saved = Saved,
                IsActive = IsActive,
                IsProtectedView = IsProtectedView,
                TabColor = TabColor,
                IsPinned = IsPinned
            };
        }

        public bool SameAs(DocumentTabModel o)
        {
            return o != null
                && Hwnd == o.Hwnd
                && string.Equals(Caption, o.Caption, StringComparison.Ordinal)
                && string.Equals(FullPath, o.FullPath, StringComparison.Ordinal)
                && Saved == o.Saved
                && IsActive == o.IsActive
                && IsProtectedView == o.IsProtectedView
                && Nullable.Equals(TabColor, o.TabColor)
                && IsPinned == o.IsPinned;
        }

        public static bool ListsEqual(IList<DocumentTabModel> a, IList<DocumentTabModel> b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
            {
                if (!a[i].SameAs(b[i])) return false;
            }
            return true;
        }
    }
}
