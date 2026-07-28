using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord
{
    /// <summary>
    /// Receiver of Reconcile results: keeps the tab strips in step with the live
    /// Word windows and hands them the current list of tabs.
    ///
    /// Implementations:
    ///  - TabPaneManager        - the classic mode (Custom Task Pane);
    ///  - NativeTabHostManager  - the in-window Win32 host (no CTP title strip),
    ///    which keeps a TabPaneManager inside as its fallback.
    ///
    /// DocumentWindowManager knows only this interface, so the modes are swapped
    /// in Connect without touching the tab model logic.
    /// </summary>
    public interface ITabPaneSync
    {
        /// <summary>Bring the strips in step with the live windows and push a tab snapshot.</summary>
        void SyncPanes(List<Word.Window> liveWindows, List<DocumentTabModel> tabs, bool tabsChanged);

        /// <summary>Push the current tab order to every live strip (after a reorder).</summary>
        void PushTabs(List<DocumentTabModel> tabs);
    }
}
