using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TabsForWord
{
    /// <summary>
    /// Store of pinned tabs and of their order across Word sessions.
    ///
    /// One file: line order = tab order (pinned tabs are written first - the calling
    /// code guarantees that, see DocumentWindowManager), and a pinned line carries the
    /// "PIN|" prefix. Works for saved documents only (the key is the full file path,
    /// case-insensitive) - just like the tab colour (TabColorStore), because unsaved
    /// documents have no stable identity between Word runs.
    ///
    ///
    /// Pinning an unsaved document lives until it is saved or closed - per session, by
    /// hwnd, in memory only. The order of unsaved documents is not restored (there
    /// would be no point: the next Word run will not see them).
    ///
    ///
    /// Runs strictly on Word main STA thread, so no locking is needed.
    /// </summary>
    public sealed class TabOrderStore
    {
        private readonly string _filePath;   // null => memory only (unit tests)
        private List<string> _orderedPaths = new List<string>();       // persistent order (saved documents)
        private readonly HashSet<string> _pinnedPaths =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<int> _pinnedHwnds = new HashSet<int>(); // per-session pinning of unsaved ones
        private bool _loaded;

        public TabOrderStore() : this(DefaultFilePath) { }

        public TabOrderStore(string filePath)
        {
            _filePath = filePath;
        }

        public static string DefaultFilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "tab-order.cfg");
            }
        }

        /// <summary>Whether a tab is pinned: by path for saved documents, by hwnd for unsaved ones.</summary>
        public bool IsPinned(string fullPath, int hwnd)
        {
            EnsureLoaded();
            return TabColorStore.IsPersistentKey(fullPath)
                ? _pinnedPaths.Contains(fullPath)
                : _pinnedHwnds.Contains(hwnd);
        }

        /// <summary>
        /// Sets the pinned flag in memory. For saved documents the file is NOT written
        /// here - the calling code (DocumentWindowManager) immediately recomputes the
        /// order and calls SaveOrder with the full current tab list (pinning and order
        /// are always written together - writing them separately would let the two
        /// drift apart).
        /// </summary>
        public void SetPinned(string fullPath, int hwnd, bool pinned)
        {
            EnsureLoaded();
            if (TabColorStore.IsPersistentKey(fullPath))
            {
                if (pinned) _pinnedPaths.Add(fullPath);
                else _pinnedPaths.Remove(fullPath);
            }
            else
            {
                if (pinned) _pinnedHwnds.Add(hwnd);
                else _pinnedHwnds.Remove(hwnd);
            }
        }

        /// <summary>The window is gone - release the session pinning of an unsaved document.</summary>
        public void ForgetWindow(int hwnd)
        {
            _pinnedHwnds.Remove(hwnd);
        }

        /// <summary>
        /// The document has acquired a path (first save) or changed it (Save As): the
        /// pinning follows the tab, otherwise it would vanish right after Ctrl+S (the key
        /// moves from hwnd to path). The order needs no separate migration: it is fully
        /// recomputed from the current tab list on the next SaveOrder (which uses the
        /// already updated paths).
        /// </summary>
        public void HandlePathChanged(int hwnd, string oldFullPath, string newFullPath)
        {
            EnsureLoaded();
            if (!TabColorStore.IsPersistentKey(newFullPath)) return;
            if (string.Equals(oldFullPath ?? string.Empty, newFullPath, StringComparison.OrdinalIgnoreCase))
                return;

            bool wasPinned;
            if (!TabColorStore.IsPersistentKey(oldFullPath))
            {
                wasPinned = _pinnedHwnds.Contains(hwnd);
                if (wasPinned) _pinnedHwnds.Remove(hwnd);
            }
            else
            {
                wasPinned = _pinnedPaths.Contains(oldFullPath);
            }
            if (!wasPinned) return;

            _pinnedPaths.Add(newFullPath);
            LoggingService.Info("Tab pin migrated to new path (hwnd=" + hwnd + ")");
        }

        /// <summary>Rank of a known path in the stored order; null when the path is unknown (a new document).</summary>
        public int? RankOf(string fullPath)
        {
            EnsureLoaded();
            if (!TabColorStore.IsPersistentKey(fullPath)) return null;
            int i = _orderedPaths.FindIndex(p => string.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase));
            return i < 0 ? (int?)null : i;
        }

        /// <summary>
        /// Saves the order and the pinning of saved documents in the currently displayed
        /// order (unsaved documents never reach the file).
        ///
        /// Paths missing from tabsInOrder (the document is closed right now) KEEP their
        /// previous position and pinning in the file - otherwise closing ALL tabs before
        /// quitting Word (the ordinary Quit path) would call SaveOrder with an empty list
        /// and wipe the pinning of documents the user opens next time (the same principle
        /// as in TabColorStore: the colour of a closed document does not disappear).
        /// </summary>
        public void SaveOrder(IEnumerable<DocumentTabModel> tabsInOrder)
        {
            EnsureLoaded();
            var paths = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in tabsInOrder)
            {
                if (!TabColorStore.IsPersistentKey(t.FullPath)) continue;
                if (!seen.Add(t.FullPath)) continue; // the document is open in 2+ windows - the same path
                paths.Add(t.FullPath);
                if (t.IsPinned) _pinnedPaths.Add(t.FullPath);
                else _pinnedPaths.Remove(t.FullPath);
            }
            foreach (var p in _orderedPaths)
            {
                if (seen.Add(p)) paths.Add(p);
            }
            _orderedPaths = paths;
            Save();
        }

        // ------------------------------------------------------------------
        // Parsing and writing (pure methods - tested without the file system)
        // ------------------------------------------------------------------

        internal struct Entry
        {
            public string Path;
            public bool Pinned;
        }

        internal static List<Entry> ParseLines(IEnumerable<string> lines)
        {
            var result = new List<Entry>();
            if (lines == null) return result;
            foreach (var raw in lines)
            {
                if (raw == null) continue;
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;

                bool pinned = false;
                if (line.StartsWith("PIN|", StringComparison.OrdinalIgnoreCase))
                {
                    pinned = true;
                    line = line.Substring(4);
                }
                if (line.Length == 0) continue;
                result.Add(new Entry { Path = line, Pinned = pinned });
            }
            return result;
        }

        internal static List<string> FormatLines(List<Entry> entries)
        {
            var lines = new List<string>
            {
                "# TabsForWord: tab order and pinned tabs " +
                "(line order = tab order; PIN|path = pinned)"
            };
            foreach (var e in entries)
                lines.Add((e.Pinned ? "PIN|" : string.Empty) + e.Path);
            return lines;
        }

        // ------------------------------------------------------------------
        // File
        // ------------------------------------------------------------------

        private void EnsureLoaded()
        {
            if (_loaded) return;
            _loaded = true;
            if (_filePath == null) return;
            try
            {
                if (!File.Exists(_filePath)) return;
                var entries = ParseLines(File.ReadAllLines(_filePath));
                var paths = new List<string>();
                foreach (var e in entries)
                {
                    paths.Add(e.Path);
                    if (e.Pinned) _pinnedPaths.Add(e.Path);
                }
                _orderedPaths = paths;
                LoggingService.Info("Tab order loaded: " + _orderedPaths.Count +
                    " (pinned: " + _pinnedPaths.Count + ")");
            }
            catch (Exception ex)
            {
                // A corrupt file must not break the add-in - carry on without a stored order.
                LoggingService.Error("Tab order load failed", ex);
            }
        }

        private void Save()
        {
            if (_filePath == null) return;
            try
            {
                var dir = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                var entries = _orderedPaths
                    .Select(p => new Entry { Path = p, Pinned = _pinnedPaths.Contains(p) })
                    .ToList();
                File.WriteAllLines(_filePath, FormatLines(entries), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Tab order save failed", ex);
            }
        }
    }
}
