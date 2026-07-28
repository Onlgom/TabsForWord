using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace TabsForWord
{
    /// <summary>
    /// Store of user tab colours.
    ///
    /// Saved documents: the key is the full file path (case-insensitive), the colours
    /// survive closing the document and restarting Word - they are kept in
    /// %LOCALAPPDATA%\TabsForWord\tab-colors.cfg.
    /// Unsaved documents: the key is the window handle, in memory only
    /// (the colour lives until the document is saved or closed - a path appears on save).
    ///
    /// File line format: RRGGBB|full\path\to\file
    /// (colour first: a path may contain any character except a line break).
    /// Runs strictly on Word main STA thread, so no locking is needed.
    /// </summary>
    public sealed class TabColorStore
    {
        private readonly string _filePath;   // null => memory only (unit tests)
        private readonly Dictionary<string, Color> _byPath =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, Color> _byHwnd = new Dictionary<int, Color>();
        private bool _loaded;

        public TabColorStore() : this(DefaultFilePath) { }

        public TabColorStore(string filePath)
        {
            _filePath = filePath;
        }

        public static string DefaultFilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "tab-colors.cfg");
            }
        }

        /// <summary>Tab colour: by path for saved documents, by hwnd for unsaved ones.</summary>
        public Color? Resolve(string fullPath, int hwnd)
        {
            EnsureLoaded();
            Color c;
            if (IsPersistentKey(fullPath))
                return _byPath.TryGetValue(fullPath, out c) ? c : (Color?)null;
            return _byHwnd.TryGetValue(hwnd, out c) ? c : (Color?)null;
        }

        /// <summary>Sets (or clears, with null) the colour; for saved documents writes the file at once.</summary>
        public void SetColor(string fullPath, int hwnd, Color? color)
        {
            EnsureLoaded();
            if (IsPersistentKey(fullPath))
            {
                if (color.HasValue) _byPath[fullPath] = color.Value;
                else _byPath.Remove(fullPath);
                Save();
            }
            else
            {
                if (color.HasValue) _byHwnd[hwnd] = color.Value;
                else _byHwnd.Remove(hwnd);
            }
        }

        /// <summary>The window is gone - release the session colour of an unsaved document.</summary>
        public void ForgetWindow(int hwnd)
        {
            _byHwnd.Remove(hwnd);
        }

        /// <summary>
        /// The document has acquired a path (first save) or changed it (Save As):
        /// the colour follows the tab. Without this, a colour chosen before saving
        /// would vanish right after Ctrl+S (the key moves from hwnd to path).
        /// On Save As the previous file keeps its colour (a copy, not a move).
        /// </summary>
        public void HandlePathChanged(int hwnd, string oldFullPath, string newFullPath)
        {
            EnsureLoaded();
            if (!IsPersistentKey(newFullPath)) return;
            if (string.Equals(oldFullPath ?? string.Empty, newFullPath, StringComparison.OrdinalIgnoreCase))
                return;

            Color c;
            bool had;
            if (!IsPersistentKey(oldFullPath))
            {
                had = _byHwnd.TryGetValue(hwnd, out c);
                if (had) _byHwnd.Remove(hwnd); // the path becomes the key
            }
            else
            {
                had = _byPath.TryGetValue(oldFullPath, out c);
            }
            if (!had) return;

            _byPath[newFullPath] = c;
            Save();
            LoggingService.Info("Tab color migrated to new path (hwnd=" + hwnd + ")");
        }

        /// <summary>The key is persistent when the path is absolute (the document is on disk).</summary>
        internal static bool IsPersistentKey(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath)) return false;
            try { return Path.IsPathRooted(fullPath); }
            catch (ArgumentException) { return false; } // invalid characters in the name
        }

        // ------------------------------------------------------------------
        // Parsing and writing (pure methods - tested without the file system)
        // ------------------------------------------------------------------

        internal static Dictionary<string, Color> ParseLines(IEnumerable<string> lines)
        {
            var result = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
            if (lines == null) return result;
            foreach (var raw in lines)
            {
                if (raw == null) continue;
                var line = raw.Trim();
                // Only "colour|path" is an entry; anything else (comments "#...",
                // junk) is ignored. There is no separate comment check on purpose:
                // a colour may itself start with a hash ("#RRGGBB|path").
                int sep = line.IndexOf('|');
                if (sep <= 0 || sep == line.Length - 1) continue;

                Color? color = ParseRgb(line.Substring(0, sep).Trim());
                string path = line.Substring(sep + 1).Trim();
                if (color == null || path.Length == 0) continue;
                result[path] = color.Value; // repeated path - the last line wins
            }
            return result;
        }

        internal static List<string> FormatLines(IDictionary<string, Color> byPath)
        {
            var lines = new List<string> { "# TabsForWord: tab colours (RRGGBB|full file path)" };
            foreach (var kv in byPath)
                lines.Add(FormatRgb(kv.Value) + "|" + kv.Key);
            return lines;
        }

        internal static Color? ParseRgb(string value)
        {
            if (string.IsNullOrEmpty(value)) return null;
            var s = value.Trim();
            if (s.StartsWith("#")) s = s.Substring(1);
            if (s.Length != 6) return null;
            int rgb;
            if (!int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb)) return null;
            return Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        }

        internal static string FormatRgb(Color c)
        {
            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
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
                var parsed = ParseLines(File.ReadAllLines(_filePath));
                foreach (var kv in parsed) _byPath[kv.Key] = kv.Value;
                LoggingService.Info("Tab colors loaded: " + _byPath.Count);
            }
            catch (Exception ex)
            {
                // A corrupt file must not break the add-in - carry on without colours.
                LoggingService.Error("Tab colors load failed", ex);
            }
        }

        private void Save()
        {
            if (_filePath == null) return;
            try
            {
                var dir = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(_filePath, FormatLines(_byPath), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Tab colors save failed", ex);
            }
        }
    }
}
