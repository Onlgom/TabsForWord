using System;
using System.IO;
using System.Text;

namespace TabsForWord
{
    /// <summary>
    /// Whether Ctrl+Tab / Ctrl+Shift+Tab interception is on (see KeyboardHookService).
    /// On by default, as it has been from the first version; it can be switched off in the
    /// settings window, after which Word own behaviour (inserting a tab character
    /// inside a table, docs/KNOWN_ISSUES.md O8) is no longer shadowed.
    ///
    /// Stored in %LOCALAPPDATA%\TabsForWord\hotkey.cfg (an enabled=1|0 line) and
    /// survives a Word restart. Runs strictly on Word main STA thread.
    /// </summary>
    public static class HotkeySettings
    {
        public const bool Default = true;

        private static bool? _current;

        public static string FilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "hotkey.cfg");
            }
        }

        /// <summary>Current value (lazily loaded from file; on failure => enabled).</summary>
        public static bool CtrlTabEnabled
        {
            get
            {
                if (_current.HasValue) return _current.Value;
                bool v = Default;
                try
                {
                    if (File.Exists(FilePath)) v = Parse(File.ReadAllLines(FilePath));
                }
                catch (Exception ex)
                {
                    LoggingService.Error("Hotkey setting load failed", ex);
                }
                _current = v;
                return v;
            }
        }

        /// <summary>Saves the value (and remembers it as current).</summary>
        public static void Save(bool enabled)
        {
            _current = enabled;
            try
            {
                var dir = Path.GetDirectoryName(FilePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(FilePath, new[]
                {
                    "# TabsForWord: Ctrl+Tab / Ctrl+Shift+Tab interception for switching tabs",
                    "enabled=" + (enabled ? "1" : "0")
                }, Encoding.UTF8);
                LoggingService.Info("Ctrl+Tab hotkey setting saved: " + (enabled ? "1" : "0"));
            }
            catch (Exception ex)
            {
                LoggingService.Error("Hotkey setting save failed", ex);
            }
        }

        /// <summary>Test-only: drop the cache so the file is read again.</summary>
        internal static void ResetCacheForTests()
        {
            _current = null;
        }

        /// <summary>Pure parser of the file lines (unit-tested without the file system).</summary>
        internal static bool Parse(string[] lines)
        {
            if (lines == null) return Default;
            foreach (var raw in lines)
            {
                if (raw == null) continue;
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                if (!line.Substring(0, eq).Trim().Equals("enabled", StringComparison.OrdinalIgnoreCase)) continue;

                var v = line.Substring(eq + 1).Trim();
                if (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
                if (v == "0" || v.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
            }
            return Default;
        }
    }
}
