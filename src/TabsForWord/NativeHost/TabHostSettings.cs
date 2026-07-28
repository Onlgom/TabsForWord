using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace TabsForWord.NativeHost
{
    /// <summary>How the tab strip is hosted.</summary>
    internal enum TabHostMode
    {
        /// <summary>The classic (fallback) mode: a Custom Task Pane with its own title strip.</summary>
        CustomTaskPane,

        /// <summary>The main mode (since v1.6.0): our own Win32 host inside the Word window.</summary>
        Native
    }

    /// <summary>
    /// Tab strip mode settings. Read once, when the add-in starts.
    ///
    /// The file %LOCALAPPDATA%\TabsForWord\native-host.cfg holds key=value lines:
    ///   mode=native | ctp      - the strip mode (native by default);
    ///   dump=1                 - dump the Word window tree when a host is created;
    ///   reserve=1              - reserve the space (push the document area down)
    ///                            instead of covering it (1 by default);
    ///   bg=auto | spec | #RRGGBB - strip background: auto (default) matches the Word
    ///                            work area; spec is the specification 2b colour;
    ///                            #RRGGBB is a fixed colour.
    ///
    /// The WORDTABS_NATIVE_HOST=1|0 environment variable wins over mode from the file.
    /// With neither file nor variable the result is Native + reserve (decision ADR-015;
    /// before v1.6.0 the default was CustomTaskPane). Back to the classic mode:
    /// mode=ctp in the config (installer\Classic-mode.cmd). On top of that, repeated
    /// failures of the in-window host make NativeTabHostManager fall back to CTP on its
    /// own for the rest of the Word session.
    /// </summary>
    internal static class TabHostSettings
    {
        private static bool _loaded;
        private static TabHostMode _mode = TabHostMode.Native;
        private static bool _dumpWindowTree;
        private static bool _reserveSpace = true;
        private static string _background = "auto";

        internal static TabHostMode Mode { get { EnsureLoaded(); return _mode; } }
        internal static bool DumpWindowTree { get { EnsureLoaded(); return _dumpWindowTree; } }
        internal static bool ReserveSpace { get { EnsureLoaded(); return _reserveSpace; } }

        /// <summary>The bg= value: "auto", "spec" or "#RRGGBB".</summary>
        internal static string Background { get { EnsureLoaded(); return _background; } }

        internal static string ConfigPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "native-host.cfg");
            }
        }

        /// <summary>Test-only: pin the settings without reading the file or the environment.</summary>
        internal static void OverrideForTests(TabHostMode mode, bool dumpWindowTree, bool reserveSpace)
        {
            _loaded = true;
            _mode = mode;
            _dumpWindowTree = dumpWindowTree;
            _reserveSpace = reserveSpace;
            _background = "spec";
        }

        /// <summary>
        /// Parses "#RRGGBB" (case-insensitive, the hash may be omitted).
        /// Returns null for anything malformed.
        /// </summary>
        internal static Color? ParseColor(string value)
        {
            if (string.IsNullOrEmpty(value)) return null;
            var s = value.Trim();
            if (s.StartsWith("#")) s = s.Substring(1);
            if (s.Length != 6) return null;
            int rgb;
            if (!int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb)) return null;
            return Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        }

        private static void EnsureLoaded()
        {
            if (_loaded) return;
            _loaded = true;
            try
            {
                string[] lines = null;
                try
                {
                    if (File.Exists(ConfigPath)) lines = File.ReadAllLines(ConfigPath);
                }
                catch (Exception ex)
                {
                    LoggingService.Warn("native-host.cfg read failed: " + ex.Message);
                }

                string env = null;
                try { env = Environment.GetEnvironmentVariable("WORDTABS_NATIVE_HOST"); }
                catch { }

                bool dump, reserve;
                string background;
                _mode = Parse(lines, env, out dump, out reserve, out background);
                _dumpWindowTree = dump;
                _reserveSpace = reserve;
                _background = background;

                // The mode is always logged: since v1.6.0 the main mode is Native, and
                // when reading field reports one has to see that the classic CTP was
                // switched on deliberately (by config) rather than by itself.
                if (_mode == TabHostMode.Native)
                    LoggingService.Info("Tab host mode: Native (dump=" +
                        (dump ? "1" : "0") + ", reserve=" + (reserve ? "1" : "0") + ")");
                else
                    LoggingService.Info("Tab host mode: CustomTaskPane (classic, set by config/env)");
            }
            catch (Exception ex)
            {
                // An unexpected failure to read the settings => the classic mode, which is
                // guaranteed to work (it is also the emergency contour of the native host).
                LoggingService.Error("TabHostSettings load failed; using CustomTaskPane", ex);
                _mode = TabHostMode.CustomTaskPane;
                _dumpWindowTree = false;
                _reserveSpace = false;
                _background = "auto";
            }
        }

        /// <summary>The same parser without the background (used by the unit tests).</summary>
        internal static TabHostMode Parse(IEnumerable<string> configLines, string envValue,
            out bool dumpWindowTree, out bool reserveSpace)
        {
            string background;
            return Parse(configLines, envValue, out dumpWindowTree, out reserveSpace, out background);
        }

        /// <summary>
        /// Pure settings parser (unit-tested without the file system).
        /// envValue wins over mode from the file.
        /// Defaults (no file / no key): mode=native, reserve=1, dump=0, bg=auto.
        /// </summary>
        internal static TabHostMode Parse(IEnumerable<string> configLines, string envValue,
            out bool dumpWindowTree, out bool reserveSpace, out string background)
        {
            var mode = TabHostMode.Native;
            dumpWindowTree = false;
            reserveSpace = true;
            background = "auto";

            if (configLines != null)
            {
                foreach (var raw in configLines)
                {
                    if (raw == null) continue;
                    var line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    var key = line.Substring(0, eq).Trim().ToLowerInvariant();
                    var value = line.Substring(eq + 1).Trim().ToLowerInvariant();

                    switch (key)
                    {
                        case "mode":
                            // "nativeexperimental" was the mode name before v1.6.0; it is still
                            // read for configs left over from the experimental builds.
                            if (value == "native" || value == "nativeexperimental") mode = TabHostMode.Native;
                            else if (value == "ctp" || value == "customtaskpane") mode = TabHostMode.CustomTaskPane;
                            break;
                        case "dump":
                            dumpWindowTree = value == "1" || value == "true";
                            break;
                        case "reserve":
                            reserveSpace = value == "1" || value == "true";
                            break;
                        case "bg":
                            if (value == "auto" || value == "spec" || ParseColor(value) != null)
                                background = value;
                            break;
                    }
                }
            }

            if (!string.IsNullOrEmpty(envValue))
            {
                var v = envValue.Trim();
                if (v == "1") mode = TabHostMode.Native;
                else if (v == "0") mode = TabHostMode.CustomTaskPane;
            }

            return mode;
        }
    }
}
