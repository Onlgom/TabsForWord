using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace TabsForWord
{
    /// <summary>
    /// User tab size: a multiplier on top of the automatic DPI scale
    /// (final strip scale = DPI/96 x this multiplier).
    /// Stored in %LOCALAPPDATA%\TabsForWord\tab-size.cfg (a scale=1.15 line),
    /// applied in every Word window and surviving a restart.
    /// Runs strictly on Word main STA thread, so no locking is needed.
    /// </summary>
    public static class TabSizeSettings
    {
        public const float Min = 0.7f;
        public const float Max = 2.0f;
        public const float Default = 1.0f;

        private static float? _current;

        public static string FilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "tab-size.cfg");
            }
        }

        /// <summary>Current multiplier (lazily loaded from file; on failure => 1.0).</summary>
        public static float Current
        {
            get
            {
                if (_current.HasValue) return _current.Value;
                float v = Default;
                try
                {
                    if (File.Exists(FilePath)) v = Parse(File.ReadAllLines(FilePath));
                }
                catch (Exception ex)
                {
                    LoggingService.Error("Tab size load failed", ex);
                }
                _current = v;
                return v;
            }
        }

        /// <summary>Saves the multiplier (and remembers it as current).</summary>
        public static void Save(float scale)
        {
            float v = Clamp(scale);
            _current = v;
            try
            {
                var dir = Path.GetDirectoryName(FilePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(FilePath, new[]
                {
                    "# TabsForWord: tab size (multiplier on top of the DPI scale)",
                    "scale=" + v.ToString("0.###", CultureInfo.InvariantCulture)
                }, Encoding.UTF8);
                LoggingService.Info("Tab size scale saved: " + v.ToString("0.###", CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                LoggingService.Error("Tab size save failed", ex);
            }
        }

        /// <summary>Test-only: drop the cache so the file is read again.</summary>
        internal static void ResetCacheForTests()
        {
            _current = null;
        }

        /// <summary>Pure parser of the file lines (unit-tested without the file system).</summary>
        internal static float Parse(string[] lines)
        {
            if (lines == null) return Default;
            foreach (var raw in lines)
            {
                if (raw == null) continue;
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                if (!line.Substring(0, eq).Trim().Equals("scale", StringComparison.OrdinalIgnoreCase)) continue;

                float v;
                if (float.TryParse(line.Substring(eq + 1).Trim(),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                    return Clamp(v);
            }
            return Default;
        }

        internal static float Clamp(float v)
        {
            if (float.IsNaN(v) || float.IsInfinity(v)) return Default;
            if (v < Min) return Min;
            if (v > Max) return Max;
            return v;
        }
    }
}
