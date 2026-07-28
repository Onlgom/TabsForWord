using System;
using System.Globalization;
using System.IO;
using System.Text;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord
{
    /// <summary>Interface language of the tab strip.</summary>
    public enum UiLang
    {
        /// <summary>Neutral language; also the fallback for every locale we do not translate.</summary>
        En,
        Ru
    }

    /// <summary>
    /// Picks the interface language and remembers an explicit user choice.
    ///
    /// Detection order:
    ///   1. lang=en|ru in %LOCALAPPDATA%\TabsForWord\language.cfg (lang=auto means "detect");
    ///   2. the UI language of Word itself (LanguageSettings, msoLanguageIDUI) - the right
    ///      signal for an Office add-in: an English Windows with a Russian Word must show
    ///      Russian tabs, and Windows locale alone would get that wrong;
    ///   3. Windows UI culture, if Word refuses to answer.
    /// Anything that is not Russian resolves to English (see <see cref="FromLcid"/>).
    ///
    /// Adding a language later: extend UiLang, FromLcid and the optional parameters of
    /// Strings.T - nothing else knows about languages.
    /// </summary>
    public static class UiLanguage
    {
        public const UiLang Default = UiLang.En;

        private static UiLang? _current;
        private static bool _explicitChoice;   // language.cfg holds a real language, not "auto"

        // Kept only to re-detect when the user switches back to "same as Word".
        // Cleared in Shutdown() so the add-in does not hold Word alive.
        private static Word.Application _word;

        public static string FilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "language.cfg");
            }
        }

        /// <summary>Language in use. Falls back to the neutral one until Init() runs.</summary>
        public static UiLang Current
        {
            get { return _current.HasValue ? _current.Value : Default; }
        }

        /// <summary>True when the user pinned a language explicitly (auto-detection is off).</summary>
        public static bool IsExplicit { get { return _explicitChoice; } }

        /// <summary>
        /// Resolves the language once, at add-in start. Never throws: any failure leaves
        /// the neutral language, which is always safe to draw.
        /// </summary>
        public static void Init(Word.Application word)
        {
            _word = word;
            try
            {
                // Detection starts from scratch every time: Init also runs when the user
                // switches back to "same as Word", and a stale "explicit" flag would then
                // make the settings window show a pinned language that is no longer set.
                _explicitChoice = false;

                UiLang? fromFile = ParseFile();
                if (fromFile.HasValue)
                {
                    _explicitChoice = true;
                    Apply(fromFile.Value, "language.cfg");
                    return;
                }

                int lcid = DetectWordUiLcid(word);
                if (lcid > 0)
                {
                    Apply(FromLcid(lcid), "Word UI language (LCID " +
                        lcid.ToString(CultureInfo.InvariantCulture) + ")");
                    return;
                }

                Apply(FromLcid(CultureInfo.CurrentUICulture.LCID), "Windows UI culture " +
                    CultureInfo.CurrentUICulture.Name);
            }
            catch (Exception ex)
            {
                LoggingService.Error("UI language detection failed; using " + Default, ex);
                Apply(Default, "fallback");
            }
        }

        private static void Apply(UiLang lang, string reason)
        {
            _current = lang;
            Strings.Use(lang);
            LoggingService.Info("UI language: " + lang + " (" + reason + ")");
        }

        /// <summary>Releases the Word reference taken in Init (add-in shutdown).</summary>
        public static void Shutdown()
        {
            _word = null;
        }

        /// <summary>Saves an explicit choice (null = "same as Word") and applies it at once.</summary>
        public static void Save(UiLang? lang)
        {
            try
            {
                var dir = Path.GetDirectoryName(FilePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(FilePath, new[]
                {
                    "# TabsForWord: interface language",
                    "# lang=auto (follow Word), en, ru",
                    "lang=" + (lang.HasValue ? lang.Value.ToString().ToLowerInvariant() : "auto")
                }, Encoding.UTF8);
                LoggingService.Info("UI language setting saved: " +
                    (lang.HasValue ? lang.Value.ToString() : "auto"));
            }
            catch (Exception ex)
            {
                LoggingService.Error("UI language setting save failed", ex);
            }

            _explicitChoice = lang.HasValue;
            if (lang.HasValue)
            {
                Apply(lang.Value, "user choice");
            }
            else
            {
                // Back to auto: redo the same detection the add-in does at startup.
                _current = null;
                Init(_word);
            }
        }

        /// <summary>Russian for any Russian locale (ru-RU, ru-MD, ...); English for everything else.</summary>
        internal static UiLang FromLcid(int lcid)
        {
            const int LangRussian = 0x19;
            return (lcid & 0x3FF) == LangRussian ? UiLang.Ru : UiLang.En;
        }

        /// <summary>Reads lang= from the config. Returns null for "auto", a bad value or no file.</summary>
        internal static UiLang? ParseFile()
        {
            try
            {
                if (!File.Exists(FilePath)) return null;
                return Parse(File.ReadAllLines(FilePath));
            }
            catch (Exception ex)
            {
                LoggingService.Warn("language.cfg read failed: " + ex.Message);
                return null;
            }
        }

        /// <summary>Pure parser (unit-tested without the file system).</summary>
        internal static UiLang? Parse(string[] lines)
        {
            if (lines == null) return null;
            foreach (var raw in lines)
            {
                if (raw == null) continue;
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                if (!line.Substring(0, eq).Trim().Equals("lang", StringComparison.OrdinalIgnoreCase)) continue;

                var v = line.Substring(eq + 1).Trim().ToLowerInvariant();
                if (v == "ru" || v == "russian") return UiLang.Ru;
                if (v == "en" || v == "english") return UiLang.En;
                return null;   // "auto" and anything unknown: detect
            }
            return null;
        }

        /// <summary>
        /// UI language of Word itself. Returns 0 when Word cannot answer (old build,
        /// COM failure) so the caller can fall back to the Windows culture.
        /// </summary>
        private static int DetectWordUiLcid(Word.Application word)
        {
            if (word == null) return 0;
            try
            {
                return word.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            }
            catch (Exception ex)
            {
                LoggingService.Warn("Word UI language unavailable: " + ex.Message);
                return 0;
            }
        }

        /// <summary>Test hook: force a language without touching disk or Word.</summary>
        internal static void OverrideForTests(UiLang lang)
        {
            _current = lang;
            _explicitChoice = true;
            Strings.Use(lang);
        }
    }
}
