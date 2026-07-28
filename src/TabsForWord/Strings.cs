using System;

namespace TabsForWord
{
    /// <summary>
    /// Every piece of text the user can see in the tab strip, in one table.
    ///
    /// Why properties and not .resx / satellite assemblies: the whole install story of
    /// this add-in is "one DLL, per-user, no administrator" - install.ps1, repair.ps1,
    /// diagnose.ps1 and the packaging script all assume a single file. Satellite
    /// assemblies would add a ru\TabsForWord.resources.dll that each of them has to
    /// learn about, and a missing satellite degrades silently to another language.
    /// With ~35 strings a compiled table is cheaper and the compiler guarantees that
    /// no translation is missing (see ADR-016).
    ///
    /// Adding a language: add it to UiLang, map it in UiLanguage.FromLcid and give T()
    /// one more optional parameter - untranslated strings fall back to English on their
    /// own, so call sites can be filled in gradually.
    ///
    /// Log messages are deliberately NOT here: they stay English always, so that a
    /// report from any machine reads the same way (see LoggingService).
    /// </summary>
    internal static class Strings
    {
        private static UiLang _lang = UiLanguage.Default;

        internal static void Use(UiLang lang) { _lang = lang; }

        /// <summary>English is the base; a missing translation falls back to it.</summary>
        private static string T(string en, string ru)
        {
            if (_lang == UiLang.Ru && !string.IsNullOrEmpty(ru)) return ru;
            return en;
        }

        // ---- Tab context menu -------------------------------------------------

        internal static string MenuPin { get { return T("Pin tab", "Закрепить вкладку"); } }
        internal static string MenuUnpin { get { return T("Unpin tab", "Открепить вкладку"); } }
        internal static string MenuClose { get { return T("Close", "Закрыть"); } }
        internal static string MenuCloseOthers { get { return T("Close others", "Закрыть остальные"); } }
        internal static string MenuOpenFolder { get { return T("Open file location", "Открыть папку файла"); } }
        internal static string MenuTabColor { get { return T("Tab color", "Цвет вкладки"); } }
        internal static string MenuColorCustom { get { return T("More colors…", "Другой цвет…"); } }
        internal static string MenuColorReset { get { return T("Reset color", "Сбросить цвет"); } }
        internal static string MenuTabSize { get { return T("Tab size", "Размер вкладок"); } }

        // ---- Tooltips ---------------------------------------------------------

        internal static string TipProtectedView { get { return T("(Protected View)", "(Защищённый просмотр)"); } }
        internal static string TipUnsaved { get { return T("(unsaved changes)", "(есть несохранённые изменения)"); } }
        internal static string TipCloseDocument { get { return T("Close document", "Закрыть документ"); } }
        internal static string TipNewDocument { get { return T("New document", "Новый документ"); } }
        internal static string TipAllTabs { get { return T("All tabs", "Все вкладки"); } }
        internal static string TipExpandBar { get { return T("Expand the tab bar", "Развернуть панель вкладок"); } }
        internal static string TipCollapseBar { get { return T("Collapse the tab bar", "Свернуть панель вкладок"); } }

        // ---- All-tabs popup ---------------------------------------------------

        internal static string PopupNoMatches { get { return T("No matches", "Нет совпадений"); } }
        internal static string PopupSearchHint { get { return T("Search documents…", "Поиск документов…"); } }
        internal static string PopupSettings { get { return T("Settings…", "Настройки…"); } }

        // ---- Settings window --------------------------------------------------

        internal static string SettingsTitle { get { return T("Settings", "Настройки"); } }
        internal static string SettingsSectionSize { get { return T("Tab size", "Размер вкладок"); } }
        internal static string SettingsSectionKeyboard { get { return T("Keyboard", "Клавиатура"); } }
        internal static string SettingsCtrlTab { get { return T("Switch tabs with Ctrl+Tab", "Переключение вкладок сочетанием Ctrl+Tab"); } }
        internal static string SettingsSectionLanguage { get { return T("Language", "Язык"); } }
        internal static string SettingsLangAuto { get { return T("Same as Word", "Как в Word"); } }
        internal static string SettingsLangRestartHint { get { return T("applies to new windows", "применяется в новых окнах"); } }

        // Language names stay in their own language on purpose: a person who switched
        // Word to a language they cannot read must still find their way back.
        internal static string LangNameEn { get { return "English"; } }
        internal static string LangNameRu { get { return "Русский"; } }

        // ---- Tab size names ---------------------------------------------------

        internal static string[] SizeNames
        {
            get
            {
                return new[]
                {
                    T("Compact", "Компактные"),
                    T("Normal", "Обычные"),
                    T("Large", "Крупные"),
                    T("Extra large", "Очень крупные")
                };
            }
        }

        // ---- Tab color names (order matches TabTheme presets) -----------------

        internal static string[] ColorNames
        {
            get
            {
                return new[]
                {
                    T("Red", "Красный"),
                    T("Orange", "Оранжевый"),
                    T("Yellow", "Жёлтый"),
                    T("Green", "Зелёный"),
                    T("Teal", "Бирюзовый"),
                    T("Blue", "Синий"),
                    T("Purple", "Фиолетовый"),
                    T("Pink", "Розовый")
                };
            }
        }
    }
}
