using System;
using System.Collections.Generic;
using TabsForWord;

namespace TabsForWord.Tests
{
    /// <summary>
    /// Unit tests for the logic that does not need a running Word.
    /// A minimal runner with no external framework: PASS/FAIL plus an exit code.
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;

        private static int Main()
        {
            Run("DocumentTabModel.SameAs: equal objects", TestSameAsEqual);
            Run("DocumentTabModel.SameAs: Saved differs", TestSameAsSavedDiffers);
            Run("DocumentTabModel.ListsEqual: empty and equal lists", TestListsEqual);
            Run("DocumentTabModel.ListsEqual: different order => not equal", TestListsOrderMatters);
            Run("DocumentTabModel.ListsEqual: null tolerance", TestListsNull);
            Run("PvKey: negative and stable", TestPvKeyNegativeStable);
            Run("PvKey: case-insensitive", TestPvKeyCaseInsensitive);
            Run("PvKey: null does not throw", TestPvKeyNull);
            Run("TabStripControl: construction and UpdateTabs do not throw", TestTabStripControl);
            Run("TabStripControl 2b: strip height and collapsing", TestTabStrip2bCollapse);
            Run("TabStripControl 2b: overflow in a narrow strip", TestTabStrip2bOverflow);
            Run("TabListPopup: construction and the search filter", TestTabListPopup);
            Run("LoggingService: writing does not throw", TestLogging);
            Run("NativeHost: ComputePanelRect - the ordinary case", TestPanelRectNormal);
            Run("NativeHost: ComputePanelRect - clamps and degenerate cases", TestPanelRectClamps);
            Run("NativeHost: ComputePanelRect - Word side panes", TestPanelRectSidePane);
            Run("NativeHost: ScaleHeight/Scale96 - DPI conversion", TestDpiScaling);
            Run("NativeHost: Parse - Native + reserve by default", TestSettingsDefault);
            Run("NativeHost: Parse - mode=native, flags and legacy configs", TestSettingsNative);
            Run("NativeHost: Parse - the environment variable wins over the file", TestSettingsEnvOverride);
            Run("NativeHost: locator - _WwG canvas and the container under the ribbon", TestLocatorCanvasAnchor);
            Run("NativeHost: locator - full-height containers are skipped", TestLocatorCanvasOnly);
            Run("NativeHost: locator - fallback without a canvas", TestLocatorFallback);
            Run("NativeHost: locator - hidden canvas and an empty tree", TestLocatorNegative);
            Run("NativeHost: locator - a half-built Word window is rejected", TestLocatorWindowNotLaidOut);
            Run("NativeHost: locator - the File menu (full-page UI) is not a failure", TestLocatorFullPageUi);
            Run("NativeHost: locator - the Degraded flag on fallback paths", TestLocatorDegradedFlag);
            Run("NativeHost: WindowTreeDiagnostics.Format - dump structure", TestTreeFormat);
            Run("NativeHost: live SetParent hosting in a WinForms window", TestNativeHostLive);
            Run("NativeHost: locator - the ribbon shadow (DropShadow)", TestLocatorShadow);
            Run("NativeHost: Parse - the bg setting and colour parsing", TestSettingsBackground);
            Run("TabColorStore: parsing and the file line format", TestColorStoreParseFormat);
            Run("TabColorStore: saved by path, unsaved by hwnd", TestColorStoreResolve);
            Run("TabColorStore: file write and read (roundtrip)", TestColorStoreFileRoundtrip);
            Run("TabColorStore: colour migration on the first save", TestColorStoreMigrateFirstSave);
            Run("TabColorStore: colour migration on Save As", TestColorStoreMigrateSaveAs);
            Run("DocumentTabModel.SameAs: TabColor differs", TestSameAsColorDiffers);
            Run("TabTheme: the pastel tints are lighter than the source colour", TestThemeTints);
            Run("TabStripControl: coloured tabs draw without throwing", TestTabStripColored);
            Run("TabStripControl: the DPI override scales the strip", TestDpiOverride);
            Run("TabSizeSettings: parsing and clamps", TestSizeSettingsParse);
            Run("TabStripControl: the user size on top of DPI", TestUserScale);
            Run("TabStripControl: overflow - tabs do not escape the scroll zone", TestOverflowClipping);
            Run("TabStripControl: overflow - captions land on their own tabs", TestOverflowTextPlacement);

            Run("DocumentTabModel.SameAs: IsPinned differs", TestSameAsPinDiffers);
            Run("TabOrderStore: parsing and the file line format (PIN|path)", TestOrderStoreParseFormat);
            Run("TabOrderStore: saved by path, unsaved by hwnd", TestOrderStoreResolvePin);
            Run("TabOrderStore: file write and read (roundtrip)", TestOrderStoreFileRoundtrip);
            Run("TabOrderStore: RankOf and SaveOrder - restoring the order", TestOrderStoreRankAndSaveOrder);
            Run("TabOrderStore: closing every tab does not wipe pinning or order", TestOrderStoreSurvivesAllTabsClosed);
            Run("TabOrderStore: pinning migration on the first save", TestOrderStoreMigratePin);
            Run("HotkeySettings: parsing and the default value", TestHotkeySettingsParse);
            Run("TabStripControl: pinned tabs draw without throwing", TestTabStripPinnedRender);
            Run("TabListPopup: pinned tabs and the settings button do not throw", TestTabListPopupPinnedAndSettings);
            Run("Language: parsing language.cfg", TestLanguageParse);
            Run("Language: LCID -> interface language", TestLanguageFromLcid);
            Run("Language: captions switch and are never empty", TestStringsSwitch);
            Run("Language: the strip and the menu build in English", TestTabStripEnglish);

            Console.WriteLine();
            Console.WriteLine("Total: PASS=" + _passed + ", FAIL=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("[PASS] " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name + " -- " + ex.Message);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        private static DocumentTabModel Tab(int hwnd, string caption, bool saved = true,
            bool active = false, bool pv = false, string path = null, bool pinned = false)
        {
            return new DocumentTabModel
            {
                Hwnd = hwnd,
                Caption = caption,
                FullPath = path ?? ("C:\\Docs\\" + caption),
                Saved = saved,
                IsActive = active,
                IsProtectedView = pv,
                IsPinned = pinned
            };
        }

        // ------------------------------------------------------------------

        private static void TestSameAsEqual()
        {
            Assert(Tab(1, "a.docx").SameAs(Tab(1, "a.docx")), "identical tabs must be equal");
        }

        private static void TestSameAsSavedDiffers()
        {
            Assert(!Tab(1, "a.docx", saved: true).SameAs(Tab(1, "a.docx", saved: false)),
                "a change of Saved must count as a difference (the unsaved indicator)");
        }

        private static void TestListsEqual()
        {
            Assert(DocumentTabModel.ListsEqual(new List<DocumentTabModel>(), new List<DocumentTabModel>()),
                "empty lists are equal");
            var a = new List<DocumentTabModel> { Tab(1, "a"), Tab(2, "b", active: true) };
            var b = new List<DocumentTabModel> { Tab(1, "a"), Tab(2, "b", active: true) };
            Assert(DocumentTabModel.ListsEqual(a, b), "identical lists are equal");
        }

        private static void TestListsOrderMatters()
        {
            var a = new List<DocumentTabModel> { Tab(1, "a"), Tab(2, "b") };
            var b = new List<DocumentTabModel> { Tab(2, "b"), Tab(1, "a") };
            Assert(!DocumentTabModel.ListsEqual(a, b), "tab order matters");
        }

        private static void TestListsNull()
        {
            var a = new List<DocumentTabModel> { Tab(1, "a") };
            Assert(!DocumentTabModel.ListsEqual(a, null), "a list is not equal to null");
            Assert(!DocumentTabModel.ListsEqual(null, a), "null is not equal to a list");
        }

        private static void TestPvKeyNegativeStable()
        {
            int k1 = DocumentWindowManager.PvKey("C:\\Docs\\file.docx");
            int k2 = DocumentWindowManager.PvKey("C:\\Docs\\file.docx");
            Assert(k1 < 0, "a PV key must be negative (so it never collides with an HWND)");
            Assert(k1 == k2, "a PV key must be stable for the same path");
            int other = DocumentWindowManager.PvKey("C:\\Docs\\other.docx");
            Assert(k1 != other, "different paths must give different keys");
        }

        private static void TestPvKeyCaseInsensitive()
        {
            Assert(DocumentWindowManager.PvKey("C:\\A\\B.DOCX") == DocumentWindowManager.PvKey("c:\\a\\b.docx"),
                "Windows paths are case-insensitive");
        }

        private static void TestPvKeyNull()
        {
            int k = DocumentWindowManager.PvKey(null);
            Assert(k < 0, "a null path must not throw and still gives a negative key");
        }

        // ------------------------------------------------------------------
        // Interface language (stage 29)
        // ------------------------------------------------------------------

        private static void TestLanguageParse()
        {
            Assert(TabsForWord.UiLanguage.Parse(null) == null, "no file => auto-detect");
            Assert(TabsForWord.UiLanguage.Parse(new[] { "lang=auto" }) == null, "lang=auto => auto-detect");
            Assert(TabsForWord.UiLanguage.Parse(new[] { "# comment", "", "lang = RU" }) == TabsForWord.UiLang.Ru,
                "lang=RU (case and spaces) => Russian");
            Assert(TabsForWord.UiLanguage.Parse(new[] { "lang=english" }) == TabsForWord.UiLang.En,
                "lang=english => English");
            Assert(TabsForWord.UiLanguage.Parse(new[] { "lang=klingon" }) == null,
                "an unknown language => auto-detect rather than a crash");
        }

        private static void TestLanguageFromLcid()
        {
            Assert(TabsForWord.UiLanguage.FromLcid(1049) == TabsForWord.UiLang.Ru, "1049 (ru-RU) => Russian");
            Assert(TabsForWord.UiLanguage.FromLcid(2073) == TabsForWord.UiLang.Ru, "2073 (ru-MD) => Russian");
            Assert(TabsForWord.UiLanguage.FromLcid(1033) == TabsForWord.UiLang.En, "1033 (en-US) => English");
            Assert(TabsForWord.UiLanguage.FromLcid(1031) == TabsForWord.UiLang.En,
                "an untranslated language (de-DE) => English as the neutral one");
            Assert(TabsForWord.UiLanguage.FromLcid(0) == TabsForWord.UiLang.En, "a zero LCID => English");
        }

        private static void TestStringsSwitch()
        {
            // Every caption is read by reflection: a new menu item cannot be left
            // untranslated - the test sees both an empty string and one equal to the Russian.
            var type = typeof(TabsForWord.TabStripControl).Assembly.GetType("TabsForWord.Strings");
            Assert(type != null, "the Strings class was found");
            var props = type.GetProperties(System.Reflection.BindingFlags.Static |
                                           System.Reflection.BindingFlags.NonPublic |
                                           System.Reflection.BindingFlags.Public);
            Assert(props.Length >= 20, "at least 20 captions in the table (currently " + props.Length + ")");

            var ru = new Dictionary<string, string>();
            var en = new Dictionary<string, string>();
            foreach (var lang in new[] { TabsForWord.UiLang.Ru, TabsForWord.UiLang.En })
            {
                TabsForWord.UiLanguage.OverrideForTests(lang);
                var into = lang == TabsForWord.UiLang.Ru ? ru : en;
                foreach (var p in props)
                {
                    var value = p.GetValue(null, null);
                    var arr = value as string[];
                    var text = arr != null ? string.Join("|", arr) : (string)value;
                    Assert(!string.IsNullOrEmpty(text), "caption " + p.Name + " is not empty (" + lang + ")");
                    into[p.Name] = text;
                }
            }

            // The language names are identical on purpose (each is written in its own language).
            int translated = 0;
            foreach (var kv in ru)
            {
                if (kv.Key == "LangNameEn" || kv.Key == "LangNameRu") continue;
                if (kv.Value != en[kv.Key]) translated++;
            }
            Assert(translated >= 20, "at least 20 captions are translated (currently " + translated + ")");

            TabsForWord.UiLanguage.OverrideForTests(TabsForWord.UiLang.Ru);
        }

        private static void TestTabStripEnglish()
        {
            TabsForWord.UiLanguage.OverrideForTests(TabsForWord.UiLang.En);
            try
            {
                using (var ctrl = new TabStripControl())
                {
                    ctrl.Size = new System.Drawing.Size(600, 34);
                    ctrl.UpdateTabs(new List<DocumentTabModel>
                    {
                        Tab(1, "Contract.docx", active: true),
                        Tab(2, "Report.docx", saved: false)
                    });
                    using (var bmp = new System.Drawing.Bitmap(ctrl.Width, ctrl.Height))
                        ctrl.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, ctrl.Width, ctrl.Height));

                    // Switching the language on the fly drops the context-menu cache and does not break drawing.
                    ctrl.RefreshLocalizedUi();
                    using (var bmp = new System.Drawing.Bitmap(ctrl.Width, ctrl.Height))
                        ctrl.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, ctrl.Width, ctrl.Height));
                }
            }
            finally
            {
                TabsForWord.UiLanguage.OverrideForTests(TabsForWord.UiLang.Ru);
            }
        }

        private static void TestTabStripControl()
        {
            using (var ctrl = new TabStripControl())
            {
                ctrl.Size = new System.Drawing.Size(600, 34);
                // in-window mode: the flush fill and the background override do not break drawing
                ctrl.SetActiveTabFlushBottom(true);
                ctrl.SetPanelBackgroundOverride(System.Drawing.Color.FromArgb(0xE8, 0xE9, 0xEF));
                ctrl.SetPanelBackgroundOverride(null);
                ctrl.UpdateTabs(new List<DocumentTabModel>());
                ctrl.UpdateTabs(new List<DocumentTabModel> { Tab(1, "Contract.docx", active: true) });
                ctrl.UpdateTabs(new List<DocumentTabModel>
                {
                    Tab(1, "Contract.docx"),
                    Tab(2, "Report.docx", saved: false, active: true),
                    Tab(-5, "From the internet.docx", pv: true)
                });
                // 15 tabs => overflow and scrolling must not throw
                var many = new List<DocumentTabModel>();
                for (int i = 1; i <= 15; i++) many.Add(Tab(i, "Document-" + i + ".docx", active: i == 15));
                ctrl.UpdateTabs(many);
            }
        }

        private static void TestTabStrip2bCollapse()
        {
            using (var ctrl = new TabStripControl())
            {
                int expanded = ctrl.DesiredContentHeight;
                Assert(expanded >= 37, "an expanded strip is at least 37 px (specification 2b, section 1)");

                ctrl.Size = new System.Drawing.Size(900, expanded);
                ctrl.UpdateTabs(new List<DocumentTabModel>
                {
                    Tab(1, "Supply contract_v3.docx", active: true, saved: false),
                    Tab(2, "Report for July", saved: false),
                    Tab(3, "Document1")
                });

                ctrl.SetCollapsed(true);
                Assert(ctrl.Collapsed, "after SetCollapsed(true) the strip is collapsed");
                Assert(ctrl.DesiredContentHeight < expanded,
                    "a collapsed strip is shorter than an expanded one (section 8)");

                ctrl.SetCollapsed(false);
                Assert(!ctrl.Collapsed, "after SetCollapsed(false) the strip is expanded");
                Assert(ctrl.DesiredContentHeight == expanded, "the height is restored");
            }
        }

        private static void TestTabStrip2bOverflow()
        {
            using (var ctrl = new TabStripControl())
            {
                // A narrow strip plus 25 tabs with long names: even shrinking, the compact
                // mode and scrolling must not throw.
                ctrl.Size = new System.Drawing.Size(480, ctrl.DesiredContentHeight);
                var many = new List<DocumentTabModel>();
                for (int i = 1; i <= 25; i++)
                    many.Add(Tab(i, "A very long document name number " + i + ".docx",
                        active: i == 13, saved: i % 3 != 0));
                ctrl.UpdateTabs(many);

                // Changing the active tab while overflowing (EnsureActiveVisible)
                for (int i = 0; i < many.Count; i++) many[i].IsActive = i == 24;
                ctrl.UpdateTabs(new List<DocumentTabModel>(many));
            }
        }

        private static void TestTabListPopup()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    var tabs = new List<DocumentTabModel>
                    {
                        Tab(1, "Contract.docx", active: true),
                        Tab(2, "Report.docx", saved: false),
                        Tab(3, "Letter.docx")
                    };
                    using (var popup = new TabListPopup(tabs, 1f, delegate { }, delegate { }))
                    {
                        // the constructor applies the filter and the layout - nothing may throw
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("TabListPopup: " + failure.Message, failure);
        }

        private static void TestLogging()
        {
            LoggingService.Info("Test run marker");
            LoggingService.Error("Test error marker", new InvalidOperationException("test"));
        }

        // ------------------------------------------------------------------
        // NativeHost: pure geometry
        // ------------------------------------------------------------------

        private static void TestPanelRectNormal()
        {
            var r = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1200, 800), 150, 37);
            Assert(r.X == 0 && r.Y == 150 && r.Width == 1200 && r.Height == 37,
                "the strip spans the full client width, pinned to the top of the document area");
        }

        private static void TestPanelRectSidePane()
        {
            // A Word side pane is open (Navigation, 350 px on the left): the document area
            // is 350..1920, so the tab strip only occupies that range.
            var r = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1920, 1000), 178, 37, 350, 1570);
            Assert(r.X == 350 && r.Width == 1570 && r.Y == 178,
                "the strip does not cover the left Word side pane");

            // Right side pane: the document area is 0..1500 out of 1920
            var r2 = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1920, 1000), 178, 37, 0, 1500);
            Assert(r2.X == 0 && r2.Width == 1500, "the strip does not cover the right side pane");

            // The area is wider than the client - it gets trimmed
            var r3 = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 600), 100, 37, 100, 5000);
            Assert(r3.X == 100 && r3.Width == 900, "the width is clamped to the client area");

            // A degenerate document area => Empty
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 600), 100, 37, 1200, 300).IsEmpty,
                "a document area outside the client => Empty");
        }

        private static void TestPanelRectClamps()
        {
            // a negative top is clamped to 0
            var r1 = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 600), -20, 37);
            Assert(r1.Y == 0, "a negative top is clamped to 0");
            // a top below the bottom is clamped so that the strip still fits
            var r2 = TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 600), 900, 37);
            Assert(r2.Y == 600 - 37, "the strip does not spill past the bottom of the client area");
            // degenerate cases
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(0, 600), 10, 37).IsEmpty, "zero client width => Empty");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 20), 0, 37).IsEmpty, "a window shorter than the strip => Empty");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ComputePanelRect(
                new System.Drawing.Size(1000, 600), 10, 0).IsEmpty, "zero strip height => Empty");
        }

        private static void TestDpiScaling()
        {
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ScaleHeight(37, 96, 96) == 37,
                "the same DPI - the height does not change");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ScaleHeight(37, 96, 144) == 56,
                "96->144: 37*1.5 = 55.5 -> 56 (rounding)");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ScaleHeight(56, 144, 96) == 37,
                "the reverse conversion 144->96");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.ScaleHeight(37, 0, 144) == 37,
                "an invalid source DPI - no change");
            Assert(TabsForWord.NativeHost.NativeHostLayoutMath.Scale96(8, 192) == 16,
                "the 8 px threshold at 200% => 16 px");
        }

        // ------------------------------------------------------------------
        // NativeHost: settings
        // ------------------------------------------------------------------

        private static void TestSettingsDefault()
        {
            bool dump, reserve;
            // ADR-015 (v1.6.0): a clean installation = native + reserved space.
            var mode = TabsForWord.NativeHost.TabHostSettings.Parse(null, null, out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.Native,
                "no file and no variable => the main Native mode");
            Assert(reserve, "reserve is on by default (the strip pushes the document down rather than covering it)");
            Assert(!dump, "dump is off by default");

            mode = TabsForWord.NativeHost.TabHostSettings.Parse(
                new[] { "# comment", "", "junk", "mode=ctp" }, null, out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.CustomTaskPane,
                "mode=ctp => the classic CustomTaskPane");
        }

        private static void TestSettingsNative()
        {
            bool dump, reserve;
            var mode = TabsForWord.NativeHost.TabHostSettings.Parse(
                new[] { "mode = Native", "dump=1", "reserve=true" }, null, out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.Native,
                "mode=native (case- and space-insensitive) => Native");
            Assert(dump, "dump=1 is recognised");
            Assert(reserve, "reserve=true is recognised");

            // Configs left over from the experimental builds (the mode name before v1.6.0).
            mode = TabsForWord.NativeHost.TabHostSettings.Parse(
                new[] { "mode=nativeexperimental" }, null, out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.Native,
                "the legacy name mode=nativeexperimental => Native");

            // The overlay variant: reserve is switched off explicitly, the default does not override it.
            mode = TabsForWord.NativeHost.TabHostSettings.Parse(
                new[] { "mode=native", "reserve=0" }, null, out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.Native && !reserve,
                "reserve=0 turns the reservation off (the strip lies over the document)");
        }

        private static void TestSettingsEnvOverride()
        {
            bool dump, reserve;
            var mode = TabsForWord.NativeHost.TabHostSettings.Parse(
                new[] { "mode=native" }, "0", out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.CustomTaskPane,
                "WORDTABS_NATIVE_HOST=0 disables native from the file");
            mode = TabsForWord.NativeHost.TabHostSettings.Parse(null, "1", out dump, out reserve);
            Assert(mode == TabsForWord.NativeHost.TabHostMode.Native,
                "WORDTABS_NATIVE_HOST=1 enables native without a file");
        }

        // ------------------------------------------------------------------
        // NativeHost: the document-area locator
        // ------------------------------------------------------------------

        private static TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot Win(
            long hwnd, long parent, string cls, int x, int y, int w, int h, bool visible = true)
        {
            return new TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot
            {
                Hwnd = new IntPtr(hwnd),
                Parent = new IntPtr(parent),
                ClassName = cls,
                ScreenRect = new System.Drawing.Rectangle(x, y, w, h),
                Visible = visible
            };
        }

        private static void TestLocatorCanvasAnchor()
        {
            // The typical Word hierarchy: OpusApp(0x1000) -> _WwF (the whole client
            // area) -> _WwB (under the ribbon) -> _WwG (the canvas).
            var root = new IntPtr(0x1000);
            var client = new System.Drawing.Rectangle(100, 100, 1200, 800);
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x1100, 0x1000, "_WwF", 100, 100, 1200, 800),
                Win(0x1200, 0x1100, "_WwB", 100, 250, 1100, 650),
                Win(0x1300, 0x1200, "_WwG", 130, 280, 1000, 600),
                Win(0x1400, 0x1000, "NetUIHWND", 100, 100, 1200, 150) // the ribbon
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 96);
            Assert(zone != null, "a zone must be found");
            Assert(zone.AnchorHwnd == new IntPtr(0x1200),
                "the anchor is _WwB (the topmost container starting under the ribbon), not the full-height _WwF");
            Assert(zone.CanvasHwnd == new IntPtr(0x1300), "the canvas is _WwG");
            Assert(zone.AnchorRect.Top == 250, "the zone top matches the top of _WwB");
        }

        private static void TestLocatorCanvasOnly()
        {
            // Every container reaches the top of the client area, so the canvas itself
            // becomes the anchor.
            var root = new IntPtr(0x2000);
            var client = new System.Drawing.Rectangle(0, 0, 1000, 700);
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x2100, 0x2000, "_WwF", 0, 0, 1000, 700),
                Win(0x2200, 0x2100, "_WwB", 0, 0, 1000, 700),
                Win(0x2300, 0x2200, "_WwG", 20, 180, 960, 500)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 96);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x2300),
                "full-height containers are skipped, the anchor is the _WwG canvas");
        }

        private static void TestLocatorFallback()
        {
            // There is no canvas (its class was renamed in some future Word, say) - the
            // topmost suitable window below the client top is taken.
            var root = new IntPtr(0x3000);
            var client = new System.Drawing.Rectangle(0, 0, 1000, 700);
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x3100, 0x3000, "RibbonThing", 0, 0, 1000, 140),      // starts at the top - skipped
                Win(0x3200, 0x3000, "SmallThing", 0, 200, 300, 400),      // too narrow (<40%)
                Win(0x3300, 0x3000, "ContentThing", 0, 140, 1000, 560),   // suitable, and topmost
                Win(0x3400, 0x3000, "LowerThing", 0, 300, 1000, 300)      // lower - loses
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 96);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x3300),
                "the fallback picks the topmost sufficiently large window under the ribbon");
            Assert(zone.Reason.Contains("fallback"), "the reason is marked as a fallback");
        }

        private static void TestLocatorNegative()
        {
            var root = new IntPtr(0x4000);
            var client = new System.Drawing.Rectangle(0, 0, 1000, 700);
            // a hidden canvas is ignored and there are no other candidates
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x4100, 0x4000, "_WwG", 0, 180, 960, 400, visible: false)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 96);
            Assert(zone == null, "a hidden canvas must not be chosen");

            zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client,
                new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>(), 96);
            Assert(zone == null, "an empty tree => null (the strip hides, Word keeps working)");
        }

        private static void TestLocatorWindowNotLaidOut()
        {
            // A field case (machine A, 2026-07-26): Word had only just created the window,
            // the layout was not built yet and the canvas covered the WHOLE client area.
            // The locator used to take it as the anchor, the strip landed at 0,0 over the
            // title bar and the ribbon, and the zone was cached for the whole session.
            // Correct behaviour: no zone is found and the search repeats later.
            var root = new IntPtr(0x7000);
            var client = new System.Drawing.Rectangle(0, 0, 2534, 1528);
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x7100, 0x7000, "_WwF", 0, 0, 2534, 1528),
                Win(0x7200, 0x7100, "_WwB", 0, 0, 2534, 1528),
                Win(0x7300, 0x7200, "_WwG", 0, 0, 2534, 1528)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 144);
            Assert(zone == null, "a half-built window => no zone (the strip does not land over the ribbon)");

            // The same window a second later, with the layout built - the zone is found.
            var laidOut = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x7100, 0x7000, "_WwF", 0, 267, 2534, 1228),
                Win(0x7200, 0x7100, "_WwB", 0, 267, 2534, 1228),
                Win(0x7300, 0x7200, "_WwG", 0, 300, 2534, 1195)
            };
            zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, laidOut, 144);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x7100),
                "once laid out, the anchor is _WwF (the topmost container under the ribbon)");
            Assert(!zone.Degraded, "a proper choice is not marked Degraded");
        }

        private static void TestLocatorFullPageUi()
        {
            // A field case (machine C, 2026-07-28): the user opened the File menu
            // (Backstage). Word puts up a full-page UI window and HIDES the whole
            // document subtree. There is no document area, so "no zone" is correct -
            // but it is a normal state, not a failure: no failure counted, no
            // window-tree dump, and the strip must not fall back to a task pane.
            var root = new IntPtr(0x9000);
            var client = new System.Drawing.Rectangle(0, 0, 1920, 1040);
            var backstage = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x9100, 0x9000, "FullpageUIHost", 0, 0, 1920, 1040),
                Win(0x9110, 0x9100, "NetUIHWND", 0, 0, 1920, 1040),
                Win(0x9200, 0x9000, "_WwF", 0, 215, 1920, 803, visible: false),
                Win(0x9300, 0x9200, "_WwB", 0, 215, 1920, 803, visible: false),
                Win(0x9400, 0x9300, "_WwG", 24, 239, 1879, 779, visible: false)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, backstage, 96);
            Assert(zone == null, "the File menu is open => no zone (the strip must not float over it)");
            Assert(TabsForWord.NativeHost.WordNativeWindowLocator.IsFullPageUiActive(client, backstage),
                "a visible FullpageUIHost across the client area is recognised");

            // The File menu is closed: the same host window is still there, but hidden.
            var closed = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x9100, 0x9000, "FullpageUIHost", 0, 0, 1920, 1040, visible: false),
                Win(0x9200, 0x9000, "_WwF", 0, 215, 1920, 803),
                Win(0x9300, 0x9200, "_WwB", 0, 215, 1920, 803),
                Win(0x9400, 0x9300, "_WwG", 24, 239, 1879, 779)
            };
            Assert(!TabsForWord.NativeHost.WordNativeWindowLocator.IsFullPageUiActive(client, closed),
                "a hidden FullpageUIHost is not the File menu");
            var back = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, closed, 96);
            Assert(back != null && back.AnchorHwnd == new IntPtr(0x9200),
                "once the File menu is closed the zone is found again");

            // A small full-page host (Word keeps such leftovers around) is not the File menu.
            var leftover = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x9100, 0x9000, "FullpageUIHost", 0, 0, 100, 100)
            };
            Assert(!TabsForWord.NativeHost.WordNativeWindowLocator.IsFullPageUiActive(client, leftover),
                "a 100x100 FullpageUIHost does not cover the document");

            // The important negative: a half-built window (machine A) has no full-page UI,
            // so it stays a real "not found" - counters, retries and the dump keep working.
            var halfBuilt = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x9200, 0x9000, "_WwF", 0, 0, 1920, 1040),
                Win(0x9300, 0x9200, "_WwB", 0, 0, 1920, 1040),
                Win(0x9400, 0x9300, "_WwG", 0, 0, 1920, 1040)
            };
            Assert(TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, halfBuilt, 96) == null,
                "a half-built window still yields no zone");
            Assert(!TabsForWord.NativeHost.WordNativeWindowLocator.IsFullPageUiActive(client, halfBuilt),
                "a half-built window must NOT be excused as the File menu");
            Assert(!TabsForWord.NativeHost.WordNativeWindowLocator.IsFullPageUiActive(client, null),
                "a null tree does not throw");
        }

        private static void TestLocatorDegradedFlag()
        {
            // Anchor = canvas: a workable choice, but marked Degraded - the host will
            // periodically try to upgrade to the container under the ribbon.
            var root = new IntPtr(0x8000);
            var client = new System.Drawing.Rectangle(0, 0, 1000, 700);
            var canvasOnly = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x8100, 0x8000, "_WwF", 0, 0, 1000, 700),
                Win(0x8200, 0x8100, "_WwB", 0, 0, 1000, 700),
                Win(0x8300, 0x8200, "_WwG", 20, 180, 960, 500)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, canvasOnly, 96);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x8300) && zone.Degraded,
                "anchor = canvas is marked Degraded");

            // The generic fallback without a canvas is Degraded as well.
            var noCanvas = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x8400, 0x8000, "ContentThing", 0, 140, 1000, 560)
            };
            zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, noCanvas, 96);
            Assert(zone != null && zone.Degraded, "a fallback without a canvas is marked Degraded");

            // A proper choice (canvas + container under the ribbon) is not Degraded.
            var full = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x8500, 0x8000, "_WwF", 0, 178, 1000, 522),
                Win(0x8600, 0x8500, "_WwG", 20, 200, 960, 500)
            };
            zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, full, 96);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x8500) && !zone.Degraded,
                "the _WwF anchor under the ribbon - a proper choice");
        }

        private static void TestTreeFormat()
        {
            var root = new IntPtr(0x5000);
            var nodes = new List<TabsForWord.NativeHost.WindowTreeDiagnostics.Node>
            {
                new TabsForWord.NativeHost.WindowTreeDiagnostics.Node
                {
                    Hwnd = root, ClassName = "OpusApp", Text = "Document1 - Word", Visible = true
                },
                new TabsForWord.NativeHost.WindowTreeDiagnostics.Node
                {
                    Hwnd = new IntPtr(0x5100), Parent = root, ClassName = "_WwF", Visible = true
                },
                new TabsForWord.NativeHost.WindowTreeDiagnostics.Node
                {
                    Hwnd = new IntPtr(0x5200), Parent = new IntPtr(0x5100), ClassName = "_WwG", Visible = true
                }
            };
            var text = TabsForWord.NativeHost.WindowTreeDiagnostics.Format(root, nodes, "unit-test");
            Assert(text.Contains("OpusApp") && text.Contains("_WwF") && text.Contains("_WwG"),
                "the dump contains every window class");
            Assert(text.Contains("  0x5100"), "a child window is printed with an indent");
            Assert(text.Contains("    0x5200"), "a grandchild is printed with a double indent");
            Assert(text.Contains("unit-test"), "the dump context is recorded");
        }

        private static void TestLocatorShadow()
        {
            // The real Word structure: DropShadow (a thin shadow strip) sits right on top
            // of the document area. The strip has to know its HWND so it can place itself
            // below it in the z-order.
            var root = new IntPtr(0x6000);
            var client = new System.Drawing.Rectangle(0, 0, 1200, 900);
            var nodes = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x6100, 0x6000, "DropShadow", 0, 178, 1200, 6),
                Win(0x6200, 0x6000, "_WwF", 0, 178, 1200, 700),
                Win(0x6300, 0x6200, "_WwB", 0, 178, 1200, 700),
                Win(0x6400, 0x6300, "_WwG", 24, 202, 1150, 676)
            };
            var zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, nodes, 96);
            Assert(zone != null && zone.AnchorHwnd == new IntPtr(0x6200), "the anchor is _WwF");
            Assert(zone.ShadowHwnd == new IntPtr(0x6100), "the ribbon shadow was found");

            // A thick or distant DropShadow window does not qualify
            var far = new List<TabsForWord.NativeHost.WordNativeWindowLocator.WindowSnapshot>
            {
                Win(0x6100, 0x6000, "DropShadow", 0, 500, 1200, 6),   // not at the top of the zone
                Win(0x6200, 0x6000, "_WwF", 0, 178, 1200, 700),
                Win(0x6400, 0x6200, "_WwG", 24, 202, 1150, 676)
            };
            zone = TabsForWord.NativeHost.WordNativeWindowLocator.FindContentZone(root, client, far, 96);
            Assert(zone != null && zone.ShadowHwnd == IntPtr.Zero, "a foreign shadow is not taken");
        }

        private static void TestSettingsBackground()
        {
            // colour parsing
            var c = TabsForWord.NativeHost.TabHostSettings.ParseColor("#E8E9EF");
            Assert(c.HasValue && c.Value.R == 0xE8 && c.Value.G == 0xE9 && c.Value.B == 0xEF,
                "#RRGGBB parses");
            Assert(TabsForWord.NativeHost.TabHostSettings.ParseColor("e8e9ef").HasValue,
                "the hash may be omitted");
            Assert(!TabsForWord.NativeHost.TabHostSettings.ParseColor("#XYZ").HasValue &&
                   !TabsForWord.NativeHost.TabHostSettings.ParseColor("").HasValue &&
                   !TabsForWord.NativeHost.TabHostSettings.ParseColor("#12345").HasValue,
                "junk is rejected");

            // the bg setting
            bool dump, reserve;
            string bg;
            TabsForWord.NativeHost.TabHostSettings.Parse(null, null, out dump, out reserve, out bg);
            Assert(bg == "auto", "bg=auto by default");
            TabsForWord.NativeHost.TabHostSettings.Parse(new[] { "bg=spec" }, null, out dump, out reserve, out bg);
            Assert(bg == "spec", "bg=spec is accepted");
            TabsForWord.NativeHost.TabHostSettings.Parse(new[] { "bg=#A1B2C3" }, null, out dump, out reserve, out bg);
            Assert(bg == "#a1b2c3", "bg=#RRGGBB is accepted (lower-cased after trimming)");
            TabsForWord.NativeHost.TabHostSettings.Parse(new[] { "bg=junk" }, null, out dump, out reserve, out bg);
            Assert(bg == "auto", "an invalid bg is ignored (auto stays)");
        }

        // ------------------------------------------------------------------
        // NativeHost: live SetParent hosting (no Word: a form imitates the window)
        // ------------------------------------------------------------------

        private static void TestNativeHostLive()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    TabsForWord.NativeHost.TabHostSettings.OverrideForTests(
                        TabsForWord.NativeHost.TabHostMode.CustomTaskPane, false, false);

                    // The host now works in physical pixels (per-monitor-v2): the test form
                    // is created in the same context, otherwise on a monitor scaled other
                    // than 100% its coordinates would be virtualised and the comparisons
                    // with LastAppliedRect would differ by the DPI factor.
                    TabsForWord.NativeHost.NativeWin32.SetThreadDpiAwarenessContextSafe(
                        TabsForWord.NativeHost.NativeWin32.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

                    using (var form = new System.Windows.Forms.Form())
                    {
                        form.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        form.Location = new System.Drawing.Point(60, 60);
                        form.Size = new System.Drawing.Size(820, 560);
                        form.ShowInTaskbar = false;

                        // The "document area": a visible child control below the "ribbon"
                        var fakeCanvas = new System.Windows.Forms.Panel
                        {
                            Left = 10,
                            Top = 120,
                            Width = 760,
                            Height = 320,
                            BackColor = System.Drawing.Color.White
                        };
                        form.Controls.Add(fakeCanvas);
                        form.Show();
                        System.Windows.Forms.Application.DoEvents();

                        var mgr = new TabsForWord.NativeHost.NativeTabHostManager(new TabPaneManager());
                        try
                        {
                            int hwnd = form.Handle.ToInt32();

                            var host = mgr.EnsureHostForWindow(hwnd);
                            Assert(host != null, "a host must be created for a live window");
                            var host2 = mgr.EnsureHostForWindow(hwnd);
                            Assert(ReferenceEquals(host, host2),
                                "a repeated EnsureHostForWindow does not create a duplicate");

                            host.PushTabs(new List<DocumentTabModel>
                            {
                                Tab(1, "Contract.docx", active: true),
                                Tab(2, "Report.docx", saved: false)
                            });
                            host.UpdateLayout();
                            System.Windows.Forms.Application.DoEvents();

                            var rect = host.LastAppliedRect;
                            Assert(!rect.IsEmpty, "the layout was applied");
                            Assert(rect.X == fakeCanvas.Left && rect.Width == fakeCanvas.Width,
                                "the strip spans the horizontal range of the document area " +
                                "(not the full width: Word side panes are not covered)");
                            Assert(rect.Y == fakeCanvas.Top,
                                "the strip is pinned to the top of the document area (fallback anchor)");
                            Assert(rect.Height > 0 && rect.Height < 200, "the strip height is sensible");

                            Assert(host.HostHandle != IntPtr.Zero, "the host HWND exists");
                            // the host HWND really is a child of the parent window
                            Assert(TabsForWord.NativeHost.NativeWin32.IsChild(form.Handle, host.HostHandle),
                                "the host HWND is a child window of the form");

                            // A repeated UpdateLayout is idempotent
                            host.UpdateLayout();
                            Assert(host.LastAppliedRect == rect, "a repeated layout changes nothing");

                            mgr.RemoveHostForWindow(hwnd);
                            mgr.RemoveHostForWindow(hwnd); // removing twice does not throw
                            Assert(host.HostHandle == IntPtr.Zero, "after removal the host HWND is destroyed");
                        }
                        finally
                        {
                            mgr.Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("NativeHostLive: " + failure.Message, failure);
        }

        // ------------------------------------------------------------------
        // Tab colours
        // ------------------------------------------------------------------

        private static void TestColorStoreParseFormat()
        {
            var parsed = TabColorStore.ParseLines(new[]
            {
                "# comment",
                "",
                "D93025|C:\\Docs\\Contract.docx",
                "#1A73E8|C:\\Docs\\Report.docx",   // a leading hash is allowed
                "junk with no separator",
                "GGGGGG|C:\\Docs\\broken-colour.docx",
                "D93025|",                        // empty path - ignored
                "0F9DA8|C:\\Docs\\Contract.docx"   // repeated path - the last one wins
            });
            Assert(parsed.Count == 2, "exactly 2 valid entries were parsed, not " + parsed.Count);
            Assert(parsed["C:\\Docs\\contract.DOCX"] == System.Drawing.Color.FromArgb(0x0F, 0x9D, 0xA8),
                "the path is case-insensitive and a repeated path overwrites the colour");
            Assert(parsed["C:\\Docs\\Report.docx"] == System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8),
                "a colour with a leading hash was parsed");

            var lines = TabColorStore.FormatLines(parsed);
            var reparsed = TabColorStore.ParseLines(lines);
            Assert(reparsed.Count == parsed.Count, "format/parse roundtrip loses nothing");
            Assert(TabColorStore.FormatRgb(System.Drawing.Color.FromArgb(0x0F, 0x9D, 0xA8)) == "0F9DA8",
                "FormatRgb - six hexadecimal digits");
        }

        private static void TestColorStoreResolve()
        {
            var store = new TabColorStore(null); // no file
            var red = System.Drawing.Color.FromArgb(0xD9, 0x30, 0x25);
            var blue = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);

            // A saved document: the key is the path; the hwnd does not matter
            store.SetColor("C:\\Docs\\Contract.docx", 111, red);
            Assert(store.Resolve("C:\\Docs\\CONTRACT.docx", 999) == red,
                "the colour of a saved document is found by path (case-insensitively)");

            // Unsaved: the key is the hwnd (FullPath is a name without a path)
            store.SetColor("Document1", 222, blue);
            Assert(store.Resolve("Document1", 222) == blue, "the colour of an unsaved document is found by hwnd");
            Assert(store.Resolve("Document1", 333) == null, "a different hwnd has no colour");

            // Reset, and cleanup when the window closes
            store.SetColor("C:\\Docs\\Contract.docx", 111, null);
            Assert(store.Resolve("C:\\Docs\\Contract.docx", 111) == null, "the colour was reset");
            store.ForgetWindow(222);
            Assert(store.Resolve("Document1", 222) == null, "closing the window forgets the session colour");
        }

        private static void TestColorStoreFileRoundtrip()
        {
            var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "TabsForWord-test-colors-" + Guid.NewGuid().ToString("N") + ".cfg");
            try
            {
                var green = System.Drawing.Color.FromArgb(0x1E, 0x8E, 0x3E);
                var store1 = new TabColorStore(file);
                store1.SetColor("C:\\Docs\\Report for July.docx", 1, green);
                Assert(System.IO.File.Exists(file), "the colour file is created on the first save");

                var store2 = new TabColorStore(file); // "Word restart"
                Assert(store2.Resolve("C:\\Docs\\Report for July.docx", 42) == green,
                    "the colour survived reloading the store (spaces in the path)");

                store2.SetColor("C:\\Docs\\Report for July.docx", 42, null);
                var store3 = new TabColorStore(file);
                Assert(store3.Resolve("C:\\Docs\\Report for July.docx", 7) == null,
                    "resetting the colour is saved to the file as well");
            }
            finally
            {
                try { System.IO.File.Delete(file); } catch { }
            }
        }

        private static void TestColorStoreMigrateFirstSave()
        {
            var store = new TabColorStore(null);
            var red = System.Drawing.Color.FromArgb(0xD9, 0x30, 0x25);

            // An unsaved document got a colour and was then saved (Ctrl+S)
            store.SetColor("Document1", 111, red);
            store.HandlePathChanged(111, "Document1", "C:\\Docs\\New.docx");
            Assert(store.Resolve("C:\\Docs\\New.docx", 111) == red,
                "after the first save the colour is found by path");
            Assert(store.Resolve("Document2", 111) == null,
                "the hwnd entry was cleared after the migration");

            // A document without a colour - the migration creates nothing
            store.HandlePathChanged(222, "Document2", "C:\\Docs\\Other.docx");
            Assert(store.Resolve("C:\\Docs\\Other.docx", 222) == null,
                "with no colour the migration creates nothing");

            // Saving without changing the path (the same path) is a no-op
            store.HandlePathChanged(111, "C:\\Docs\\New.docx", "C:\\Docs\\NEW.docx");
            Assert(store.Resolve("C:\\Docs\\New.docx", 111) == red,
                "a change of case in the path does not count as a change of path");
        }

        private static void TestColorStoreMigrateSaveAs()
        {
            var store = new TabColorStore(null);
            var blue = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);

            // Save As: the colour follows the tab, the previous file keeps its own
            store.SetColor("C:\\Docs\\Old.docx", 111, blue);
            store.HandlePathChanged(111, "C:\\Docs\\Old.docx", "C:\\Docs\\New.docx");
            Assert(store.Resolve("C:\\Docs\\New.docx", 111) == blue,
                "Save As moves the colour onto the new path");
            Assert(store.Resolve("C:\\Docs\\Old.docx", 999) == blue,
                "the previous file keeps its colour (a copy, not a move)");
        }

        private static void TestSameAsColorDiffers()
        {
            var a = Tab(1, "a.docx");
            var b = Tab(1, "a.docx");
            b.TabColor = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);
            Assert(!a.SameAs(b), "gaining a colour is a difference (the strip must repaint)");
            a.TabColor = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);
            Assert(a.SameAs(b), "the same colour - the tabs are equal");
        }

        private static void TestThemeTints()
        {
            var c = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);
            var active = TabTheme.TintActiveBg(c);
            var inactive = TabTheme.TintInactiveBg(c);
            var hover = TabTheme.TintInactiveHoverBg(c);
            var border = TabTheme.TintBorder(c);

            Func<System.Drawing.Color, int> lum = x => x.R + x.G + x.B;
            Assert(lum(active) > lum(inactive), "the active pastel is lighter than the inactive one");
            Assert(lum(inactive) > lum(hover), "hover is richer (darker) than the ordinary inactive tab");
            Assert(lum(hover) > lum(c), "every pastel is lighter than the source colour");
            Assert(lum(border) > lum(c), "the border is lighter than the source (a pastel outline)");
            // The pastel is light enough for the amber indicators to stay readable on top
            Assert(active.R >= 200 && active.G >= 200 && active.B >= 200,
                "the active pastel is close to white (>=82% white)");
        }

        private static void TestTabStripColored()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using (var ctrl = new TabStripControl())
                    {
                        ctrl.Size = new System.Drawing.Size(700, 40);
                        var t1 = Tab(1, "Contract.docx", active: true, saved: false);
                        t1.TabColor = System.Drawing.Color.FromArgb(0x1A, 0x73, 0xE8);
                        var t2 = Tab(2, "Report.docx", saved: false);
                        t2.TabColor = System.Drawing.Color.FromArgb(0xD9, 0x30, 0x25);
                        var t3 = Tab(3, "Letter.docx");
                        ctrl.UpdateTabs(new List<DocumentTabModel> { t1, t2, t3 });

                        // Drawing into a bitmap (the full paint path) - nothing may throw
                        using (var bmp = new System.Drawing.Bitmap(ctrl.Width, ctrl.Height))
                            ctrl.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, ctrl.Width, ctrl.Height));

                        // A colour change => the model counts as changed => a repaint
                        var t1b = Tab(1, "Contract.docx", active: true, saved: false);
                        t1b.TabColor = null;
                        ctrl.UpdateTabs(new List<DocumentTabModel> { t1b, t2, t3 });
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("TabStripColored: " + failure.Message, failure);
        }

        private static void TestDpiOverride()
        {
            using (var ctrl = new TabStripControl())
            {
                ctrl.SetUserScale(1f); // a test of exact numbers must not depend on the machine tab-size.cfg
                int baseH = ctrl.DesiredContentHeight;

                ctrl.SetDpiOverride(144); // 150%
                Assert(ctrl.EffectiveDpi == 144, "EffectiveDpi equals the override that was passed");
                Assert(ctrl.DesiredContentHeight == 56, "the strip is 56 px at 150% (37*1.5)");

                ctrl.SetDpiOverride(192); // 200%
                Assert(ctrl.DesiredContentHeight == 74, "the strip is 74 px at 200%");

                ctrl.SetDpiOverride(20);  // junk - ignored, the scale is not lost
                Assert(ctrl.EffectiveDpi == 192, "a junk DPI was ignored");

                ctrl.SetDpiOverride(144);
                ctrl.SetCollapsed(true);
                Assert(ctrl.DesiredContentHeight == 24, "the collapsed sliver is 24 px at 150% (16*1.5)");
                ctrl.SetCollapsed(false);

                ctrl.SetDpiOverride(0);   // back to the control DeviceDpi
                Assert(ctrl.DesiredContentHeight == baseH, "0 returns the height from DeviceDpi");
            }
        }

        private static void TestSizeSettingsParse()
        {
            Assert(TabSizeSettings.Parse(new[] { "# comment", "scale=1.15" }) == 1.15f,
                "scale=1.15 is parsed");
            Assert(TabSizeSettings.Parse(new[] { "SCALE = 0.85" }) == 0.85f,
                "case and spaces do not matter");
            Assert(TabSizeSettings.Parse(new[] { "scale=junk" }) == TabSizeSettings.Default,
                "junk => the default multiplier");
            Assert(TabSizeSettings.Parse(null) == TabSizeSettings.Default, "null => the default");
            Assert(TabSizeSettings.Parse(new[] { "scale=99" }) == TabSizeSettings.Max,
                "too large => the maximum");
            Assert(TabSizeSettings.Parse(new[] { "scale=0.1" }) == TabSizeSettings.Min,
                "too small => the minimum");
            Assert(TabSizeSettings.Clamp(float.NaN) == TabSizeSettings.Default, "NaN => the default");
        }

        private static void TestUserScale()
        {
            using (var ctrl = new TabStripControl())
            {
                ctrl.SetDpiOverride(96);
                ctrl.SetUserScale(1f);
                Assert(ctrl.DesiredContentHeight == 37, "the base height is 37 at 100% / Normal");

                ctrl.SetUserScale(1.3f); // "Extra large"
                Assert(ctrl.UserScale == 1.3f, "the multiplier is remembered");
                Assert(ctrl.DesiredContentHeight == 48, "37 * 1.3 = 48 (the strip grew)");

                ctrl.SetDpiOverride(144); // a 150% monitor plus large tabs
                Assert(ctrl.DesiredContentHeight == 72, "37 * 1.5 * 1.3 = 72 (DPI and size multiply)");

                ctrl.SetUserScale(0.85f); // "Compact"
                Assert(ctrl.DesiredContentHeight == 47, "37 * 1.5 * 0.85 = 47");

                ctrl.SetUserScale(99f);  // clamp
                Assert(ctrl.UserScale == TabSizeSettings.Max, "the multiplier is capped at the maximum");
            }
        }

        // Bug regression: with the last tab active and scrolled left, tab captions
        // "escaped" the scroll zone onto the arrows and the buttons (TextRenderer
        // ignores the GDI+ clip; the row of tabs is now drawn into a buffer).
        private static void TestOverflowClipping()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using (var ctrl = new TabStripControl())
                    {
                        ctrl.SetDpiOverride(96);
                        ctrl.SetUserScale(1f);
                        ctrl.Size = new System.Drawing.Size(800, 37);

                        var many = new List<DocumentTabModel>();
                        for (int i = 1; i <= 24; i++)
                            many.Add(Tab(i, "Document" + i, active: i == 23));
                        ctrl.UpdateTabs(many);
                        using (var warm = new System.Drawing.Bitmap(800, 37))
                            ctrl.DrawToBitmap(warm, new System.Drawing.Rectangle(0, 0, 800, 37)); // creates the handle

                        // moving the active tab to the last one => EnsureActiveVisible => offset = max
                        var many2 = new List<DocumentTabModel>();
                        for (int i = 1; i <= 24; i++)
                            many2.Add(Tab(i, "Document" + i, active: i == 24));
                        ctrl.UpdateTabs(many2);

                        var geo = ctrl.OverflowGeometryForTests();
                        Assert(geo.Length == 3, "24 tabs in 800 px => scrolling mode");
                        var vp = geo[0];
                        var sl = geo[1];
                        var sr = geo[2];

                        var fld = typeof(TabStripControl).GetField("_scrollOffset",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        int max = (int)fld.GetValue(ctrl);
                        Assert(max > 240, "scrolled to the right edge (offset=" + max + ")");

                        // Reference: offset=max (everything fits exactly to the zone edge).
                        // Check: offset=max-240 - the tabs then reach past the zone.
                        // Everything outside the zone (except the arrows themselves, whose
                        // look legitimately depends on the scroll state) must stay pixel for pixel.
                        using (var bmpA = new System.Drawing.Bitmap(800, 37))
                        using (var bmpB = new System.Drawing.Bitmap(800, 37))
                        {
                            ctrl.DrawToBitmap(bmpA, new System.Drawing.Rectangle(0, 0, 800, 37));
                            fld.SetValue(ctrl, max - 240);
                            ctrl.DrawToBitmap(bmpB, new System.Drawing.Rectangle(0, 0, 800, 37));

                            int diffs = 0;
                            for (int x = 0; x < 800; x++)
                            {
                                bool insideViewport = x >= vp.Left && x < vp.Right;
                                bool onArrow = (x >= sl.Left - 2 && x <= sl.Right + 2) ||
                                               (x >= sr.Left - 2 && x <= sr.Right + 2);
                                if (insideViewport || onArrow) continue;
                                for (int y = 0; y < 37; y++)
                                {
                                    if (bmpA.GetPixel(x, y) != bmpB.GetPixel(x, y)) diffs++;
                                }
                            }
                            Assert(diffs == 0,
                                "outside the scroll zone the image does not depend on scrolling (differences: " + diffs + ")");
                        }
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("OverflowClipping: " + failure.Message, failure);
        }

        // Regression: inside the scroll-zone buffer the captions were drawn WITHOUT the
        // coordinate shift (TextRenderer ignores TranslateTransform) and "slid" onto the
        // neighbouring tabs; the caption of the active tab (blue) fell outside the zone.
        // Check: at offset=max the blue text of the active tab is visible INSIDE the zone,
        // near its right edge, and not beyond it.
        private static void TestOverflowTextPlacement()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using (var ctrl = new TabStripControl())
                    {
                        ctrl.SetDpiOverride(96);
                        ctrl.SetUserScale(1f);
                        ctrl.Size = new System.Drawing.Size(800, 37);

                        var many = new List<DocumentTabModel>();
                        for (int i = 1; i <= 24; i++)
                            many.Add(Tab(i, "Document" + i, active: i == 23));
                        ctrl.UpdateTabs(many);
                        using (var warm = new System.Drawing.Bitmap(800, 37))
                            ctrl.DrawToBitmap(warm, new System.Drawing.Rectangle(0, 0, 800, 37));

                        var many2 = new List<DocumentTabModel>();
                        for (int i = 1; i <= 24; i++)
                            many2.Add(Tab(i, "Document" + i, active: i == 24));
                        ctrl.UpdateTabs(many2); // the active tab is the last one, offset = max

                        var geo = ctrl.OverflowGeometryForTests();
                        Assert(geo.Length == 3, "scrolling mode is active");
                        var vp = geo[0];

                        using (var bmp = new System.Drawing.Bitmap(800, 37))
                        {
                            ctrl.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, 800, 37));
                            int blueInside = 0;
                            for (int x = vp.Left; x < vp.Right; x++)
                            {
                                for (int y = 6; y < 35; y++)
                                {
                                    var c = bmp.GetPixel(x, y);
                                    if (c.B > c.R + 40 && c.B > 140) blueInside++;
                                }
                            }
                            Assert(blueInside > 40,
                                "the caption of the active tab is visible inside the scroll zone (blue pixels: " +
                                blueInside + ")");
                        }
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("OverflowTextPlacement: " + failure.Message, failure);
        }

        // ------------------------------------------------------------------
        // Pinned tabs and stored order
        // ------------------------------------------------------------------

        private static void TestSameAsPinDiffers()
        {
            var a = Tab(1, "a.docx");
            var b = Tab(1, "a.docx", pinned: true);
            Assert(!a.SameAs(b), "pinning is a difference (the strip must repaint and move the tab)");
            a.IsPinned = true;
            Assert(a.SameAs(b), "the same pinning - the tabs are equal");
        }

        private static void TestOrderStoreParseFormat()
        {
            var parsed = TabOrderStore.ParseLines(new[]
            {
                "# comment",
                "",
                "PIN|C:\\Docs\\Contract.docx",
                "C:\\Docs\\Report.docx",
                "   ",
                "PIN|"
            });
            Assert(parsed.Count == 2, "exactly 2 entries were parsed (an empty PIN| with no path is dropped), not " + parsed.Count);
            Assert(parsed[0].Path == "C:\\Docs\\Contract.docx" && parsed[0].Pinned, "the first entry is pinned");
            Assert(parsed[1].Path == "C:\\Docs\\Report.docx" && !parsed[1].Pinned, "the second entry is not pinned");

            var lines = TabOrderStore.FormatLines(parsed);
            var reparsed = TabOrderStore.ParseLines(lines);
            Assert(reparsed.Count == parsed.Count, "format/parse roundtrip loses nothing");
            Assert(reparsed[0].Pinned && reparsed[0].Path == parsed[0].Path,
                "the pinning and the path survived the format/parse roundtrip");
        }

        private static void TestOrderStoreResolvePin()
        {
            var store = new TabOrderStore(null); // no file

            // A saved document: the key is the path; the hwnd does not matter
            store.SetPinned("C:\\Docs\\Contract.docx", 111, true);
            Assert(store.IsPinned("C:\\Docs\\CONTRACT.docx", 999),
                "the pinning of a saved document is found by path (case-insensitively)");

            // Unsaved: the key is the hwnd
            store.SetPinned("Document1", 222, true);
            Assert(store.IsPinned("Document1", 222), "the pinning of an unsaved document is found by hwnd");
            Assert(!store.IsPinned("Document1", 333), "a different hwnd is not pinned");

            // Unpinning, and cleanup when the window closes
            store.SetPinned("C:\\Docs\\Contract.docx", 111, false);
            Assert(!store.IsPinned("C:\\Docs\\Contract.docx", 111), "the pinning was removed");
            store.ForgetWindow(222);
            Assert(!store.IsPinned("Document1", 222), "closing the window forgets the session pinning");
        }

        private static void TestOrderStoreFileRoundtrip()
        {
            var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "TabsForWord-test-order-" + Guid.NewGuid().ToString("N") + ".cfg");
            try
            {
                var store1 = new TabOrderStore(file);
                var tabs = new List<DocumentTabModel>
                {
                    Tab(1, "Contract.docx", pinned: true, path: "C:\\Docs\\Contract for July.docx"),
                    Tab(2, "Report.docx", path: "C:\\Docs\\Report.docx")
                };
                store1.SaveOrder(tabs);
                Assert(System.IO.File.Exists(file), "the order file is created on the first save");

                var store2 = new TabOrderStore(file); // "Word restart"
                Assert(store2.IsPinned("C:\\Docs\\Contract for July.docx", 999),
                    "the pinning survived reloading the store (spaces in the path)");
                Assert(store2.RankOf("C:\\Docs\\Contract for July.docx") == 0, "first in order - rank 0");
                Assert(store2.RankOf("C:\\Docs\\Report.docx") == 1, "second in order - rank 1");
                Assert(store2.RankOf("C:\\Docs\\Unknown.docx") == null, "an unknown path - rank null");
            }
            finally
            {
                try { System.IO.File.Delete(file); } catch { }
            }
        }

        private static void TestOrderStoreRankAndSaveOrder()
        {
            var store = new TabOrderStore(null);
            var tabs = new List<DocumentTabModel>
            {
                Tab(1, "b.docx", path: "C:\\Docs\\b.docx"),
                Tab(2, "a.docx", path: "C:\\Docs\\a.docx"),
                Tab(3, "unsaved", path: "Document1") // unsaved - never reaches the file
            };
            store.SaveOrder(tabs);
            Assert(store.RankOf("C:\\Docs\\b.docx") == 0, "b.docx is first");
            Assert(store.RankOf("C:\\Docs\\a.docx") == 1, "a.docx is second");
            Assert(store.RankOf("Document1") == null, "an unsaved document does not enter the stored order");

            // A repeated SaveOrder with a different order fully replaces the previous one
            var reordered = new List<DocumentTabModel>
            {
                Tab(2, "a.docx", path: "C:\\Docs\\a.docx"),
                Tab(1, "b.docx", path: "C:\\Docs\\b.docx")
            };
            store.SaveOrder(reordered);
            Assert(store.RankOf("C:\\Docs\\a.docx") == 0, "after reordering a.docx is first");
            Assert(store.RankOf("C:\\Docs\\b.docx") == 1, "after reordering b.docx is second");
        }

        // Regression: before quitting Word the user closes every document (or Reconcile
        // simply sees an emptied _tabs) - DocumentWindowManager then calls SaveOrder(_tabs)
        // with an EMPTY list. SaveOrder used to replace _orderedPaths with that empty list
        // unconditionally and wiped the pinning and the order of documents the user opens
        // next time (the pinning in memory - _pinnedPaths - was left alone, but Save() wrote
        // the INTERSECTION with the already emptied _orderedPaths to disk, that is nothing).
        private static void TestOrderStoreSurvivesAllTabsClosed()
        {
            var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "TabsForWord-test-order-closeall-" + Guid.NewGuid().ToString("N") + ".cfg");
            try
            {
                var store = new TabOrderStore(file);
                var tabs = new List<DocumentTabModel>
                {
                    Tab(1, "Contract.docx", pinned: true, path: "C:\\Docs\\Contract.docx"),
                    Tab(2, "Report.docx", path: "C:\\Docs\\Report.docx")
                };
                store.SaveOrder(tabs);

                // Every tab is closed (as before Quit) - Reconcile calls SaveOrder with an
                // empty list, while the documents themselves have not gone anywhere.
                store.SaveOrder(new List<DocumentTabModel>());

                var reloaded = new TabOrderStore(file); // "Word restart"
                Assert(reloaded.IsPinned("C:\\Docs\\Contract.docx", 999),
                    "the pinning survived a SaveOrder with an empty list (every tab closed before quitting)");
                Assert(reloaded.RankOf("C:\\Docs\\Contract.docx") == 0, "the order of the pinned document is kept");
                Assert(reloaded.RankOf("C:\\Docs\\Report.docx") == 1, "the order of the ordinary document is kept too");

                // The document is reopened and explicitly unpinned - this time the entry must go.
                var reopened = new List<DocumentTabModel>
                {
                    Tab(1, "Contract.docx", pinned: false, path: "C:\\Docs\\Contract.docx")
                };
                reloaded.SaveOrder(reopened);
                var reloaded2 = new TabOrderStore(file);
                Assert(!reloaded2.IsPinned("C:\\Docs\\Contract.docx", 999),
                    "explicitly unpinning a still-open document is written out");
            }
            finally
            {
                try { System.IO.File.Delete(file); } catch { }
            }
        }

        private static void TestOrderStoreMigratePin()
        {
            var store = new TabOrderStore(null);

            // An unsaved document was pinned and then saved (Ctrl+S)
            store.SetPinned("Document1", 111, true);
            store.HandlePathChanged(111, "Document1", "C:\\Docs\\New.docx");
            Assert(store.IsPinned("C:\\Docs\\New.docx", 111),
                "after the first save the pinning is found by path");

            // A document with no pinning - the migration creates nothing
            store.HandlePathChanged(222, "Document2", "C:\\Docs\\Other.docx");
            Assert(!store.IsPinned("C:\\Docs\\Other.docx", 222), "with no pinning the migration creates nothing");

            // Saving without changing the path (only its case) is a no-op
            store.HandlePathChanged(111, "C:\\Docs\\New.docx", "C:\\Docs\\NEW.docx");
            Assert(store.IsPinned("C:\\Docs\\New.docx", 111), "a change of case in the path does not count as a change of path");
        }

        private static void TestHotkeySettingsParse()
        {
            Assert(HotkeySettings.Parse(new[] { "# comment", "enabled=0" }) == false, "enabled=0 is parsed");
            Assert(HotkeySettings.Parse(new[] { "ENABLED = 1" }) == true, "case and spaces do not matter");
            Assert(HotkeySettings.Parse(new[] { "enabled=junk" }) == HotkeySettings.Default,
                "junk => the default value");
            Assert(HotkeySettings.Parse(null) == HotkeySettings.Default, "null => the default");
            Assert(HotkeySettings.Default, "Ctrl+Tab is on by default (the MVP behaviour is unchanged)");
        }

        private static void TestTabStripPinnedRender()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using (var ctrl = new TabStripControl())
                    {
                        ctrl.Size = new System.Drawing.Size(700, 40);
                        var t1 = Tab(1, "Important.docx", active: true, pinned: true);
                        var t2 = Tab(2, "Also important.docx", pinned: true, saved: false);
                        var t3 = Tab(3, "Ordinary.docx");
                        ctrl.UpdateTabs(new List<DocumentTabModel> { t1, t2, t3 });

                        using (var bmp = new System.Drawing.Bitmap(ctrl.Width, ctrl.Height))
                            ctrl.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, ctrl.Width, ctrl.Height));

                        // A narrow strip plus a mix of pinned and ordinary tabs - shrinking and
                        // overflow (the scroll-zone buffer) must not throw.
                        ctrl.Size = new System.Drawing.Size(420, ctrl.DesiredContentHeight);
                        var many = new List<DocumentTabModel>();
                        for (int i = 1; i <= 20; i++)
                            many.Add(Tab(i, "Document " + i, active: i == 10, pinned: i % 3 == 0));
                        ctrl.UpdateTabs(many);
                        using (var bmp2 = new System.Drawing.Bitmap(ctrl.Width, ctrl.Height))
                            ctrl.DrawToBitmap(bmp2, new System.Drawing.Rectangle(0, 0, ctrl.Width, ctrl.Height));
                    }
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("TabStripPinnedRender: " + failure.Message, failure);
        }

        private static void TestTabListPopupPinnedAndSettings()
        {
            Exception failure = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    bool settingsCalled = false;
                    var tabs = new List<DocumentTabModel>
                    {
                        Tab(1, "Important.docx", pinned: true),
                        Tab(2, "Ordinary.docx", active: true),
                        Tab(3, "More.docx", saved: false)
                    };
                    using (var popup = new TabListPopup(tabs, 1f, delegate { }, delegate { settingsCalled = true; }))
                    {
                        // the constructor with the "Settings..." footer and a pinned tab must not throw
                    }
                    Assert(!settingsCalled, "the settings callback does not fire by itself during construction");
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw new Exception("TabListPopupPinnedAndSettings: " + failure.Message, failure);
        }
    }
}
