# TEST_RESULTS.md

Date: 2026-07-23. Environment: Windows 11 Pro x64, Microsoft 365 Word x64
(16.0.20131.20154). That date and environment apply to the first sections (up to
stage 8); every later section names its own date, version and machine.

Statuses: PASS / FAIL / NOT TESTED / BLOCKED

## Automated tests (no Word) - scripts/test.ps1

| # | Scenario | Result | Comment |
|---|---|---|---|
| A1 | DocumentTabModel.SameAs: equal objects | PASS | |
| A2 | SameAs: a difference in Saved | PASS | the "*" indicator |
| A3 | ListsEqual: empty and equal lists | PASS | |
| A4 | ListsEqual: a different order => not equal | PASS | tab order is significant |
| A5 | ListsEqual: null tolerance | PASS | |
| A6 | PvKey: negative and stable | PASS | never collides with an HWND |
| A7 | PvKey: case-insensitive | PASS | |
| A8 | PvKey: null does not throw | PASS | |
| A9 | TabStripControl: construction, UpdateTabs, 15 tabs | PASS | no exceptions |
| A10 | LoggingService: writing an entry with an exception | PASS | |

Result: 10/10 PASS.

(As of 2026-07-23. The suite has grown with every stage since - each later
section gives its own count, the latest being 68/68 in stage 37. The "*"
unsaved indicator of A2 and B11 has since become a dot.)

## Integration tests on a real Word

Method: automating a live Word instance through COM + real mouse clicks on the
strip (SetCursorPos/mouse_event) + verification against the add-in log and
screenshots.

### The main scenarios (stages 4-6)

| # | Scenario | Result | Comment |
|---|---|---|---|
| B1 | The add-in loads when Word starts | PASS | "Add-in started" in the log, LoadBehavior stayed 3 |
| B2 | A tab for an already-open document (the startup scan) | PASS | Document1 is visible immediately |
| B3 | A strip in every Word window | PASS | 12 windows -> 12 strips |
| B4 | Opening a document -> a tab in every window | PASS | |
| B5 | Creating a document -> a tab | PASS | |
| B6 | Closing a document -> the tab disappears | PASS | plus "Dead tab pane removed" |
| B7 | Clicking a tab -> the window is activated | PASS | a real click; ActiveWindow changed |
| B8 | The active tab is visually highlighted | PASS | screenshots |
| B9 | The close button | PASS | a real click on the "x" |
| B10 | The "+" button -> a new document | PASS | a real click |
| B11 | The "*" unsaved indicator | PASS | appeared after an edit, disappeared after Save |
| B12 | Save As -> the tab caption updated | PASS | Document2 -> TestA.docx |
| B13 | The tooltip with the full path | NOT TESTED | the code is there; showing it was not automated |
| B14 | Overflow: the arrows scroll, [+] stays on the right, the UI does not break | PASS | 12 tabs, screenshots |
| B15 | Auto-scrolling to the active tab | PASS | Document13 visible after it was opened |

### Robustness scenarios (stage 7)

| # | Scenario | Result | Comment |
|---|---|---|---|
| C1 | A new empty document | PASS | |
| C2 | An ordinary DOCX | PASS | |
| C3 | A DOC (compatibility mode) | PASS | the tab "TestOld.doc" |
| C4 | A read-only document | PASS | the tab is there and works |
| C5 | Several documents | PASS | |
| C6 | 10+ documents (12) | PASS | overflow + scrolling |
| C7 | Save | PASS | the "*" was cleared |
| C8 | Save As | PASS | the caption updated |
| C9 | Closing a saved document | PASS | no dialog |
| C10 | Closing an unsaved document | PASS | the standard Word dialog was shown |
| C11 | Cancelling the close in the Word dialog | PASS | Esc -> "Close cancelled by user", the tab stayed |
| C12 | Switching with Alt+Tab (an external activation) | PASS | the active tab updated |
| C13 | Switching with Word's own means | PASS | Window.Activate |
| C14 | Closing the active document | PASS | |
| C15 | Closing an inactive document | PASS | a mass close of 11 background ones |
| C16 | Quitting Word | PASS | a correct OnDisconnection, "Add-in stopped" |
| C17 | Starting Word again | PASS | the strips were restored |

### Additional scenarios

| # | Scenario | Result | Comment |
|---|---|---|---|
| D1 | Protected View: a file with the Mark-of-the-Web | PASS | a grey [PV] tab; there is no strip inside the PV window (an Office limitation) |
| D2 | PV -> "Enable Editing" | PASS | the tab converts from PV to an ordinary one; a NullReference bug was found and fixed (see KNOWN_ISSUES) |
| D3 | One document in two windows ("New Window") | PASS | the tabs "TestA.docx:1" and "TestA.docx:2" |
| D4 | The close button on ":2" closes only that window | PASS | the document stays open |
| D5 | The add-in does not interfere with Word's dialogs | PASS | the standard save dialog works normally |
| D6 | Word neither crashes nor hangs across the whole test session | PASS | LoadBehavior=3, DisabledItems empty |

## Not tested / testing limitations

As of 2026-07-23 (stage 8). Several items were covered later: multi-DPI in the
per-monitor DPI sections (stage 15 and the three-monitor run), Word 2024 in the
field (machines A and B, stage 23). Still NOT TESTED: Word 2016/2019/2021 and
32-bit Word.

- The tooltip (B13): the code and the control's unit test were checked, but the
  popup itself was never captured in a screenshot.
- Multiple monitors / DPI scaling other than 100 %: NOT TESTED (one machine).
- Running alongside add-ins other than OneNote/PDFMaker: those are present on the
  machine and no conflicts were observed.
- Word 2021/2024 (not M365): NOT TESTED - the machine only has M365.

## Tab colour (stage 13, 2026-07-24, v1.2.0)

Automated tests: 33/33 PASS (6 new: parsing and formatting tab-colors.cfg, the
path and hwnd keys, a file round-trip, SameAs by TabColor, the pastel tones,
drawing coloured tabs).

E2E on a real Word M365 x64 (native+reserve,
tests/manual/e2e-tab-colors.ps1):

| # | Scenario | Result | Comment |
|---|---|---|---|
| E1 | Right-click on a tab -> the menu opened | PASS | the classic Win32 menu |
| E2 | The "Tab colour" submenu -> "Red" | PASS | the log says "Tab color set: D93025" |
| E3 | The menu closed after the choice | PASS | |
| E4 | The fill became pastel red | PASS | the pixel went 235,237,239 -> 241,176,172 |
| E5 | The colour was written to tab-colors.cfg | PASS | the key is the full file path |
| E6 | "Reset colour" | PASS | the log says "Tab color cleared", the entry was removed from the file |
| E7 | Setting a colour again (green) | PASS | the log says "Tab color set: 1E8E3E" |
| E8 | Restarting Word: the colour was loaded | PASS | the log says "Tab colors loaded: 1" |
| E9 | Restarting Word: the tab is pastel green | PASS | the pixel is 214,235,220 |
| E10 | WINWORD exited cleanly, no ERROR in the log | PASS | both sessions |

The colour of an unsaved document (the first version of the run coloured the
initial "Document1"): the pastel is applied but nothing is written to the file -
it lives in memory until the document is saved or closed, exactly as designed.

"More colours…" (tests/manual/e2e-tab-colors-dialog.ps1): the system ChooseColor
opened inside Word, OK applied the colour (the log says "Tab color set"), the
entry in tab-colors.cfg was created, and Word exited cleanly - 7/7 PASS.

## The installer and diagnostics (stage 14, 2026-07-24, v1.3.0)

| # | Check | Result | Comment |
|---|---|---|---|
| F1 | install.ps1: the installation log was created and the environment recorded | PASS | Logs\install-20260724-185134.log |
| F2 | Registration in the 64-bit registry view | PASS | the installer reads every key back itself |
| F3 | Registration in the 32-bit view (WOW6432Node) | PASS | CodeBase/Assembly are read from Registry32 |
| F4 | The assembly name in the registry = the real one from the DLL (1.3.0.0) | PASS | it used to be hard-coded as 1.0.0.0 |
| F5 | diagnose.ps1: a ZIP on the desktop with every section | PASS | a 59 KB report plus the Logs folder |
| F6 | The diagnostics summary finds event-log errors | PASS | it found WINWORD entries from earlier crash tests |
| F7 | The "Connect instantiated" beacon and the Environment line in the log | PASS | v1.3.0.0, Word 16.0.20131, ru-RU |
| F8 | The full tab-colour E2E on v1.3.0 | PASS | 15/15, Word exited cleanly, 0 ERROR |

NOT TESTED (no machine available): a real 32-bit Word - the WOW6432Node
registration was verified by reading the registry, and loading under 32-bit Word
is supported by reasoning only (mscoree + AnyCPU). A diagnostics report from
someone else's machine will show the actual state.

## Per-monitor DPI (stage 15, 2026-07-25, v1.3.1)

Configuration: monitor 1 - 1920x1080, scale 150 % (DPI 144, the system one);
monitor 2 - 2560x1440, scale 100 % (DPI 96). The Word window is moved by the
script tests/manual/e2e-dpi-monitors.ps1; the drawn height is measured from the
bottom line of the strip on a screenshot (physical pixels).

| # | Check | Before the fix | After | Result |
|---|---|---|---|---|
| G1 | The DPI of the Word window on the monitors (144/96) | OK | OK | PASS |
| G2 | The strip height at 150 % (expected 56) | 56 | 56 | PASS |
| G3 | The drawn tabs at 150 % (expected 56) | 37 (too small) | 56 | PASS |
| G4 | The strip height at 100 % (expected 37) | 25 (clipped) | 37 | PASS |
| G5 | The drawn tabs at 100 % (expected 37) | 25 | 37 | PASS |
| G6 | Moving back to 150 % restores the size | OK | OK | PASS |
| G7 | Word exited cleanly, no ERROR in the log | OK | OK | PASS |
| G8 | Regression: the full tab-colour E2E at 150 % | - | 15/15 | PASS |

Automated tests 34/34 (new: the DPI override scales DesiredContentHeight, junk
values are ignored, a collapsed strip scales too).

A quirk (documented in ADR-010): on a monitor whose scale differs from the system
one, the strip is scaled by the system - the sizes are exact but the font is
slightly softer than Word's own text. NOT TESTED: changing a monitor's scale on
the fly without moving the window (a clean check would need signing out and back
in); it is expected to be covered by the same logic (the roughly 1 Hz
reconciliation).

## Manual tab size (stage 16, 2026-07-25, v1.4.0)

| # | Check | Result | Comment |
|---|---|---|---|
| H1 | The menu: right-click -> "Tab size" -> "Large" | PASS | the strip went 56 -> 64 px (x1.15 on a 150 % monitor) |
| H2 | The multiplier was written (the log + tab-size.cfg) | PASS | "Tab size scale saved: 1.15", scale=1.15 |
| H3 | Restarting Word: the size persisted | PASS | the strip is 64 px again |
| H4 | Going back to "Normal" | PASS | 64 -> 56 px |
| H5 | Word exited cleanly, no ERROR in the log | PASS | both sessions |
| H6 | Regression: the full colour E2E after the menu changed | PASS | 15/15 |

Automated tests 36/36 (new: parsing and clamping tab-size.cfg; the multiplier on
top of the DPI override - 37 -> 48 at x1.3, 72 at 150 % + x1.3, 47 at 150 % +
x0.85). The "Normal"/"Large" screenshots were compared visually
(e2e6-size-*.png).

## Field fixes (stage 17, 2026-07-25, v1.4.1)

| # | Check | Result | Comment |
|---|---|---|---|
| I1 | The Navigation pane: the side dock appeared | PASS | MsoCommandBarDock 0,178 350x832 - at its natural top |
| I2 | The tab strip shrank to the document area | PASS | 0..1920 -> 350..1920, exactly along _WwF |
| I3 | The tab strip does NOT overlap the side dock | PASS | the rectangles do not intersect |
| I4 | Closing the side pane restores the full width | PASS | 1570 -> 1920 |
| I5 | The collapse and all-tabs buttons: centred, larger | PASS | the zoomed screenshot e2e7-buttons-zoom.png was compared visually |
| I6 | Reserve mode after the fit was sped up (125 ms) | PASS | e2e-reserve-mode 11/11: resize, the ribbon, no oscillation |
| I7 | Tab-colour regression | PASS | 15/15 (one run had a single failed step; two full repeats had none) |
| I8 | Word exited cleanly, no ERROR in the logs | PASS | every run |

Automated tests 37/37 (new: ComputePanelRect with side panes - left, right,
clamps, degenerate cases). The window-structure research is
probe-side-panes.ps1 (Word's side panes are MsoCommandBarDock and start at the
natural top of the document area; _WwF shrinks in width when they open).

The Editor and Transcribe panes were never opened live (there is no reliable
programmatic way: none of the ExecuteMso candidates fit) - their geometry is the
same as Navigation's (MsoCommandBarDock), which was covered by E2E. If a user
still sees an overlap, send a screenshot and a case will be added.

## Optimising the test harness (stage 18, 2026-07-25, the add-in code was unchanged)

| # | Check | Result | Comment |
|---|---|---|---|
| J1 | The new smoke test e2e-smoke.ps1 | PASS 10/10 | 28 seconds, one Word run |
| J2 | e2e-tab-colors.ps1 after moving to menu polling | PASS 15/15 | 51 s (was about 2 min) |
| J3 | e2e-tab-size.ps1 after moving to polling | PASS 8/8 | 46 s (was about 2 min) |
| J4 | e2e-tab-colors-dialog.ps1 after moving to polling | PASS 7/7 | 29 s (was about 1 min) |
| J5 | A BOM on every .ps1 with Cyrillic in tests/manual | PASS | protection against PowerShell 5.1 |

A mandatory run matrix was introduced (docs/TESTING.md): a feature's own E2E
runs only for the feature that was touched; after
any change to the rendering, at least the smoke test; the full set before
rebuilding the packages for distribution. The flake policy: one repeat of a
script; two failures in a row count as a bug rather than something to re-run.

## Scrolling the tabs on overflow (stage 19, 2026-07-25, v1.4.2)

Repro: 24 documents, a 1920 px window => the arrow mode; the active tab is the
last one, and scrolling left happens by clicking the left arrow. Before the fix
the tab captions were drawn outside the scroll zone - over the right arrow, over
"+", right up to the right edge (cause: TextRenderer/GDI ignores the GDI+ clip;
the fills were cut, the text was not).

| # | Check | Result | Comment |
|---|---|---|---|
| K1 | Clicking the left arrow from the active last tab: no artefacts on the right | PASS | bluePx 29 -> 0 (the button zone is 105 px) |
| K2 | Touchpad: WM_MOUSEHWHEEL scrolls the tabs | PASS | smoothly and proportionally to delta |
| K3 | Touchpad: returning to the stop restores the picture | PASS | pixel for pixel (MD5) |
| K4 | No artefacts after touchpad scrolling | PASS | bluePx=0 |
| K5 | The arrows are vectors and are noticeable | PASS | the zoomed screenshot was compared visually |
| K6 | Word exited cleanly, no ERROR in the log | PASS | 24 documents opened and closed |
| K7 | The smoke test on the final binary | PASS | 10/10 |

Automated tests 38/38; the new unit test: two renders (offset=max and max-240)
must be pixel-for-pixel identical outside the scroll zone (apart from the arrows
themselves - their appearance legitimately depends on the state). The Word-free
rig (STA + reflection) is preserved in the session history; the key diagnostic
tool is a direct dump of the layout.

## Polishing the scroll buffer (stage 20, 2026-07-25, v1.4.3)

From a screenshot sent by the user: with many tabs the captions ran over the
neighbouring tabs and were cut without an ellipsis. The cause is that stage 19
drew the row of tabs into a buffer through TranslateTransform, and TextRenderer
(GDI) ignores a Graphics transform exactly as it ignores a clip: the fills moved,
the text did not.

| # | Check | Result | Comment |
|---|---|---|---|
| L1 | Captions on their own tabs, with an ellipsis | PASS | an explicit shift of the bounds instead of a transform |
| L2 | Unit: the blue caption of the active tab is visible inside the zone | PASS | before the fix the text fell outside the zone |
| L3 | A dragged tab: its caption is drawn | PASS | the same defect: TranslateTransform in its bitmap |
| L4 | e2e-overflow-scroll on the final binary | PASS | 8/8 |
| L5 | smoke on the final binary | PASS | 10/10 |

Automated tests 39/39. A lesson for future edits: NEVER rely on
Graphics.Clip/TranslateTransform where TextRenderer does the drawing - only
explicit coordinates (also recorded in the project memory).

## Crisp tab text in scrolling mode (stage 21, 2026-07-25, v1.4.4)

The user noticed that with many tabs the font looked "bold and fuzzy". Confirmed
by zooming in on the rig: GDI text (TextRenderer) on a bitmap surface is drawn
without antialiasing; an opaque buffer does NOT help (verified separately).

| # | Check | Result | Comment |
|---|---|---|---|
| M1 | Captions in scrolling mode = the ordinary ones (zoom comparison) | PASS | a second pass over the screen DC |
| M2 | Clipping the caption to the zone (the rect bounds passed to DrawText) | PASS | unit: the renders are identical outside the zone |
| M3 | The caption of a floating tab while dragging | PASS | also onto the screen DC |
| M4 | e2e-overflow-scroll on v1.4.4 | PASS | 8/8 |
| M5 | smoke on v1.4.4 | PASS | 10/10 |

Automated tests 39/39. Along the way THE SMOKE TEST was fixed: after a click on a
tab a DIFFERENT Word window becomes active (SDI), possibly in windowed mode - the
test was aiming at the coordinates of the old window and kept "missing" the menu
(the user spotted it while watching a run); now the test switches to the new
active window and maximises it.

## Fixes from external audit 2 (2026-07-25, v1.4.5)

Fixed: the tab colour disappearing after the first save (migrating the key from
hwnd to path, and from path to path on Save As), the order in which the
_initialized flag is set in Connect (now after full initialisation, with a
rollback on failure), the atomicity of subscribing to and unsubscribing from Word
events, counting native-host layout failures towards the CTP fallback counter,
the check for a closed Word in repair.ps1, and the selective decoding of
Resiliency in diagnose.ps1.

| # | Check | Result | Comment |
|---|---|---|---|
| N1 | Automated tests (2 new: colour migration) | PASS | 41/41 |
| N2 | e2e-smoke on v1.4.5 | PASS | 10/10 |
| N3 | e2e-tab-colors on v1.4.5 | PASS | 15/15; the log shows the transient "no content zone" at startup with no false fallback |
| N4 | e2e-reserve-mode on v1.4.5 | PASS | 11/11 |
| N5 | The live scenario "colour an unsaved document -> Ctrl+S -> the colour stays" | NOT TESTED | covered by unit tests (the migration) plus e2e-tab-colors (colours in general); there is no separate E2E step that saves |
| N6 | A live CTP fallback when the zone is never found | NOT TESTED | it needs a Word with a changed window tree - nothing to simulate that with; the counter logic was verified by reading the code and by the locator unit tests |
| N7 | repair.ps1, an actual run (Word closed) | PASS | LoadBehavior=3, the Resiliency scan, an honest "our add-in is not there" |
| N8 | diagnose.ps1, an actual run | PASS | Resiliency "empty", the warning about the ZIP contents is printed; the test ZIP was deleted |
| N9 | The repair.ps1 branch "Word is running -> refuse" | NOT TESTED | the block repeats byte for byte the long-working guard in uninstall.ps1 |

The targeted Resiliency cleanup (only the TabsForWord entries) was verified
separately on a temporary registry key with fake "ours"/"someone else's" entries
(checkpoint-22).

## Stage 22 - pinning tabs, persisting the order, the settings window (2026-07-25, v1.5.0)

A multi-agent adversarial review (3 lenses: correctness, safety for Word, design
consistency, plus a re-check of the findings) found 5 confirmed problems; 2 of
them are real bugs that a live E2E run caught independently (losing the pinning
when all tabs are closed before quitting; the repaint being skipped when the model
is mutated in place) - all 5 were fixed before the commit.

| # | Check | Result | Comment |
|---|---|---|---|
| O1 | Automated tests (9 new: TabOrderStore, HotkeySettings, pinned rendering) | PASS | 51/51 |
| O2 | e2e-smoke on v1.5.0 | PASS | 10/10 |
| O3 | e2e-settings-pins-order.ps1 (new, the full scenario) | PASS | 26/26; pin, unpin and pin again through the menu, moving a tab to the front, the order and the pinning survive a Word restart EVEN when the documents are opened in a different order, the settings window (size + Ctrl+Tab) opens from the footer of the all-tabs menu and applies instantly (verified functionally in both directions - Ctrl+Tab really does stop and start working) |
| O4 | e2e-tab-colors.ps1 (regression - the new "Pin" item shifted the menu's Down navigation) | PASS | 14/14, the script was updated (Down x3 -> x4) |
| O5 | e2e-tab-size.ps1 (the same regression) | PASS | 8/8, the script was updated (Down x4 -> x5) |
| O6 | e2e-reserve-mode.ps1 | PASS | 11/11 |
| O7 | The pin separator in scrolling mode (visually) | NOT TESTED | fixed in the code (the drawing was added to the overflow buffer, the same trick as for the pin icon), but no separate E2E with the scenario "20+ tabs, some pinned, scrolling" exists - it is only covered by an automated test that no exception is thrown |
| O8 | RadioButton/CheckBox in SettingsForm - the visual theme (not the classic look) | PASS (by inspection) | the screenshot e2e8-settings-open.png was taken during E2E; the window opens and paints with no exceptions in the log |

The first run of e2e-settings-pins-order.ps1 (before the fixes) caught a
production-code bug outright: `Tab order loaded: 0 (pinned: 0)` after a restart
even though the document had been pinned and saved in the previous session -
SaveOrder was wiping the file when all the tabs were closed. After the fix
(TabOrderStore.SaveOrder appends rather than replaces the paths outside the
current set) it reads `Tab order loaded: 3 (pinned: 1)`, and the pinned document
is correctly first even when it was opened last.

## Multi-DPI on three monitors (2026-07-25, the add-in code was unchanged)

Three monitors connected SIMULTANEOUSLY at DIFFERENT scales: 100 % (DPI 96),
125 % (DPI 120), 150 % (DPI 144). The Word window is moved by the script
tests/manual/e2e-dpi-monitors.ps1 (rewritten: it auto-discovers every monitor
through EnumDisplayMonitors + GetDpiForMonitor, takes the monitor with the lowest
DPI as the baseline and checks the scaling relative to it, so the test is not tied
to a particular layout). The height is measured from the host frame
(GetWindowRect) and from the bottom line of the strip on a screenshot (physical
pixels).

| # | Check | Result | Comment |
|---|---|---|---|
| P1 | The Word window DPI = the monitor DPI (96/120/144) | PASS | per-monitor DPI on every screen |
| P2 | The strip height scales with the DPI | PASS | 43 -> 54 -> 65 px (exactly x1.25 and x1.5 of the 43 baseline) |
| P3 | The drawn content scales | PASS | 43 -> 53 -> 63 px (a scan of the bottom line) |
| P4 | Moving back to the baseline monitor restores the size | PASS | 43 -> 43 px |
| P5 | Word exited cleanly, no ERROR in the log | PASS | the reserve log: anchor moved 43 -> 54 -> 64 -> 43 |

Result: 13/13 PASS. The baseline strip height at 100 % is now 43 px (not the 37
in section 9 of NATIVE_TAB_HOST.md) because of the user tab-size setting
(_userScale, stage 16) - the DPI scaling is orthogonal to it. A capture flake (a
screenshot of a not-yet-painted strip when the mouse moves or the focus changes
during a run) was closed by repeating the measurement inside the script. This
closes item 6 of section 11 of NATIVE_TAB_HOST.md (previously NOT TESTED: a second
monitor at a different scale was physically unavailable).

## The installer's prerequisite checks: .NET and the presence of Word (2026-07-26, the add-in code was unchanged)

Changes following external audit 3 (ADR-014): a missing .NET 4.8 became a fatal
installation error instead of a false "success"; the detection of Word was
strengthened to three independent signals; the same threshold, link and
explanations were duplicated in diagnose.ps1. ONLY the installer scripts and the
documentation changed - TabsForWord.dll was not rebuilt (its SHA256 matches
src\bin\Release), so the automated tests and the feature E2E runs of the
docs/TESTING.md matrix are unaffected and were not re-run.

| # | Check | Result | Comment |
|---|---|---|---|
| Q1 | The syntax of install.ps1 and diagnose.ps1 (the AST parser) | PASS | no parser errors in either |
| Q2 | The structure of install.ps1: 8 functions, no duplicates, 7 exit points | PASS | 5 fatal refusals + the read-back + success |
| Q3 | Find-WordInfo on this machine: Word found | PASS | App Paths, x64, 16.0.20131.20154, C2R O365HomePremRetail |
| Q4 | Signal 2 (winword.exe in the Office folders) - on its own | PASS | `...\Microsoft Office\root\Office16\WINWORD.EXE` |
| Q5 | Signal 3 (the Word.Application COM registration) - on its own | PASS | found in LocalMachine/Registry64 and /Registry32 |
| Q6 | The three branches of the .NET check (the logic) | PASS | an old version -> refusal; no .NET -> refusal; the registry unreadable but the files present -> a warning |
| Q7 | The real state of .NET on the machine passes the check | PASS | version 4.8.09221 (Release=533509) -> the installation continues |
| Q8 | A human-readable version instead of Release=NNNNNN | PASS | the `Version` value is taken from the same key |
| Q9 | The text of the fatal message (rendering) | PASS | the link, the clipboard, the hint about the Windows build, the note about administrator rights, the numbered instruction 1-4 |
| Q10 | The .NET download link is live and localised | PASS | HTTP 200, the Russian page title "Download .NET Framework 4.8" |
| Q11 | Synchronisation of the 2 folders + 2 ZIPs (install/diagnose/INSTALL_RU) | PASS | 12 items by SHA256, no discrepancies |
| Q12 | The DLL in the packages did not change | PASS | the hash matches src\TabsForWord\bin\Release |

Result: 12/12 PASS.

NOT TESTED (no suitable machine - verified by simulating the logic, not live):
- **a machine without .NET 4.8 or with an old version** - both fatal branches of
  Q6 were checked by running the logic with substituted values rather than on a
  real system (compare limitation O16: as of 2026-07-26 Windows 10 had never
  been run at all; machine C gave the first, clean, Windows 10 field report on
  2026-07-28 - see stage 34);
- **a machine with no Word installed** - the fatal "Word not found" branch was
  never exercised live; on this machine all three signals fired independently
  (Q3-Q5), so only the "Word is present" scenario is confirmed;
- **the hint about Windows builds below 18362** - it was rendered with a
  substituted build number 17763; no real old Windows was available.

NOT TESTED (deliberately not run):
- **an end-to-end run of install.ps1** - the installer was not launched: that
  would re-register the add-in and requires Word to be completely closed;
- **interactive input in a real console** - `Read-Host` (the "open the browser?"
  prompt), `Set-Clipboard` and `Start-Process` were replaced by stubs while
  rendering Q9; all three are wrapped in try/catch, so their failure does not
  affect the exit code, but the actual behaviour in a live console was never
  observed.

## Field reports from two machines + fixes (2026-07-26, v1.5.1, stage 23)

Input: two Diagnostics ZIP reports (machines A and B, both Windows 11 Pro 25H2
26200.8875, Word x64 16.0.17932.20884 ProPlus2024, DLL v1.5.0 with the same
SHA256, DPI 150 %). The reports themselves said "Windows 10 Pro 25H2":
Diagnostics up to 1.5 printed the registry ProductName, which still reads
"Windows 10" on Windows 11; build 26200 is Windows 11 25H2 (Diagnostics 1.6
labels build 22000 and later as Windows 11, see stage 37). Both installations succeeded, the strip worked, 0 ERROR, Word
did not crash, and Resiliency never disabled our add-in.

| # | Check | Result | Details |
|---|---|---|---|
| F1 | Installation on someone else's machine (machine A) | PASS | 4/4 registry read-backs, LoadBehavior=3 |
| F2 | Installation on someone else's machine (machine B) | PASS | the same |
| F3 | The strip appears and the tabs work | PASS | machine A: 5 windows; machine B: 3 windows |
| F4 | Size / colour / pinning / collapsing in the field | PASS | machine A went through all four |
| F5 | A clean unload when Word quits | PASS | machine B: the hooks were removed, the hosts destroyed |
| F6 | The position of the strip in the first Word window | **FAIL -> fixed** | machine A: rect=0,0 (over the ribbon), the anchor was a half-built canvas |
| F7 | The delay before the strip appears at startup | PASS (with a note) | machine B: 1 locator miss, the strip appeared after about 1 s |

The fixes and their verification (build v1.5.1):

| # | Check | Result | Details |
|---|---|---|---|
| R1 | A half-built Word window is rejected by the locator | PASS | a new unit test built on the measurements from machine A (0,0,2534x1528) |
| R2 | After the layout the anchor is `_WwF` | PASS | the same test, second phase |
| R3 | The Degraded mark on the fallback routes | PASS | a unit test (the anchor is the canvas; a fallback with no canvas) |
| R4 | Focus returns to the document after clicking a tab | PASS | a probe on a live Word: the character arrived |
| R5 | Focus returns to the document after clicking the collapse button | **FAIL -> fixed** | before the fix the character was lost; after it, it arrives |
| R6 | Automated tests | PASS | 53/53 (+2 new) |
| R7 | The full E2E set on v1.5.1 | PASS | smoke 10/10, reserve 11/11, overlay-interaction 9/9, overlay-buttons-pv 15/15, colors 15/15, colors-dialog 7/7, size 8/8, settings-pins-order all OK, side-panes 7/7, overflow-scroll 7/7, dpi-monitors 9/9 |
| R8 | The packages were rebuilt | PASS | both folders + the ZIP on DLL v1.5.1 |

Defects in the test harness found by the same run (fixed):
- three scripts changed `native-host.cfg` and did not restore the user's mode -
  after a run the strip stayed in overlay and covered the ruler;
- the check "the character reached the document" compared against 'x' and failed
  under a Russian keyboard layout (VK 'X' produces a Cyrillic letter) - it now
  compares the text length;
- collapsing the ribbon in e2e-reserve-mode was done by double-clicking (120,103)
  and depended on the ribbon state left over from previous runs - it was replaced
  by ExecuteMso('MinimizeRibbon') after forcing a known state;
- e2e-overlay-buttons-pv relied on a test1.docx left over from previous runs - it
  now creates the fixture itself.

NOT TESTED: reproducing the F6 race on demand - the moment when Word hands out a
half-built window tree cannot be controlled from outside; the fix is covered by a
unit test on the coordinates captured from machine A and by the absence of
regressions across the full E2E set.

## Diagnostics v1.2 - parsing the add-in log (2026-07-26, the add-in code was unchanged)

The reason: on both field machines the automatic summary said "No obvious problems
found", even though machine B's log had warnings and machine A's log had a line
about the strip being placed at the very top of the window. The summary only
looked at the registry, the presence of logs and the Windows event log - it never
read the add-in's own log.

How the field data was verified: the logs of machines A and B were placed in a
separate folder, `LOCALAPPDATA` was temporarily pointed at it, and the REAL
`installer/diagnose.ps1` was run over them (not a copy of the logic). The user's
own working folder was not touched.

| # | Check | Result | Details |
|---|---|---|---|
| D1 | The script runs without errors and its encoding is intact | PASS | the UTF-8 BOM `EF BB BF` is in place |
| D2 | ERROR lines are grouped by kind | PASS | the local logs: `x5 Test error marker` |
| D3 | Transient WARNs do not raise a false alarm | PASS | 91 x "no content zone" - the "routine" section, never reaching the SUMMARY |
| D4 | Window-tree dumps are flagged as a problem | PASS | 8 dumps -> an item in the SUMMARY |
| D5 | The strip at the very top of the window is caught (machine A's case) | PASS | machine A's log -> the summary reported the strip having been placed at the very top of the window once |
| D6 | On a healthy log the D5 check stays silent | PASS | machine B's logs and the local ones - the item did not appear |
| D7 | The add-in version is taken from the log | PASS | machines A and B -> 1.5.0.0; locally -> 1.5.1.0 |
| D8 | "The last run" really is the last one | PASS | the files are parsed from oldest to newest (there was a bug: it showed yesterday's) |
| D9 | The hwnd is masked in the grouping key | PASS | `root=HWND (nodes=N)`; document names never reach the summary |
| D10 | The three copies of diagnose.ps1 match | PASS | SHA256 `7657DBF3…` in installer/ and in both packages |

Deliberate decisions: the TWO most recent daily logs are parsed (counters over a
whole week would mix a fresh session with earlier versions of the add-in); "no
content zone" is NOT treated as a problem by its counter - the sign of a real
failure is the window-tree dumps the add-in writes after six consecutive failures
on one window.

## Stage 28 - the native mode by default (2026-07-26, v1.6.0)

The change affects the choice of mode at startup rather than the rendering, but
the matrix requires the full set before rebuilding a package - so the full set was
run.

The new feature test `tests/manual/e2e-default-mode.ps1` (2 Word sessions):

| # | Check | Result | Details |
|---|---|---|---|
| A1 | The add-in starts WITHOUT native-host.cfg | PASS | "Add-in started" in the log |
| A2 | The default is native | PASS | "Tab host mode: Native" |
| A3 | The default is reserve=1 | PASS | "Tab host mode: Native (dump=0, reserve=1)" |
| A4 | The strip is a child window of the Word window itself | PASS | rect=0,178 1920x43 (a WinForms window whose parent is OpusApp) |
| A5 | The strip is NOT at the very top of the window (not over the ribbon) | PASS | top=178 |
| A6 | Not a single ERROR line | PASS | errors=0 |
| B1 | mode=ctp: the add-in starts | PASS | "Add-in started" |
| B2 | mode=ctp: the mode really is the classic one | PASS | "Tab host mode: CustomTaskPane" |
| B3 | mode=ctp: there is NO strip window of our own under the Word window | PASS | the control lives inside the CTP |
| B4 | Not a single ERROR line | PASS | errors=0 |

The full E2E set on a real Word (Word M365 x64, Windows 11 26200, DLL 1.6.0.0):

| Script | Result |
|---|---|
| e2e-default-mode | PASS (10/10) |
| e2e-smoke | PASS |
| e2e-reserve-mode | PASS |
| e2e-side-panes | PASS |
| e2e-overflow-scroll | PASS on the second run (a flake, see below) |
| e2e-tab-colors | PASS |
| e2e-tab-colors-dialog | PASS |
| e2e-tab-size | PASS |
| e2e-overlay-interaction | PASS |
| e2e-overlay-buttons-pv | PASS |
| e2e-settings-pins-order | PASS |
| e2e-dpi-monitors | PASS (7/7; this time 2 monitors were connected - 96 and 120 DPI) |

Automated tests: 53/53 PASS (the default checks were rewritten for native, and
parsing the old `mode=nativeexperimental` plus an explicit `reserve=0` were
added).

Honestly about the flake: on its first run `e2e-overflow-scroll` FAILED the check
"no tab artefacts above the right-hand buttons after scrolling" (bluePx=35 against
a tolerance of 0); on the second it was bluePx=0. Under the flake policy
(docs/TESTING.md) one repeat is allowed; two failures in a row would count as a
bug. The check is pixel-based and sensitive to mouse movement during a run, and
the rendering code did not change in this stage.

A test-harness defect fixed in the same stage: seven E2E scripts restored
`native-host.cfg` only if it HAD existed before the run; if the file was absent,
the mode the script had imposed was left on the machine. Now such a file is
deleted.

Diagnostics 1.3 (in the same stage): a "Tab bar mode" line was added to the
automatic summary. Verified by running the REAL installer/diagnose.ps1:
- on this machine's live logs (the last session was native) -> the mode was
  reported as the main one (tabs inside the Word window), the add-in version
  1.6.0.0 - PASS;
- on the same log truncated at the last "Tab host mode: CustomTaskPane" line
  (with `LOCALAPPDATA` temporarily substituted; the user's folder was untouched)
  -> the mode was reported as CLASSIC plus the path to native-host.cfg - PASS;
- the SHA256 of diagnose.ps1 in installer/ and in the rebuilt package match -
  PASS.

## Stage 29 - the interface language (2026-07-27, v1.7.0)

Automated tests: 57/57 PASS (53 before; +4 for the language). The key new one is
"the captions switch and none of them is empty": it takes EVERY property of
`Strings` by reflection, so a new menu item cannot be left untranslated - the test
sees both an empty string and a string identical to the Russian one.

The new feature E2E `tests/manual/e2e-language.ps1` (2 Word sessions), 8/8 PASS:

| # | Check | Result | Details |
|---|---|---|---|
| A1 | The add-in starts with lang=en | PASS | |
| A2 | The language was taken from language.cfg | PASS | "UI language: En (language.cfg)" |
| A3 | The context menu REALLY is English | PASS | read out of the live Word menu through MN_GETHMENU + GetMenuStringW: "Pin tab \| Close \| Close others \| Tab color \| Tab size \| Open file location" |
| A4 | No Cyrillic is left in the menu | PASS | |
| A5 | Not a single ERROR line | PASS | errors=0 |
| B1 | With lang=auto the language comes from Word itself | PASS | "UI language: Ru (Word UI language (LCID 1049))" - Word on this machine is Russian |
| B2 | The menu matches the detected language | PASS | "Закрепить вкладку \| Закрыть \| …" |
| B3 | Not a single ERROR line | PASS | errors=0 |

The main thing this closes: `LanguageSettings.LanguageID[msoLanguageIDUI]` works
through the NoPIA-embedded interop and returns an LCID on a real Word.

Regression per the matrix: e2e-smoke PASS; e2e-settings-pins-order (the settings
window changed - a "Language" section was added) PASS on the second run.

Honestly about the flake: the first run of e2e-settings-pins-order FAILED the
check "restart: stored order loaded" - a FOURTH document ended up in
`tab-order.cfg` (exactly 3 were expected; the log said
"Tab order loaded: 4 (pinned: 1)"); the repeat run gave exactly 3 and was fully
green. The cause was state leaking between runs in the same batch (a document from
the previous script got into the order file) - the same family of test-harness
defects that was fixed in stage 23. The check was NOT weakened; instead the error
message now prints the "Tab order loaded: N" line itself, so that the next such
case explains itself.

## Stage 30 - English in the installer and Diagnostics (2026-07-27, the DLL was unchanged)

Scripts are not compiled, so the first check is parsing them with the PowerShell
parser (`[Parser]::ParseFile`) - all four: install, uninstall, repair, diagnose.
That is exactly what caught a real error: `uninstall.ps1` and `repair.ps1` had
been overwritten wholesale and lost their BOM - PowerShell 5.1 read the Cyrillic
as ANSI and the file stopped parsing (a known trap of this project,
docs/ENVIRONMENT). The BOM was restored and both parse again.

| # | Check | Result | Details |
|---|---|---|---|
| L1 | The syntax of all four scripts | PASS | after the BOM was restored |
| L2 | `install.ps1 -Lang en` on a real machine | PASS | the whole output is English, the installation succeeded, the registry read-backs are 3/3 |
| L3 | The name in Word's add-in list with -Lang en | PASS | `FriendlyName = Tabs for Word` |
| L4 | `install.ps1 -Lang ru` | PASS | the output is Russian, `FriendlyName = Вкладки для Word` (the machine was returned to its original state) |
| L5 | `repair.ps1 -Lang en` | PASS | "DONE. Start Word and check the tab bar." |
| L6 | `diagnose.ps1 -Lang en` | PASS | the report is entirely English: "Diagnostics script version: 1.4", "Tab bar mode: main (tabs inside the Word window)", "SUMMARY (automatic analysis)" |
| L7 | `diagnose.ps1 -Lang ru` | PASS | the same report in Russian |
| L8 | The report file name is ASCII | PASS | `TabsForWord-diagnostics-<PC>-<timestamp>.zip` (it used to contain Cyrillic) |
| L9 | The package contents and the matching copies of the scripts | PASS | the SHA256 of installer/ and of the package match for all four |

The test ZIP reports were deleted from the desktop after the check.

A deliberate departure from the original plan: the Diagnostics report was NOT made
"always English". It follows the installation language (English for any non-Russian
locale), because a person must be able to read what they are sending; that is
promised outright in the privacy section of INSTALL_RU.md. The strings of the
add-in's own log stay English always - those are what the script parses.

## Stage 31 - English instructions (2026-07-27, the DLL was unchanged)

| # | Check | Result | Details |
|---|---|---|---|
| I1 | Both instructions and both READMEs are in the package | PASS | INSTALL_EN/RU.md, UNINSTALL_EN/RU.md, README.txt, ПРОЧТИ_МЕНЯ.txt |
| I2 | README.txt and ПРОЧТИ_МЕНЯ.txt are UTF-8 with a BOM | PASS | otherwise Notepad and `type` show mojibake |
| I3 | Cross-links RU <-> EN in all four documents | PASS | verified by reading |
| I4 | The file names in the instructions match the real ones | PASS | Diagnostics.cmd, Classic-mode.cmd, In-window-tabs.cmd (after the stage-30 rename) |
| I5 | The repository README is English, the Russian one moved out | PASS | README.md <-> README.ru.md |

The instructions also describe the language choice: for the add-in, from Word's
language and by hand in the settings window; for the scripts, from the Windows
language and through the `-Lang` parameter.

## Stage 32 - translating the code comments (2026-07-27)

About 1470 lines were translated: the comments in every file under `src/` and
`tests/`, the names and messages of the automated tests, the comments of the build
and E2E scripts, and the error texts in the .csproj. Only DATA was deliberately
left in Russian: the Russian half of the `T(en, ru)` pairs in `Strings.cs` and in
the installer scripts, the Russian `ПРОЧТИ_МЕНЯ.txt` in `package.ps1`, and the
wrapper file `Диагностика.cmd`.

| # | Check | Result | Details |
|---|---|---|---|
| C1 | A Release build with no warnings | PASS | 0 warnings, 0 errors |
| C2 | The automated tests after the translation | PASS | 57/57 (the test names are English now) |
| C3 | **The code did not change, only the comments** | PASS | for every `src/*.cs` file the HEAD version and the working copy were compared with the comments stripped out; the only differences are where string literals were translated DELIBERATELY (see below) |
| C4 | The syntax of every .ps1 | PASS | `[Parser]::ParseFile` over installer/, scripts/, tests/manual/ |
| C5 | The BOM of the .ps1 files with Cyrillic was preserved | PASS | build/test/build-release/package |
| C6 | The smoke E2E on a real Word | PASS | the strip, a click, the menu, collapsing, a clean exit |

The deliberate string changes (check C3 found exactly these and nothing else):
- the header comments the add-in writes into its own .cfg files
  (`tab-colors.cfg`, `tab-order.cfg`, `tab-size.cfg`, `hotkey.cfg`,
  `language.cfg`) - the parsers already skipped those lines (they start with `#`),
  so the behaviour does not change;
- the `Reason` strings of the document-area locator - they go into the log, and by
  the decision of ADR-016 the log is always English; `diagnose.ps1` looks for
  different substrings and is unaffected.

## Stage 33 - translating the engineering documents (2026-07-27, the DLL was unchanged)

The public documents under `docs/` were translated into English:
ENVIRONMENT, TESTING, ARCHITECTURE, KNOWN_ISSUES, RESEARCH, NATIVE_TAB_HOST,
DECISIONS and this file. INSTALL_RU.md and UNINSTALL_RU.md stay Russian by design
(their English counterparts are INSTALL_EN.md and UNINSTALL_EN.md), as does
README.ru.md. The author's own working notes stay Russian too and are not part
of the published repository at all.

No code was touched in this stage, so by the docs/TESTING.md matrix nothing had to
be re-run. What was checked instead:

| # | Check | Result | Details |
|---|---|---|---|
| T1 | No Cyrillic is left in the translated documents | PASS | except the deliberate literals: the Russian `FriendlyName`, `ПРОЧТИ_МЕНЯ.txt`, `Диагностика.cmd` and the Russian menu line quoted by the language test |
| T2 | Facts that had gone stale were corrected, not carried over | PASS | ARCHITECTURE (persistence appeared after the MVP), RESEARCH (the strip did end up inside the Word window), NATIVE_TAB_HOST (the default in the class table, the strip width, the dump threshold 3 -> 6, the resize settle 250 -> 125 ms, the ASCII switch names), TESTING (the missing E2E rows) |
| T3 | The Russian originals remain available | PASS | in the commit history, one commit per document |

The reason the documents were translated at all is in ADR-016: the repository is
being prepared to be public, and an engineering document nobody can read is worth
as little as an interface nobody can read.

### The personal-data proofread found a real leak in the DLL

Scanning the shipped binary - not just the text files - turned up the one place
where anonymising the documents would not have helped: the Release DLL embedded
the absolute path of its own .pdb - the full build path, starting from the user
profile folder and therefore including the account name. That string travelled inside
every copy that had already been handed out, so the account name of the machine
the add-in was built on was readable by anyone who opened the file in a text
editor. Fixed in the .csproj: `PathMap` rewrites the build root out of the
metadata, `DebugType=embedded` removes the reference to an external file
altogether, and `Deterministic` makes the output reproducible. As a side effect
the field stack traces now carry line numbers, with nothing extra to ship.

| # | Check | Result | Details |
|---|---|---|---|
| T4 | No absolute path is left in the rebuilt DLL | PASS | a byte scan for `<letter>:\…` in ASCII and UTF-16: nothing; the only hit for "andre" is the substring inside the method name `CollapsedExpandRect` |
| T5 | No separate .pdb next to the DLL any more | PASS | bin\Release contains the DLL only |
| T6 | Release build with no warnings | PASS | 0 warnings, 0 errors; the DLL grew 154 -> 180 KB (the embedded PDB) |
| T7 | Automated tests on the rebuilt DLL | PASS | 57/57 |
| T8 | The smoke E2E on a real Word | PASS | 10/10 - the strip appears, a click switches the document, the menu, collapsing, a clean exit, 0 ERROR |

Honestly about two invalid runs: the first two smoke runs each FAILED the step
"tab click switches document". They do NOT count as "two failures in a row = a
bug" under the flake policy, because the precondition of the E2E scripts was
broken - the mouse was being used during the run (the user said so). The third
run, with the machine left alone, was 10/10. The rule is worth repeating: a run
during which the mouse or keyboard is touched is not a failed run, it is not a
run at all.

### The dress rehearsal: the full E2E set on the clean DLL (2026-07-27, v1.7.0)

Required by the docs/TESTING.md matrix before rebuilding the packages for
distribution. Every script was run in one batch, in the order of the matrix, with
the machine left alone; the driver waits for Word to close between scripts rather
than killing it. Total 8 min 15 s of wall clock.

| Script | Result | Steps | Time |
|---|---|---|---|
| e2e-smoke | PASS | 10/10 | 27 s |
| e2e-default-mode | PASS | 10/10 | 32 s |
| e2e-language | PASS | 8/8 | 33 s |
| e2e-tab-colors | PASS | 15/15 | 50 s |
| e2e-tab-colors-dialog | PASS | 7/7 | 28 s |
| e2e-tab-size | PASS | 8/8 | 42 s |
| e2e-dpi-monitors | PASS | 6/6 | 30 s |
| e2e-side-panes | PASS | 7/7 | 23 s |
| e2e-overflow-scroll | PASS | 7/7 | 46 s |
| e2e-reserve-mode | PASS | 11/11 | 44 s |
| e2e-overlay-interaction | PASS | 9/9 | 33 s |
| e2e-overlay-buttons-pv | PASS | 15/15 | 41 s |
| e2e-settings-pins-order | PASS | 25/25 | 66 s |

13/13 scripts PASS, 0 failed steps, no repeats needed - no flakes in this run.
The only WARN lines in the add-in log are the familiar transient
"Locator: no content zone found" at startup, which the diagnostics classifies as
routine (check D3) and which never turned into a fallback.

**An honest caveat about the DPI check.** `e2e-dpi-monitors` passed, but this time
only ONE monitor was connected, at 100 % - the script said so itself: "all
monitors share one scale - DPI virtualisation was NOT exercised". So the run
confirms the strip reports the monitor DPI and keeps its geometry, but it does NOT
re-confirm multi-DPI behaviour. That evidence comes from the earlier runs: three
monitors at 100/125/150 % (13/13 PASS, the section above) and stage 28 (two
monitors, 96 and 120 DPI). Nothing in this stage touched the DPI code - the change
was `DebugType`/`PathMap` in the .csproj - so re-running multi-DPI was not
required, but the green tick here must not be read as covering it.

| # | Check | Result | Details |
|---|---|---|---|
| T9 | The full E2E set on the rebuilt DLL | PASS | 13/13 scripts, 138 steps, 0 failures |
| T10 | The packages were rebuilt (build-release.ps1) | PASS | build 0 warnings, unit tests 57/57, package + ZIP |
| T11 | The DLL in the package folder has no absolute path | PASS | a byte scan of `<letter>:\…` in ASCII and UTF-16: nothing |
| T12 | The DLL inside TabsForWord.zip has no absolute path | PASS | the same scan of the entry read out of the archive |
| T13 | LICENSE.txt travels in the package | PASS | 19 entries in the ZIP, LICENSE.txt among them |

With this the distribution packages are clean: what gets handed out no longer
carries the account name of the machine it was built on.

## Stage 34 - the File menu is not a locator failure (2026-07-28, v1.7.1)

### Where it came from: the third field report

A Diagnostics report arrived from machine C - and it was a clean one: Windows 10
Pro 22H2 (19045.6456), Word x64 16.0.20228.20110 (consumer M365), 100 % scaling,
Russian Word. All three of those are new ground for the field: the earlier
reports came from Windows 11 25H2 (build 26200), Word 16.0.17932 and 150 %
scaling, so C was the first Windows 10 run of any kind. The install
log is flawless, both registry views are correct, LoadBehavior=3, zero ERROR
lines, no Resiliency entry of ours, nothing in the Windows event log, and the
person said everything worked.

Two things were confirmed by that report before anything was fixed:

| # | Check | Result | Details |
|---|---|---|---|
| F1 | The locator race fix (v1.5.1) holds in the field | PASS | the second window answered "no content zone" once while it was being built, retried and landed correctly at `0,178,1920x37`; the strip never appeared at y=0 (the explicit check in the diagnostics stayed silent) |
| F2 | Side panes are handled in the field (stage 17) | PASS | the first window started with the Navigation pane open and the strip took `350...1920` - the horizontal range of `_WwF`, not the whole window |

The one alarm in the summary came from a window-tree dump, and the dump explained
itself: a visible `FullpageUIHost` covering the whole client area while `_WwF`,
`_WwB`, `_WwG`, every `MsoCommandBarDock` and every `MsoWorkPane` were hidden -
Word with the File menu (Backstage) open, for about 84 seconds. See ADR-017.

### The latent defect this uncovered, and the proof that it was real

Reasoning about the code said that returning "failure" for the File menu would,
with a SINGLE Word window open, reach the manager's threshold of five failed
syncs and switch the add-in to the CTP fallback. Reasoning is not evidence, so
the previous build was run against a live Word - at the time the 1.7.0 DLL was
still in the package folder under `release/` (it has since been replaced by
1.7.2), so the control run cost one start of Word.

| # | Check | Result | Details |
|---|---|---|---|
| C1 | Old build (1.7.0), one document, the File menu open | REPRODUCED | after ~4 s: `[ERROR] Native host failed; falling back to CustomTaskPane: native host unusable for every window 5 syncs in a row`. The strip silently turned into the classic task pane |
| C2 | New build (1.7.1), the same scenario | PASS | no fallback, no dump, one INFO line, the strip returns |

Why nobody had reported it: the counter only advances when EVERY window fails in
the same sync, and machine C had two windows open during its trip to the File
menu. On that machine the poll also froze while Backstage was up (gaps of 27 and
51 seconds between ticks), which hid the problem further. On the development
machine the poll keeps ticking at 1 Hz, so the fallback fires within seconds -
the same add-in, two different symptoms, one cause.

### The fix

| # | Check | Result | Details |
|---|---|---|---|
| U1 | Unit tests on the new DLL | PASS | 58/58 (+1: "locator - the File menu (full-page UI) is not a failure") |
| U2 | A half-built window is NOT excused as the File menu | PASS | asserted explicitly in the new test: the machine A race keeps its counters, its retries and its dump |
| U3 | A hidden or tiny `FullpageUIHost` is not the File menu | PASS | the hidden host and a 100x100 leftover both return false |
| U4 | Release build | PASS | 0 warnings, 0 errors, version 1.7.1.0 |
| E1 | e2e-backstage.ps1 on a real Word | PASS | 12/12 - the strip hides while the full-page UI is up and comes back 0.3 s after it closes, at the top of the document area; no dump, no CTP fallback, one INFO line, 0 ERROR |
| D1 | The real diagnose.ps1 run against the real field logs | PASS | run with `$env:LOCALAPPDATA` pointed at a copy of machine C's logs: the false alarm is gone from the summary, and the report now says "of them 1 were taken while the full-page Word UI (the File menu) was open" |

The E2E deliberately opens exactly one document: that is the case that used to
break. The full-page UI is opened with Ctrl+O, which is language-independent -
the Alt+F key tip only works on an English ribbon, and this Word is Russian.

### The dress rehearsal: the full E2E set on v1.7.1 (2026-07-28)

Required by the docs/TESTING.md matrix before rebuilding the packages. All
fourteen scripts were run in one batch, in the order of the matrix, with the
machine left alone; the driver waits for Word to disappear between scripts rather
than killing it. Total 11 min 13 s of wall clock, 150 steps.

| Script | Result | Steps | Time |
|---|---|---|---|
| e2e-smoke | PASS | 10/10 | 27 s |
| e2e-default-mode | PASS | 10/10 | 33 s |
| e2e-backstage | PASS | 12/12 | 31 s |
| e2e-language | PASS | 8/8 | 33 s |
| e2e-tab-colors | PASS | 15/15 | 49 s |
| e2e-tab-colors-dialog | PASS | 7/7 | 28 s |
| e2e-tab-size | PASS (2nd run) | 8/8 | 43 s |
| e2e-dpi-monitors | PASS | 6/6 | 30 s |
| e2e-side-panes | PASS | 7/7 | 23 s |
| e2e-overflow-scroll | PASS | 7/7 | 48 s |
| e2e-reserve-mode | PASS | 11/11 | 44 s |
| e2e-overlay-interaction | PASS | 9/9 | 33 s |
| e2e-overlay-buttons-pv | PASS | 15/15 | 42 s |
| e2e-settings-pins-order | PASS (2nd run) | 25/25 | 66 s |

14/14 scripts PASS. Two of them needed the one repeat the flake policy allows,
and both are worth writing down, because in both cases the ADD-IN was right and
the SCRIPT was wrong about its own preconditions.

**e2e-tab-size, first run: 3 failed steps.** Every expectation in that script is
computed from the height it measures first, which it takes to be the "Normal"
size. It measured 48 px - that is 37 x 1.3, the "Extra large" multiplier - even
though the script deletes tab-size.cfg before starting Word, and the add-in
defaults to 1.0 with no file. The heights it measured afterwards were exactly
right for what it selected (43 px = 1.15, 37 px = 1.0), so the add-in behaved
correctly throughout and only the derived expectations were off. The batch log
shows every earlier script running at 43 px (the user's own 1.15), so the 1.3
appeared for one session and never again: the repeat, with the identical setup,
started at 37 px. It is recorded here as an unexplained one-off rather than
explained away. What was fixed is the fragile assumption: the script now selects
"Normal" through the menu and confirms it from the log line
("Tab size scale saved: 1") before measuring the baseline. Re-run afterwards:
9/9 PASS, baseline 37 px.

**e2e-settings-pins-order, first run: 1 failed step** - "restart: stored order
loaded" expects `Tab order loaded: 3` and got `5`. The stored order deliberately
keeps documents that are no longer open (that is a tested feature: closing every
tab must not wipe pinning or order), so an exact count is a fragile assertion
about a file that legitimately grows. The next step - the pinned document is
first after the restart - passed, which is the behaviour the check exists for.
The repeat was 25/25.

| # | Check | Result | Details |
|---|---|---|---|
| P1 | The full E2E set on the new DLL | PASS | 14/14 scripts, 150 steps |
| P2 | The user's settings survived the batch | PASS | tab-size.cfg back to the user's 1.15, no native-host.cfg left behind (there was none), hotkey and language untouched |
| P3 | The packages were rebuilt (build-release.ps1) | PASS | 0 warnings, unit tests 58/58, package + ZIP, DLL 1.7.1.0 |
| P4 | No absolute path in the packaged DLL | PASS | a byte scan for `<letter>:\…` in ASCII and UTF-16, both in the package folder and inside the ZIP: nothing |
| P5 | The packaged diagnose.ps1 carries the fix | PASS | version 1.5, the FullpageUIHost rule present, BOM intact, parses |
| P6 | The ZIP is complete | PASS | 19 entries, LICENSE.txt among them |

## Stage 35 - the polish pass before publication (2026-07-28, v1.7.2)

A read-through of the whole source tree before the repository goes public, backed
by scripted sweeps (duplicate lines, unused members, GDI object lifetimes,
untranslated literals, culture-sensitive parsing, debug leftovers, method length).
The codebase came out of it in good shape - no logic bug was found - but two
things did matter.

**A leak that would have been published.** `release/build-log.txt` is a tracked
file, and it carried the absolute build path with the account name of the machine
that produced it - three times. This is the same class of leak stage 33 fixed
INSIDE the DLL with `PathMap`, and it slipped through then because the check
scanned the binary, not the repository. `build-release.ps1` now rewrites the repo
root out of the log before saving it (and writes it as UTF-8 with a BOM, so the
Russian MSBuild output is readable rather than mojibake).

**A real micro-bug in the keyboard hook.** `KeyboardHookService` acted on any
hook code `>= 0`, which includes `HC_NOREMOVE` - the code Windows uses when a
message is being PEEKED at and will be delivered again. On a peek the tab would
switch twice for one Ctrl+Tab. It never showed on this Word (the E2E has always
measured a single step), but it depended on how the host pumps messages, so it
was luck rather than design. Now only `HC_ACTION` is handled.

| # | Check | Result | Details |
|---|---|---|---|
| S1 | No personal data in tracked files | PASS | after the build-log fix: no absolute paths, no machine names, no account names. The three remaining hits for "andre" are a Microsoft blog URL in RESEARCH, the sentence in this file describing them, and the substring inside the method name `CollapsedExpandRect` |
| S2 | Duplicate comment fragments left by the stage 32 translation | FIXED | seven of them (four exact repeats, three where a line repeated the tail of the one above); a second detector was written for the partial case |
| S3 | Dead code | FIXED | `GetParent`, `SWP_HIDEWINDOW`, `SW_SHOWNA` and the whole unused `SetThreadDpiHostingBehavior` block removed from NativeWin32 |
| S4 | Duplicated helper | FIXED | three byte-identical `RoundedRect` implementations (strip, all-tabs menu, settings window) merged into `Glyphs.RoundedRect` |
| S5 | Stale comments from before v1.6.0 | FIXED | the strip described itself as living in a Custom Task Pane; the reserve section header still called itself off-by-default; "MVP" as a name for the first version |
| S6 | `UiLanguage.Init` state | FIXED | the "explicit language" flag was never cleared on re-detection - harmless today because the only caller clears it first, but it made the invariant depend on the caller |
| S7 | GDI/GDI+ object lifetimes | PASS | every brush, pen, font, bitmap and path is either in a `using` or disposed in `Dispose`; `GetDC` is released in a `finally` |
| S8 | Culture-sensitive parsing / debug leftovers / untranslated literals | PASS | no `ToLower()` without invariant, no `Parse` without `TryParse`, no `MessageBox`/`Console`/`#if DEBUG`, every visible caption goes through `Strings` |
| S9 | Unit tests, smoke and the Ctrl+Tab E2E on the polished DLL | PASS | 58/58; e2e-smoke 10/10, e2e-overlay-interaction 9/9, e2e-settings-pins-order 25/25 (the last two are the ones that exercise Ctrl+Tab) |
| S10 | Packages rebuilt at 1.7.2 | PASS | 0 warnings, no absolute paths in the DLL in the folder or inside the ZIP, 19 entries |

### The rename: WordTabsMvp -> TabsForWord (2026-07-28)

Everyone who had been given the add-in had already uninstalled it, so the project
was free to drop "Mvp" from its internal identity before the first public
release: the assembly, the namespace, the ProgId, the registry key, the settings
folder, the log and diagnostics file names. What a user sees in Word did not
change - the add-in has been listed as "Tabs for Word" since 1.6.0.

The replacement was done at the BYTE level rather than through text APIs, for one
reason: the repository holds .ps1 files that must keep their UTF-8 BOM and .cmd
files that must not have one, and rewriting those through PowerShell is precisely
how that gets destroyed (it has happened twice in this project). Both names are
ASCII and, as it turns out, exactly eleven bytes long, so the substitution could
not disturb an encoding, a BOM or a line ending. 446 replacements across 77 files;
the paths (`src/`, `tests/`, the solution and both project files) were moved with
`git mv`.

| # | Check | Result | Details |
|---|---|---|---|
| R1 | No trace of the old name is left | PASS | zero occurrences in tracked files and zero in tracked paths |
| R2 | Release build under the new name | PASS | 0 warnings; the assembly identifies itself as `TabsForWord, Version=1.7.2.0` |
| R3 | Unit tests | PASS | 58/58 |
| R4 | The old registration was removed first | PASS | uninstall.ps1 cleared the add-in key and both registry views; otherwise Word would have loaded the old and the new add-in side by side and every E2E would have been meaningless |
| R5 | Fresh install of the renamed package | PASS | both registry views written and read back: `Class=TabsForWord.Connect`, `Assembly=TabsForWord, Version=1.7.2.0`, CodeBase pointing at `%LOCALAPPDATA%\TabsForWord\TabsForWord.dll`, LoadBehavior=3 |
| R6 | The package rebuilt with the new names | PASS | `release/TabsForWord-Installer/` + `TabsForWord.zip`, 19 entries |
| R7 | A stale name in a user instruction | FIXED | INSTALL_RU told the reader to look for "Вкладки для Word (MVP)" in Word's COM add-ins list, but the "(MVP)" was dropped from that name back in 1.6.0. The English instruction was already correct - only the Russian one had been left behind |
| R8 | The full E2E set on the renamed build | PASS | 14/14 scripts, 150 steps, 9 min 35 s, **no repeats at all** - the two scripts that had each needed a second run in the previous rehearsal passed first time (the tab-size baseline is now forced rather than assumed) |

A rename that touches COM registration is the one change that can make the add-in
silently not load, which is why it was verified by the whole matrix rather than
the smoke test: if the ProgId, the CLSID entries, the assembly name and the
CodeBase had drifted apart, nothing would have started at all.

**Deliberately not changed.** Five methods are longer than 80 lines
(`OnPaint` 118, `ComputeLayout` 110, `BuildSnapshot` 107, `TabListPopup` ctor 102,
`UpdateLayout` 97). Each is one cohesive job and each is commented; splitting them
the day before publication would mean a full E2E round for a readability gain, so
they stay as they are. The two "#RRGGBB" parsers (`TabColorStore.ParseRgb` and
`TabHostSettings.ParseColor`) also stay separate: merging them would tie the
window-host settings to the tab colour store for eight lines of arithmetic.

## Stage 37 (v1.7.3, 2026-09-26) - tab switching, the ruler after a move, AutoSave

Environment for everything measured here: the development machine - Windows 11
Pro 25H2 (build 26200), Word M365 x64 16.0.20326, ONE monitor 1920x1080 at
100 %, tab-size multiplier 1.3 (a 48 px strip). Clicks and drags were real mouse
input driven by a script, on temporary documents only. Check IDs restart in every
section (R8, for example, exists both in stage 23 and in the rename subsection),
so cite a check by section and ID.

### Tab switching: a click that did not switch

The report: now and then a click on a tab presses the tab but the window does not
change, perhaps because some windows are maximised and some are not. The user's
own add-in log of 2026-09-17 showed the same tab requested three times within
7 s, and the same pattern twice more; each time the PREVIOUS window's document
area was re-laid out about 1.3-2 s after the click ("Reserve: anchor moved"). The
old log line "Window activated" was written unconditionally, so the log could
not show whether a switch stuck. The Windows Application log of that day also
has a Word hang (AppHang, event 1002) and GPU driver timeouts (LiveKernelEvent
141) - system-level issues the add-in does not cause, but they can make a switch
look as if it did not happen.

The reproduction on v1.7.2 (4 temporary documents, maximised and normal windows
mixed):

| # | Check | Result | Details |
|---|---|---|---|
| W1 | Clicks on the strip of the FOREGROUND window, mixed window states | PASS | 40/40 switched |
| W2 | The same with Word's full-screen mode (View.FullScreen) | PASS | 30/30; Word's full-screen mode applies to ALL Word windows at once |
| W3 | Clicks on the strip of a BACKGROUND (visible, non-active) normal window, 28 px inside a tab | PASS | 30/30 |
| W4 | The same, 6 px inside the tab's left edge | **FAIL (reproduced)** | 3/30 missed: the window whose strip was clicked came to the front instead of the requested one |

W4 only exists with non-maximised windows - with every window maximised only the
foreground window's strip is visible - which matches the "windowed vs full
screen" remark. The root cause, confirmed by the reproduction and by the code:
the press itself activated the background window (WM_MOUSEACTIVATE went up to
OpusApp, and `OnMouseDown` gave the strip the focus, which also activates an
inactive top-level window). Word's WindowActivate made that window's tab the
active one; the active tab is wider (semibold caption, always-visible close box),
so the tabs shifted under the cursor BEFORE the click was resolved, and the click
landed on the neighbouring tab (often the clicked window's own) or on empty
space.

Other defects found by analysis and fixed together: restoring a minimised window
through COM `WindowState = Normal` turned a maximised window into an ordinary one
(which itself creates the mix of states); a double click on a tab performed two
switches; a "drag" that ended where it started swallowed the click, and a lost
button-up could turn plain hovering into a phantom drag; middle-click (close) on
a background window's tab could close the NEIGHBOURING tab, and the context menu
could open for the neighbour; the Ctrl+Tab hook also fired inside Word's own
dialogs that have tabs (Font, Paragraph).

What v1.7.3 does: in the in-window mode a left or middle press on a tab answers
WM_MOUSEACTIVATE (and WM_POINTERACTIVATE for touch/pen) with "no activate" and no
longer takes the focus; the hit is taken on the layout the user saw; the switch is
posted to run after the click is over and double-click presses are de-duplicated;
minimised windows are restored with SC_RESTORE; the result of every switch is
logged ("Window activated: hwnd=N (tab, sfw=1, in front, from ... to ..., 31 ms)");
a check at +250 ms and +1200 ms writes a WARN "Activate check: target ... is not
in front ..." and re-raises the target ONCE, only if the foreground went back to
the window the switch started from and there was no input since the click. The
classic Custom Task Pane mode keeps the old press behaviour.

The results on v1.7.3 (the final build):

| # | Check | Result | Details |
|---|---|---|---|
| W5 | Background-strip clicks 6 px from the tab's left edge (the W4 case) | PASS | 0/40 missed (was 3/30) |
| W6 | Foreground-strip clicks, mixed window states | PASS | 0/30 missed |
| W7 | The add-in log for those 70 switches | PASS | no "Activate check" warning, no ERROR |
| W8 | Word itself undoing a switch (the 2026-09-17 pattern) | NOT TESTED | not reproduced, so not proven either way; if it happens, 1.7.3 logs it and recovers once - field logs of 1.7.3 will tell |
| W9 | Touch/pen press on a tab (WM_POINTERACTIVATE) | NOT TESTED | no touch or pen device on this machine |
| W10 | Middle-click close and the context menu on a BACKGROUND window's tab, live | NOT TESTED | the rule is covered by unit tests (the WM_MOUSEACTIVATE decision incl. the middle button, middle click on the same tab) |
| W11 | Double click on a tab = one switch; restoring a minimised maximised window keeps it maximised; Ctrl+Tab ignored inside the Font/Paragraph dialogs | NOT TESTED | fixed in the code, no live check was run for these |
| W12 | The classic Custom Task Pane mode (press behaviour unchanged) and switching between the modes | NOT TESTED | this stage ran only the in-window mode |
| W13 | Switching across several monitors | NOT TESTED | one monitor in this stage |

An intermediate build had 1 miss in 22, traced to the test harness itself;
another intermediate run was 0/40. Only the final-build numbers are in W5-W7.

### The ruler after moving the window

The request: halve the time it takes the ruler to reappear after the window is
moved. The mechanism: in reserve mode, after a SIZE change Word re-lays out its
window and puts its document container (`_WwF`) back at the natural top, so the
strip covers the ruler until the add-in shifts the anchor again. v1.7.2 waited
for 125 ms of quiet (SizeSettleMs) and then did nothing until the next WinEvent
or the ~1 Hz poll - after a snap or a maximise nothing else happens, so the ruler
waited up to ~1.1 s.

Measured with a real mouse drag of the title bar, sampling the geometry every
~17 ms (n = repetitions):

| Gesture | v1.7.2 | v1.7.3 |
|---|---|---|
| Move a normal window (no size change) | the ruler is never covered | the ruler is never covered |
| Drag a MAXIMISED window by its caption (it restores mid-drag) | covered ~150 ms at the restore, back before release | the same |
| Drag to the top edge and release (Aero Snap maximise) | anchor fixed 219-844 ms after release, median **751 ms** (n=6) | median **207 ms** (n=8, dev build), **237 ms** (n=6, final build); ruler visibly painted: median 269 ms |
| Drag to the left edge (snap to the left half) | not measured | anchor median 221 ms, ruler visibly painted 284 ms (n=8) |

The add-in log of the final run: "anchor moved ... N ms after the resize began",
values 125-250 ms (median ~156).

What v1.7.3 does: a one-shot timer per host fires exactly when the quiet period
ends instead of waiting for the poll; while the user still drags (Word's thread
in a move/size loop, or a mouse button held) it only re-checks every 50 ms and
does not move the anchor; a re-entrancy guard stops a nested layout pass from
shifting the anchor twice; every fit that answers a resize is logged with its
delay. SizeSettleMs stays 125 ms: Word's own follow-up layout pass was observed
~80 ms after a size change, and a shorter quiet time would move the anchor before
it and cause a second fit (visible flicker).

| # | Check | Result | Details |
|---|---|---|---|
| Y1 | Aero Snap maximise: time until the anchor is fitted | PASS | median 751 -> 237 ms on the final build - more than 3x faster (the request was 2x) |
| Y2 | Moving a normal window never covers the ruler | PASS | both versions |
| Y3 | A live resize by the window edge does not make the page jump on every pause | NOT TESTED | by design (no anchor move while a button is held); not measured separately |
| Y4 | The TickCount wrap bug found by the review | FIXED, live NOT TESTED | `Environment.TickCount` is negative for half of every 49.7 days of uptime; tick fields that started at 0 broke the quiet-period check (no fit until the first resize), the oscillation window of the reserve guard (ordinary fits added up over hours until the guard suspended the reservation and, after three times, gave it up for the session) and the background sampling. All tick fields now start "long ago" and the quiet period has its own flag. A live check would need 24.9+ days of uptime |
| Y5 | e2e-reserve-mode.ps1 could not fail its "no fight loop" check | FIXED | it counted the log line "Reserve: anchor shifted", which the add-in no longer writes (it writes "Reserve: anchor moved"); fixed, plus new steps "anchor fits are logged" and "no oscillation suspension". Its result is in the full E2E set below |

### AutoSave: a save on every switch (not an add-in bug)

The report: with AutoSave on, every edit is saved as soon as the user switches to
another document; the unsaved dot disappears and switching away from such a tab
lags. This is the AutoSave switch in Word's title bar, for OneDrive/SharePoint
files. The add-in never calls Save/SaveAs and never cancels a save
(DocumentBeforeSave only schedules a reconcile).

The experiment: two temporary test documents in the user's OneDrive (with the
user's permission; deleted afterwards), opened by Word as cloud documents with
AutoSave on. Each repetition: type 3 characters, wait 250 ms, switch, poll
`Document.Saved`.

| Method | Saved after the switch |
|---|---|
| No switch (baseline) | 1 of 4 repetitions saved after 7.4 s; 3 of 4 not saved within 12 s |
| Tab click (the add-in, v1.7.2) | 219-312 ms (4/4) |
| Alt+Tab | 218-281 ms (4/4); the target window comes to the front only after the save |
| SetForegroundWindow from outside Word | 0-15 ms (3/3) |
| Word's own Window.Activate (COM from outside) | 125 ms (3/3) |
| Alt+Tab with the add-in DISCONNECTED (COMAddIns...Connect = False) | 218-234 ms (3/3) |

The conclusion: Word's AutoSave saves a changed cloud document as soon as its
window loses activation, synchronously, whatever does the switching - with or
without the add-in; the next window comes forward only after that save, which is
the lag the user notices. The add-in will not turn AutoSave off behind the user's
back (the rule: never change how the user's documents are saved). What the user
can do is written in docs/INSTALL_EN.md and docs/INSTALL_RU.md ("Word saves the
document every time I switch tabs").

| # | Check | Result | Details |
|---|---|---|---|
| AS1 | The add-in does not save documents | PASS | by reading the code, and the Alt+Tab row with the add-in disconnected |
| AS2 | The same experiment on v1.7.3 | NOT TESTED | the switching change does not touch saving; the cloud test documents were deleted after the v1.7.2 run |

### Verification of the change

| # | Check | Result | Details |
|---|---|---|---|
| V1 | Unit tests | PASS | 68/68 (58 before; +10: the retry rule, input-since with tick wrap, Describe never reads the window title, the WM_MOUSEACTIVATE decision incl. the middle button, a press keeps the focus and empty space gives it, a jitter drag is a click, a lost button-up, a real drag reorders, a snapped-back drag is not a click, middle click on the same tab) |
| V2 | An independent code review of the change (4 reviewers) | 13 FIXED, 7 rejected | 20 findings, 13 confirmed after verification and all fixed (the watch clock taken from the request, middle/right click on background tabs, tick wrap, the focus route, the snapped-back drag, the held-button gate for the timer, the accuracy of the settle log, touch activation, the E2E grep, temporary diagnostic logging removed); 7 rejected with reasons |
| V3 | install.ps1 and Diagnostics 1.6 label build 22000 and later as Windows 11 | PASS (parsing) | both scripts parse without errors (`[Parser]::ParseFile` in PowerShell 5.1). Up to 1.5 they printed the registry ProductName, which says "Windows 10" on Windows 11 |
| V4 | Diagnostics 1.6 classifies "Activate check: target" as a serious warning | PASS (by inspection) | the pattern is in the serious list; no real log with that warning exists yet |
| V5 | A live run of the changed install.ps1 and of diagnose.ps1 1.6 | PASS | the package release/TabsForWord-Installer was installed with its install.ps1 (all registration checks OK, DLL 1.7.3.0) and diagnose.ps1 was run: both print "Windows 11 Pro 25H2 (build 26200.…)" on the development machine, where 1.5 printed "Windows 10 Pro". The summary listed 6 ERROR lines - all of them "Test error marker", written into the real add-in log by the unit test of LoggingService (a pre-existing nit of the test suite, left for a later fix: the test should write to a temporary log) |

### Full E2E set

Run 2026-09-26 16:04-16:15 on the development machine (one monitor, 100 %), the v1.7.3 build installed by
every script itself, one script at a time, nobody touching the mouse or keyboard. Flake policy: one repeat of a
failed script - not needed.

| ID | Script | Result | Steps |
|---|---|---|---|
| W1 | e2e-smoke.ps1 | PASS | 10/10 |
| W2 | e2e-default-mode.ps1 | PASS | 10/10 |
| W3 | e2e-backstage.ps1 | PASS | 12/12 |
| W4 | e2e-language.ps1 | PASS | 8/8 |
| W5 | e2e-tab-colors.ps1 | PASS | 15/15 |
| W6 | e2e-tab-colors-dialog.ps1 | PASS | 7/7 |
| W7 | e2e-tab-size.ps1 | PASS | 9/9 |
| W8 | e2e-dpi-monitors.ps1 | PASS (single monitor) | 6/6 - the script itself notes that all monitors share one scale, so DPI virtualisation was NOT exercised |
| W9 | e2e-side-panes.ps1 | PASS | 7/7 |
| W10 | e2e-overflow-scroll.ps1 | PASS | 7/7 |
| W11 | e2e-reserve-mode.ps1 | PASS | 13/13 - with the corrected log pattern it now really counts the fits: 2 after the idle phase, 6 in total (bound 25); no suspension, no give-up |
| W12 | e2e-overlay-interaction.ps1 | PASS | 9/9 (includes Ctrl+Tab) |
| W13 | e2e-overlay-buttons-pv.ps1 | PASS | 15/15 (includes Protected View) |
| W14 | e2e-settings-pins-order.ps1 | PASS | 25/25 (includes Ctrl+Tab and dragging tabs) |

14/14 scripts, 153 steps, 0 failures, 0 repeats, 11 minutes.

After this run one more defensive change went in (if posting the switch with BeginInvoke ever throws, the
"switch queued" flag is cleared and the switch runs inline, instead of every later switch being dropped). On that
final build: unit tests 68/68, e2e-smoke 10/10, e2e-overlay-interaction 9/9, e2e-settings-pins-order 25/25.
