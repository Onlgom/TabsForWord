# KNOWN_ISSUES.md

A record of the serious problems that were hit, and of the limitations that are
accepted on purpose. Nothing here is hidden: a limitation written down honestly
is worth more than a surprise later.

## Problems that were fixed

### 1. NullReferenceException in BuildSnapshot on the Protected View -> editing transition
- **Symptom:** pressing "Enable Editing" produced
  `ERROR BuildSnapshot failed (WindowActivate/DocumentChange)` in the log.
- **Cause:** during the transition `Application.ActiveProtectedViewWindow`
  returns null (not a COMException), and the guard only covered COMException.
- **Status:** RESOLVED (2026-07-23). A null guard plus wider catches in the PV
  section. Word never crashed even before the fix (the outer catch held) and the
  state recovered by itself.

### 2. An empty strip: the CTP title ate the whole height
- **Symptom:** with a pane height of 44 px the control was left about 9 px - no
  tabs were visible.
- **Cause:** Word's own pane title ("Documents" plus its buttons) counts towards
  the CTP height; measuring right after Visible returned Height=0 (the host had
  not laid the control out yet).
- **Status:** RESOLVED (2026-07-23). A one-shot height correction on the first
  SizeChanged carrying a real size. Result: a 67 px pane with 29 px of content.

### 3. A click on a tab in a background window sometimes switched to the wrong window
- **Symptom:** a tab was "pressed" but the window did not change. A field log
  (2026-09-17) shows the same tab requested three times in 7 seconds. Reproduced
  on v1.7.2 with real mouse clicks: on the strip of a visible but NOT active
  ordinary window, 6 px inside a tab's left edge, **3 clicks in 30 missed** - the
  window whose strip was clicked came to the front instead of the requested one
  (28 px inside the tab: 0/30; on the strip of the window in front: 0/40, and
  0/30 with Word's full-screen mode). It only exists with non-maximised windows:
  with every window maximised only the strip of the window in front is visible -
  hence the user's impression that it had to do with "windowed vs full screen".
- **Cause:** the press itself activated the background window
  (WM_MOUSEACTIVATE went up to OpusApp, and `UserControl.OnMouseDown` put the
  keyboard focus on the strip, which also activates an inactive top-level window).
  Word's WindowActivate -> Reconcile made that window's own tab the active one,
  which is wider (semibold text, an always-visible close box), so the tabs shifted
  under the cursor BEFORE the click was resolved and it landed on a neighbouring
  tab or on empty space. The same shift could make a middle click close the
  NEIGHBOURING tab and a right click open the menu of the neighbour.
- **Status:** RESOLVED (2026-09-26, v1.7.3). In the in-window mode a left or
  middle press on a tab body answers WM_MOUSEACTIVATE (WM_POINTERACTIVATE for touch
  and pen) with "no activate" and does not take the keyboard focus; the hit of a
  press is taken on the layout the user saw; the switch is posted to run after the
  click is over, logged with its result and checked afterwards (ADR-018). Result
  on real Word: **0 misses in 40** clicks 6 px from the edge of a background strip,
  0/30 on the strip in front with mixed window states, and no "Activate check"
  warning in the log. Fixed together: a minimised window is restored with
  SC_RESTORE, so a window that was maximised comes back maximised (the old COM
  `WindowState = Normal` turned it into an ordinary one, creating the very mix of
  maximised and ordinary windows); a double click is one switch, not two; a "drag"
  released within three drag thresholds of the press counts as a click; a lost
  button-up no longer turns plain hovering into a phantom drag that swallows the
  next click; a middle click closes only a tab that was both pressed and released
  on; the context menu is for the pressed tab.

### 4. After about 25 days of uptime the reserve could give up for the session
- **Symptom:** none reported; found by a code review in stage 37. The possible
  effect: the ruler permanently under the strip until Word restarts.
- **Cause:** `Environment.TickCount` is negative for half of every 49.7 days of
  uptime (from 24.9 to 49.7 days). Tick fields that started at 0 then broke the
  interval checks of the in-window host: (a) the quiet-period check - the reserve
  never fitted until the first resize; (b) the oscillation window of the reserve
  guard never reset, so ordinary fits added up over hours until the guard
  suspended the reservation for 30 s and, after three times, gave it up for the
  session; (c) the background colour sampling never ran.
- **Status:** RESOLVED (2026-09-26, v1.7.3). Every tick field starts "long ago",
  and the quiet period has its own flag instead of being judged by the tick alone.

## Known limitations (architectural, documented honestly)

### O1. Tabs do NOT merge the windows physically
Every document stays in its own Windows window (SDI - MDI was removed from Word
2013 onwards and there is no official API for merging; see docs/RESEARCH.md).
The tab strip is shown in every window in sync, and a click switches windows.
Alt+Tab still shows every Word window - that is Windows behaviour, not a defect.

### O2. The CTP title strip (gone in the main mode, v1.6.0; still there in the classic one)
A limitation of the official Office API: Word reserves space for its own Custom
Task Pane title (about 35 px) and enforces a minimum pane height - the strip
cannot be removed through Office. That is exactly why the in-window host was
built: since v1.6.0 the strip lives directly inside the Word window (ADR-015),
there is no title strip, the bar takes exactly its 37 px (times DPI times the
user multiplier), and the collapse button reduces it to 16 px.

The limitation applies ONLY in the classic mode (`mode=ctp`, the file
installer\Classic-mode.cmd) and after an automatic fallback to CTP: there the
strip still takes about 72 px in total, the title is left without any text, and
collapsing reduces it to whatever minimum Word allows rather than to nothing.

### O3. There is no strip inside a Protected View window
Add-ins are not hosted in PV windows (an Office security limitation). A document
in PV appears as a grey tab in the strips of ORDINARY windows; after "Enable
Editing" the tab becomes an ordinary one. If ONLY a PV window is open (no
ordinary window at all) there is nowhere to show the strip - the tabs appear
once any ordinary document is opened.

### O4. The user can hide the strip with its own close button (classic mode only)
The close button on the "Documents" title hides the strip of that window until
the end of the session (a CTP mechanism). It comes back when Word restarts or in
new windows. There is no "show it again" button in the MVP. In the main
(in-window) mode there is no CTP title, so there is nothing to hide it with.

### O5. Only the current WINWORD.EXE process
The tabs manage the windows of their own Word instance only. Separate WINWORD
processes (rare; for example Word started "as administrator" alongside an
ordinary one) get independent strips. Cross-process merging is deliberately not
implemented (fragile automation).

### O6. Word started "as administrator"
If Word runs elevated, the per-user (HKCU) COM registration may be ignored and
the strip will not appear. Start Word the ordinary way.

### O7. Overflow
With many windows the tabs first shrink evenly (the active one never below
180 px, the inactive ones down to 80 px in the compact mode), then the scroll
arrows appear. Drag-and-drop reordering was implemented in stage 12
(specification 2b).

### O8. Ctrl+Tab shadows Word's own behaviour in tables (partly lifted, stage 22)
Per specification 2b, Ctrl+Tab / Ctrl+Shift+Tab switch tabs (a local WH_KEYBOARD
hook on the Word UI thread). Word itself uses Ctrl+Tab to insert a tab character
inside a table cell - that behaviour is shadowed while the interception is on.
Since stage 22 the interception can be switched off in the settings window
(the all-tabs menu -> "Settings…" -> "Switch tabs with Ctrl+Tab"); it applies
instantly, is stored in hotkey.cfg and survives a restart. It is on by default
(the MVP behaviour is unchanged). No other shortcuts are affected, and the hook
is removed when the add-in unloads.
Since v1.7.3 the interception is active only while an enabled Word document
window of this Word process is in front. Inside Word's own dialogs that have
tabs of their own (Font, Paragraph) Ctrl+Tab reaches the dialog as usual;
before, the hook took it there too.

### O9. Tab order after drag-and-drop used to live only in the session (lifted, stage 22)
The order lives in the add-in model (the same in every window and in the all-tabs
menu) and is now stored in TabOrderStore (tab-order.cfg): it is restored once, on
the first reconcile after Word starts, and then lives in memory as before.
It only applies to saved documents (the key is the file path); the order of
unsaved documents is not restored between runs - the same as the tab colour.

### O13. Pinning: unpinning does not restore the historical position
"Pin tab" always moves the tab to the front of the pinned group. "Unpin" does
NOT teleport it back to where it was before pinning - it stays where the last
stable sort put it (usually at the border between the groups), the same way
Chrome and Edge behave. A deliberate simplification: no separate memory of
"where the tab used to be" is kept.

### O10. The key of a Protected View tab is a 32-bit hash of the path
PV windows have no HWND, so the key is a negative hash of the full source path.
A collision between two different paths is theoretically possible, and two PV
windows of the SAME file share one key by construction (one tab for both). These
are exotic scenarios; the consequence is that activating or closing hits the
first matching window. Accepted as a limitation (external audit 2, item 11).

### O11. An extremely narrow window: the scroll viewport is never narrower than one tab
In overflow mode the width of the tab zone is forced to be at least the minimum
tab width, even when there is less real space between the arrows and the "+"
button - with a window a couple of hundred pixels wide, the drawing and the
clicks can overlap the buttons. Real Word barely allows such a window. Accepted
as a limitation (external audit 2, item 12).

### O12. Keyboard focus does not auto-scroll and does not follow a reorder
The arrows / Home / End move the focus frame by index: in scrolling mode the
frame can end up outside the visible zone, and after tabs are dragged the index
is not recomputed by hwnd. The mouse and the all-tabs menu are unaffected.
Accessibility polish is out of scope for the MVP (external audit 2, item 13).
How the strip gets the keyboard focus (in-window mode, since v1.7.3): a press on
a tab no longer takes the focus - after a switch the focus goes back to the
document of the activated window. The strip takes the focus from a press on its
empty space or on a scroll arrow while its window is in front, so the keyboard
route (arrows, Home/End, Enter) stays reachable. The classic CTP mode keeps the
old behaviour (any press focuses the strip).

### O14. Soft tab text on monitors scaled differently from the system scale
On a monitor whose scale differs from the SYSTEM one (the scale the primary
monitor had at sign-in), the tab text is slightly soft or blurry.
Cause: the strip is a WinForms surface, which is always system-aware, so it is
drawn at the system scale and Windows stretches that image on monitors with a
different scale (see ADR-010 and ADR-013). It shows up ONLY in multi-monitor
setups with different scales, when the tabs are on a monitor whose scale is not
the system one; on a single monitor (any primary one) the text is crisp.
The magnitude is small - a letter edge widens from about 1 to about 2 pixels
(borderline noticeable; measured 2026-07-25 at 100 / 125 / 150 %). The SIZES of
the strip are correct on every monitor, and the monitor a document was CREATED on
does not affect sharpness (the strip repaints for the current monitor). The real
fix (our own per-monitor-aware window instead of WinForms) was deliberately
deferred until after the MVP - the reasoning and the plan are in ADR-013.

### O15. ARM64 Windows / ARM64 Office - not tested
The add-in is built as AnyCPU (net48) and has only ever been run on x64.
On ARM64 Windows with the ordinary (emulated) x64 Office it will most likely
load: registration does not depend on bitness (an Addins key without a version
number plus the CLSID in both registry views), and the CLR bitness matches the
Word process through the emulation layer. NATIVE ARM64 Office would require
.NET Framework 4.8.1 (the first version with native ARM64 support); the installer
recognises ARM64 only for the diagnostics (it writes it to the log) and has no
separate ARM64 branch or 4.8.1 check. Bottom line: never tested on any ARM
device and not officially supported. To be checked when a real ARM64 device
becomes available.

### O16. Windows 10 and Word editions other than Microsoft 365 - field-confirmed in part, not lab-tested
Every PASS of the test suite (unit tests and E2E) was produced on the development
machine: Windows 11 (build 26200) with Microsoft 365 Word x64. What else is known
comes from field reports (Diagnostics), not from test runs:
- **Windows 10 22H2** (build 19045): field-confirmed on ONE machine (machine C,
  2026-07-28, Microsoft 365 Word x64, 100 % scaling) - a clean install, zero ERROR
  lines, the user confirmed it works, checks F1/F2 in TEST_RESULTS. The E2E suite
  has never been run on Windows 10.
- **Word ProPlus 2024 x64** (16.0.17932): field-confirmed on machines A and B
  (Windows 11 25H2, build 26200, 150 %). Their reports said "Windows 10 Pro 25H2":
  that was the registry ProductName, which on Windows 11 still reads "Windows 10";
  build 26200 is Windows 11 (the installer and Diagnostics 1.6 now label builds
  from 22000 on as Windows 11).
- **Word 2016, 2019 and 2021, and 32-bit Word**: NOT TESTED anywhere. They are
  claimed by construction only (the add-in is built against the Word 15.0
  interop, ADR-005, and the CLSID is registered in both registry views for
  32-bit Word, ADR-009).
On Windows 10 builds before 1903, or on a system without .NET 4.8, the COM
component will not activate; since the installer update such a case stops the
installation with an explicit error (see the .NET check in
installer/install.ps1) instead of reporting a false success.

### O17. The main mode relies on undocumented internals of the Word window
The strip is embedded into the Word window through Win32: the anchor is found by
the window classes `_WwF` (the container under the ribbon), `_WwG` (the document
canvas) and `_WwB`. These names have been stable for decades, but Microsoft does
not document them and is not obliged to keep them. If an Office update changes
the hierarchy, the fallback heuristic of the locator kicks in, and on a complete
failure the add-in falls back to the classic CTP mode automatically (5 failed
synchronisations in a row, an ERROR line "falling back to CustomTaskPane" in the
log, and a separate item in the diagnostics report).
The risk is not removed but insured against; accepted deliberately in ADR-015.
The manual way back to the classic look is installer\Classic-mode.cmd.

### O18. The Office dark theme: the tabs stay light
The strip background is matched to the Word work area (bg=auto), so in the dark
theme it becomes dark - but the tabs themselves are drawn with the
"specification 2b" palette, which is light. A dark tab palette was never
designed; that is separate design work, out of scope for the move to the
in-window mode. Workarounds: bg=spec (a fixed light background) in
native-host.cfg, or the classic mode.

### O19. Background matching does not work in every view
bg=auto samples the colour of the Word work area from the screen (Word draws its
canvas with Direct2D, so the colour cannot be read from its DC). In Print Layout
the match is exact (delta <= 3). In Web Layout and Draft the work area is white,
so the strip keeps the specification colour. bg=#RRGGBB settles it by hand.

### O20. The strip disappears while the File menu is open, and returns about a second and a half later
The File menu (Backstage) and the pages opened from it - Print, Save As, Account -
replace the whole Word interface with a full-page screen and hide the document
area completely. There is nothing to pin the strip to at that moment, so it hides
too; leaving it visible would mean a row of tabs floating over the File screen.
On the way back Word rebuilds its layout, and the strip reappears within about a
second and a half (measured in the field: 1.4 s). Nothing is lost - the tabs, their
order, colours and pinning are all kept.
The return depends on the add-in's ~1 Hz poll. On machine C Word barely let it
tick while the full-page screen was up (gaps of 27 and 51 seconds inside a single
visit to the File menu); on the development machine it kept ticking at 1 Hz
(stage 34, TEST_RESULTS). Either way this does not delay anything, because the
ticks resume at the moment the menu closes - which is exactly when the strip is
needed again.

### O21. AutoSave (OneDrive/SharePoint) saves a document the moment you leave its window
For a document stored in OneDrive or SharePoint with Word's "AutoSave" toggle on
(the switch in the Word title bar), Word saves the changes as soon as the
document's window loses activation - synchronously, and whatever does the
switching. What the user sees: the unsaved dot disappears when they switch away,
and the next window comes forward only after that save, so the switch lags.
This is Word's behaviour, not the add-in's: the add-in never calls Save/SaveAs
and never cancels a save (DocumentBeforeSave only schedules a reconcile).
Measured 2026-09-26 on two temporary OneDrive documents (type 3 characters, wait
250 ms, switch, poll `Document.Saved`):

| How the window was left | Saved after the switch |
|---|---|
| no switch (baseline) | 1 of 4 saved after 7.4 s; 3 of 4 not within 12 s |
| a tab click (add-in v1.7.2) | 219-312 ms (4/4) |
| Alt+Tab | 218-281 ms (4/4) |
| Alt+Tab with the add-in DISCONNECTED | 218-234 ms (3/3) |
| SetForegroundWindow / Word's own Window.Activate from outside | 0-15 ms / 125 ms (3/3 each) |

The add-in does not and will not switch AutoSave off behind the user's back (it
never changes how the user's documents are saved). What the user can do: turn
the "AutoSave" toggle off for that document (Word remembers it per file), or for
every file: File > Options > Save > untick "AutoSave files stored in the Cloud by
default on Word" (the Russian wording varies between builds); then save with
Ctrl+S. AutoRecover still protects against crashes - it is a different feature
and never overwrites the document file. Details: docs/RESEARCH.md, "Word AutoSave
and window switching".
