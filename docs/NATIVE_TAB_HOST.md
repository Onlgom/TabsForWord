# NATIVE_TAB_HOST.md - tabs inside the Word window, without the CTP title strip

Started 2026-07-24 as an experiment in the branch `experiment/native-tab-host`.
Status: **THE MAIN MODE since version 1.6.0** (2026-07-26, decision ADR-015, the
branch merged into `master`). The classic Custom Task Pane remains as the
fallback mode - see sections 6 and 13.

The document is kept as an engineering record: why it is built this way, which
properties of the Word windows are relied on, and what was verified how.
Sections 1-2 were written while this was still an experiment and are left as
they were on purpose.

## 1. The goal of the experiment

Find out whether the existing tab strip (`TabStripControl`) can be placed
directly inside the Word window through the Win32 API in a stable way - so that
the strip looks like part of the Word interface and has **no Custom Task Pane
title strip** (the known limitation O2 in KNOWN_ISSUES.md, which the official
Office API cannot remove).

The main question: "is such an integration stable enough not to disturb work on
the documents?" The answer, in the end: **yes, the approach is technically
viable** (details in sections 11-13).

## 2. The original architecture (unchanged)

```
Connect -> WordEventManager -> DocumentWindowManager (the tab model)
                                   | SyncPanes/PushTabs
                             TabPaneManager -> CustomTaskPane -> TabStripControl
```

The strip is hosted in a Custom Task Pane (DockPosition=Top) in every Word
window. Word draws its own CTP title strip (about 35 px) above it, and that
cannot be removed.

## 3. The experimental architecture

```
Connect --(mode chosen at startup)--+
                                    v
DocumentWindowManager ---> ITabPaneSync (a new thin interface)
                              |-- TabPaneManager            (the classic CTP mode)
                              \-- NativeTabHostManager      (the experiment)
                                      |  dictionary: Word window HWND -> NativeTabHost
                                      |  + WinEventHookService (instant re-layout)
                                      |  + fallback to TabPaneManager on failures
                                      v
                                 NativeTabHost (one per Word window)
                                      |  a WinForms surface (WS_CHILD) -> SetParent
                                      |  into the top-level Word window (OpusApp)
                                      v
                                 TabStripControl (the very same control as in CTP)
```

The tab model, the Word events, Ctrl+Tab and the close/switch logic are reused
unchanged. The UI control is identical in both modes.

## 4. The classes that were added (src/TabsForWord/NativeHost/)

| Class | Responsibility |
|---|---|
| `NativeWin32` | P/Invoke only (x64-safe: IntPtr, Get/SetWindowLongPtr wrappers, GetDpiForWindow with a fallback) |
| `WindowTreeDiagnostics` | a snapshot and a dump of the Word window tree into Logs\WindowTree-*.txt (at most 20 files kept) |
| `TabHostSettings` + `TabHostMode` | reads native-host.cfg and WORDTABS_NATIVE_HOST (the default was CustomTaskPane back then; it is Native since v1.6.0) |
| `NativeHostLayoutMath` | pure geometry (the strip rect, DPI conversions) - covered by tests |
| `WordNativeWindowLocator` | finds the document area: the `_WwG` canvas -> the anchor container under the ribbon; a fallback that does not rely on the class name; the reason for the choice goes to the log |
| `NativeTabHost` | one host per window: creating the HWND, SetParent, layout, the reserve mode, idempotent recreation and destruction |
| `WinEventHookService` | EVENT_OBJECT_LOCATIONCHANGE for our own process only (WINEVENT_OUTOFCONTEXT) - instant re-layout without subclassing and without aggressive polling |
| `NativeTabHostManager` | the ITabPaneSync implementation; the host lifecycle; user actions; the emergency fallback to CTP |
| `ITabPaneSync` (in src/) | the single decoupling point: DocumentWindowManager knows only this interface |

The changes to the existing code are minimal: `TabPaneManager` declares that it
implements the interface (one line), `DocumentWindowManager` takes the interface
(two lines), `Connect` picks the mode at startup (about twenty lines). Nothing
was deleted.

## 5. How to switch the in-window mode on

The file `%LOCALAPPDATA%\TabsForWord\native-host.cfg` holds `key=value` lines:

```
mode=native      # native | ctp (native by default since v1.6.0)
reserve=1        # 1 = push the document area below the strip; 0 = the strip lies on top
dump=0           # 1 = dump the window tree when a host is created (for diagnostics)
bg=auto          # strip background: auto (default) = the Word work area colour,
                 # spec = the specification 2b colour (#F3F4F5), #RRGGBB = fixed
```

Since version 1.6.0 the config may be absent entirely: the defaults are
`mode=native`, `reserve=1`, `dump=0`, `bg=auto`. The file is only needed to
change something. The environment variable `WORDTABS_NATIVE_HOST=1|0` wins over
`mode` from the file.

Ready-made switches (double-click, then restart Word):

- `installer\In-window-tabs.cmd` - the main mode (native + reservation);
  part of the package;
- `installer\Classic-mode.cmd` - back to the Custom Task Pane; part of the package;
- `scripts\mode-overlay.cmd` - a service script: native without reservation (the
  strip lies over the top of the document and covers the ruler); NOT shipped.

## 6. How to get the CustomTaskPane back

Run "Classic-mode.cmd" from the unpacked package folder, the one Install.cmd was
run from (or write `mode=ctp` into native-host.cfg), and restart Word. The
switches are not copied into `%LOCALAPPDATA%\TabsForWord` - that folder holds
only the DLL, the logs and the cfg files. Deleting the config now means the MAIN mode
(native), not CTP.

The classic mode also switches itself on in two cases, with no user involved:
- the emergency fallback inside the current Word session - 5 failed
  synchronisations of the in-window host in a row (section 12), with an ERROR
  line "falling back to CustomTaskPane" in the log;
- an unexpected failure to read the settings (`TabHostSettings load failed`).
Both cases get their own item in the automatic summary of the diagnostics report.

## 7. The Word version everything was verified on

Microsoft 365 Word x64 (Windows 11 Pro 26200, Russian ribbon), July 2026.
Every observation below comes from that particular build. The v1.7.3
measurements (the click activation and the settle timer in section 9, ADR-018
and ADR-019) were made on the same machine in September 2026 with Word
16.0.20326, one 1920x1080 monitor at 100 % and the tab size multiplier 1.3
(a 48 px strip).

## 8. The Word HWND hierarchy that was discovered (from the WindowTree dumps)

```
OpusApp  "Document2 - Word"          <- the top-level window (SDI: one per document)
+-- MsoCommandBarDock "MsoDockTop"   (0,0)-(1920,178)  <- the ribbon dock
|   \-- MsoCommandBar "Ribbon"
+-- DropShadow                        the shadow strip under the ribbon
+-- MsoCommandBarDock "MsoDockBottom" the status bar
+-- _WwF                              (0,178)-(1920,1010) <- THE AREA UNDER THE RIBBON (the strip anchor)
|   \-- _WwB "Document2"
|       +-- _WwE                      the horizontal ruler (height 24)
|       +-- _WwE                      the vertical ruler (width 24)
|       +-- _WwG "Microsoft Word Document"  <- the document canvas
|       \-- NUIScrollbar -> NetUIHWND  the vertical scrollbar
\-- MsoWorkPane / NUIPane / NetUIHWND (hidden service windows)
```

`Word.Window.Hwnd` in SDI Word returns `OpusApp` itself. Every window has
`WS_CLIPSIBLINGS`, so our strip as a sibling causes no repaint conflicts.

While the File menu (Backstage) is open the same window looks completely
different - a field dump from machine C, 2026-07-28:

```
OpusApp                              visible
+-- FullpageUIHost                   (0,0)-(1920,1040)  VISIBLE, the whole client area
|   \-- NetUIHWND                    the Backstage interface itself
+-- _WwF / _WwB / _WwG               HIDDEN
+-- MsoCommandBarDock (all four)     HIDDEN
\-- MsoWorkPane (all)                HIDDEN
```

So there is no document area at all, and the correct behaviour is to hide the
strip until the menu is closed. `FullpageUIHost` is what the locator recognises
the state by (`IsFullPageUiActive`), and it is the reason the answer "no zone"
had to be split into a failure and a normal state - see ADR-017 and section 11.

## 9. How the place for the strip is chosen

1. Among the descendants of OpusApp, find the visible canvas of class `_WwG`
   (the largest one).
2. Walk up the ancestor chain of the canvas and take the topmost container whose
   top lies **below** the top of the client area (the threshold is 8 px x DPI):
   in practice that is `_WwF`. A container spanning the full client height would
   be skipped (it must contain the ribbon).
3. The strip top = the anchor top; the width = the horizontal range of the
   document area; the height = `TabStripControl.DesiredContentHeight`
   (37 px x DPI, or 16 px when collapsed).
4. If `_WwG` is not found (the class gets renamed in some future Office), the
   fallback is the topmost visible window below the client top that is at least
   40% as wide and at least 100 px tall.
5. Every choice is logged with its reason; after 6 failures in a row a window
   tree dump is written (the threshold was 3 at first; it was raised to 6 in
   v1.5.1, stage 23, together with the locator race fix).
6. "Nothing found" is two different answers. If a full-page UI window covers the
   client area (the File menu), it is a normal state: one INFO line, no failure
   counted, no dump, and the manager is told the sync succeeded - otherwise five
   polls inside the File menu would drop a single-window Word into the CTP
   fallback. Anything else is a real failure and keeps its counters (ADR-017).

The layout is refreshed instantly by the WinEvent LOCATIONCHANGE (our own
process, no DLL injection) and is backed by an idempotent ~1 Hz check from
Reconcile.

The `reserve=1` mode: after the strip is positioned, `_WwF` gets its top placed
exactly under the bottom of the strip (the bottom of `_WwF` is fixed). The
adjustment works both ways: after the strip is collapsed (37 -> 16 px) the anchor
moves up, after it is expanded it moves down; the host remembers where it put the
anchor and where the anchor's natural place is, so a change of strip height
neither "drags" it downwards nor leaves unpainted ghosts on screen (a problem of
an early version, fixed 2026-07-24). The recalculation always starts from the
actual geometry, so nothing accumulates. Protection against a layout fight with
Word: more than 8 adjustments within 2 seconds means a 30 s pause (with the
natural layout restored), and three such bursts in a row move that window to
overlay for good. Destroying the host restores the Word layout.

Appearance: the strip is placed in the z-order right BELOW the ribbon shadow
window (DropShadow), so the Word shadow falls onto the strip and it does not look
glued onto the ribbon. The strip background by default (`bg=auto`) is matched to
the colour of the Word work area. The measurement is taken FROM THE SCREEN
(GetPixel on the composited frame at points just below the strip, past the
vertical ruler): Word draws its canvas with Direct2D, and GetPixel from its DC
does not see the content. Every point is checked with WindowFromPoint for
belonging to our own window - areas covered by a foreign window are not sampled,
and white is rejected (that is the page). The matching is live and per window:
the shade of the Word work area DIFFERS between an active and an inactive window
(#E8EAEF versus #F0F0F0, for instance), so each host tracks its own window
(about 1.5 s of frequent probes after an activation change, then a check every
~3 s); the match is pixel-exact (delta = 0 in the E2E run).
In the in-window mode the fill of the active tab stops at the bottom line of the
strip, and the segment of that line is redrawn over its bottom
(SetActiveTabFlushBottom): the document is right below the strip, so letting the
fill flow through the line (which is what the CTP mode wants) produced a white
lip, and the anti-aliased edge of the fill lightened the line under the tab.

Live resize with the mouse: while the window size is changing the anchor is not
adjusted at all (the strip keeps following the window) - one adjustment after
`SizeSettleMs` = 125 ms of quiet (250 ms at first). That removes the ruler jitter
while the frame is dragged on slower machines and stops the oscillation guard
from firing on legitimate layout storms. Since v1.7.3 (ADR-019):
- a one-shot timer per host fires exactly when the quiet time ends. Before, the
  adjustment waited for the next WinEvent or the ~1 Hz poll, and after a snap or
  a maximise nothing else arrives - the ruler stayed under the strip for up to
  about a second (median 751 ms after an Aero Snap maximise; now about 210-240 ms);
- while the user still drags (GUI_INMOVESIZE on Word's thread, or a mouse button
  held) the timer does not move the anchor and looks again every 50 ms, so the
  page does not jump on every pause of the hand;
- a re-entrancy guard: if Word's WM_SIZE handler pumps messages during our
  SetWindowPos on the anchor, a nested `UpdateLayout` does not move the anchor a
  second time - it re-arms the timer instead;
- 125 ms is kept on purpose: Word's own follow-up layout pass comes about 80 ms
  after a size change, and a shorter quiet time would fit before it and then fit
  again (a visible flicker);
- every adjustment that answers a resize is logged with its delay:
  `Reserve: anchor moved by Npx (word=0x...), N ms after the resize began`
  (other adjustments are logged at most every 500 ms, without the delay).

Activation by click (since v1.7.3, ADR-018): a left or middle press on a tab
answers WM_MOUSEACTIVATE with "no activate" and does not take the keyboard focus,
so a press in a background window no longer activates that window and re-lays out
the tabs under the cursor. The switch is posted to run after the click is over:
a minimised target is restored with SC_RESTORE (it stays maximised if it was),
then Window.Activate(), SetForegroundWindow and BringWindowToTop (SetForegroundWindow
can fail silently during activation transitions - a mix of maximised and ordinary
windows - and BringWindowToTop additionally raises the window in the z-order).
The result is logged ("Window activated: hwnd=N (tab, sfw=1, in front, ...)"),
and a check at +250 ms and +1200 ms warns ("Activate check: target ... is not in
front") and re-raises the target once if the switch was undone with no user
input since the click. The focus then goes back to the document of the activated
window.

The tab context menu is a classic Win32 menu (TrackPopupMenu), not a
ContextMenuStrip. The reason (a Word crash found in field testing): the
ContextMenuStrip was disposed from its own Closed event - the ToolStrip
internals keep running after Closed, an ObjectDisposedException escaped from the
menu window procedure into the Word message loop (Word has no WinForms loop with
ThreadException) and brought WINWORD down with 0xE0434352 (repro: right-click a
tab, then click outside the menu). Even without the Dispose, the first ToolStrip
menu shown in a session would not close on a click outside (its closing depends
on WinForms message filters, which are not active inside a foreign loop). A
classic menu is modal and driven by user32 - it always closes and creates no
WinForms windows at all. In addition, the window procedures of TabStripControl,
HostSurface and TabListPopup are wrapped in try/catch, so a managed exception
from a WndProc can no longer bring Word down (it is logged as "… WndProc failed").

## 10. Scenarios that were tested (automated E2E on a real Word)

Scripts: `tests/manual/e2e-overlay-interaction.ps1`, `e2e-overlay-buttons-pv.ps1`,
`e2e-reserve-mode.ps1` (run them with Word closed; they create their own
temporary documents and use real mouse and keyboard input).

Verified and working:

- the strip appears in every Word window (one host per window, no duplicates);
- there is no CTP title strip; no CTPs are created in the in-window mode;
- the tabs show the current documents, with an amber dot for unsaved ones;
- clicking a tab activates the right window (`Window activated` in the log);
- clicking the close button closes the document the ordinary Word way;
- after a click the input focus stays in the document (typing works);
- the collapse and expand buttons work (37 <-> 16 px);
- window resize: the strip width follows the client area;
- maximise / restore - the strip stays in place;
- collapsing and expanding the ribbon - the strip moves instantly (178 <-> 78);
- a second window of a document (View -> New Window) gets its own strip;
- closing a window destroys its host; closing Word cleans everything up (the
  hooks are removed, the HWNDs destroyed, "Add-in stopped" in the log) and the
  WINWORD process exits without hanging;
- Protected View: the PV window opens without failures and the grey PV tab is
  present in the strips of ordinary windows (there is no strip inside the PV
  window itself - the same as in CTP mode);
- Ctrl+Tab / Ctrl+Shift+Tab work;
- reserve=1: the document area is pushed exactly below the strip (no overlap),
  stable while idle, during resize and when the ribbon is toggled; 2-6
  adjustments per window, with no oscillation and no shutdowns;
- 27/27 unit tests pass; the logs of the experimental runs contain 0 ERROR and
  the CTP fallback never fired.

Additionally verified after the 2026-07-24 refinements (from the user's field notes):

- three collapse/expand cycles in a row: the strip top does not move (Y=178 in
  every phase) and the strip and the document area are always flush, with no gap
  and no overlap - no ghosts;
- a chaotic mix of collapsing and resizing: the reservation is not switched off
  and the ruler does not get stuck under the strip (the strip bottom equals the
  top of the document area);
- the strip background matched the Word work area pixel for pixel (delta <= 3),
  and the strip sits in the z-order below the ribbon shadow (the shadow falls
  onto the strip);
- the tab context menu: three cycles of "right-click, then click outside",
  closing with Escape, clicking the "Close" item, and the all-tabs menu closed by
  clicking outside - Word stayed alive with zero new crash events in the Windows
  event log (this is where the 0xE0434352 crash used to be; fixed, see section 9).

## 11. Known problems and observations

1. The document zone is not found instantly when a Word window is created: for
   the first 0.5-1 s the locator sees a half-built tree (a WARN in the log) and
   the strip appears on the next check. Only noticeable in the log.
2. In overlay mode (reserve=0) the strip covers the top ~37 px of the document
   area - the ruler when it is on, the top of the page when it is off. That is
   why reserve=1 is the default.
   (The historical problem "the strip drifts down with ghosts on collapse/expand
   and the ruler gets stuck under it" was fixed 2026-07-24 - see section 9.)
   Background matching (bg=auto) works in Print Layout; in Web Layout and Draft
   the canvas is white and the specification colour is kept. In the Office dark
   theme the background will be matched dark while the tabs stay light (the
   specification 2b palette) - a dark palette is out of scope here.
3. When a window closes, a single "extra" shift of the dying `_WwF` is possible
   (a "Reserve: anchor moved" line before the destroy) - harmless, the window
   disappears anyway.
4. The strip is added to document windows only; there is none in a Protected View
   window (the same as the CTP limitation, O3).
5. The all-tabs button and the context menu: checked automatically by the E2E
   scripts (opening and closing, choosing items, no Word crashes), and a manual
   run by the user was done 2026-07-25 with no complaints. CLOSED.
6. Different DPIs and monitors: the code takes DPI into account (GetDpiForWindow
   plus a height conversion, and the WinForms DPI events). Verified 2026-07-25 on
   three monitors connected AT THE SAME TIME with different scales
   (100% / 125% / 150%): the Word window reports the per-monitor DPI on each
   screen (96/120/144), the strip height and the drawn content scale
   proportionally (43 -> 54 -> 65 px), returning to the original monitor restores
   the size, the log has no ERROR and Word exits cleanly - 13/13 PASS
   (tests/manual/e2e-dpi-monitors.ps1; details in TEST_RESULTS.md).
   Changing a monitor's scale on the fly without moving the window is still not
   automated (it needs a sign-out and back in), but it is covered by the same
   ~1 Hz check. PASS.

## 12. What happens after an Office update

- The class names `_WwG`/`_WwF`/`_WwB` have been stable for decades but are not
  documented. If they change, the fallback heuristic (section 9, item 4) takes
  over; on a complete failure the strip hides first, and both the add-in and
  Word keep working. If the failure persists for 5 synchronisations in a row,
  the add-in falls back to CTP automatically within the session (the emergency
  contour below); Classic-mode.cmd makes CTP the permanent choice.
- The internal layout (the ribbon height and so on) is read dynamically - the
  dependency is on geometry, not on constants.
- The emergency contour: 5 failed synchronisations of the in-window mode in a row
  trigger an automatic fallback to the CustomTaskPane within the same session.

## 13. Conclusion: the move was made (v1.6.0, 2026-07-26)

The observed facts in favour: the HWND host is stable (Word neither destroys nor
reparents it), the layout is tracked without subclassing and without polling,
focus and input are not broken, the reserve mode causes no layout wars, and the
unloading is clean. The multi-DPI scenarios were verified 2026-07-25 (section 11
item 6, 13/13 PASS on three monitors at 100/125/150%). Against: the Word
hierarchy is undocumented (insured against by the fallback heuristic).

Both conditions of the earlier recommendation are met: items 5 and 6 of section
11 are closed, and CTP is kept as a permanent fallback contour. Field
confirmation was added as well - the package with the in-window mode worked on
two other people's machines (machine A and machine B), both installations
succeeded and Word never crashed; the single real defect (the locator race) was
found in the logs and fixed in v1.5.1.

The decision was taken 2026-07-26: native is the main mode, the defaults in the
code are `mode=native, reserve=1`, the branch is merged into `master` and the
product ships as a single package. The full reasoning, the rejected options and
the risks are in **ADR-015** in docs/DECISIONS.md. The remaining permanent risk
(undocumented Word window classes) is recorded as limitation **O17** in
docs/KNOWN_ISSUES.md.

## 14. Rollback: how to get the pre-1.6.0 behaviour back

1. A user-level rollback (no rebuild): "Classic-mode.cmd" from the unpacked
   package folder (where Install.cmd is) - the strip returns to a Custom Task
   Pane with its title strip.
   Back again: "In-window-tabs.cmd".
2. A full code rollback: the in-window host lives in `src/TabsForWord/NativeHost/`
   and is wired in at exactly one place - the mode choice in
   `Connect.TryInitialize()`. The default value lives in `TabHostSettings`
   (`_mode`, `_reserveSpace` and the defaults inside `Parse`); setting them back
   to `CustomTaskPane`/`false` is enough to make the product CTP-only again.
3. The historical state "before the experiment" is commit `4d9a977`
   (checkpoint-12-ui-2b), the last one in `master` before the merge.
