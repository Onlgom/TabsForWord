# ARCHITECTURE.md - the architecture of TabsForWord (stage 2)

Written 2026-07-23, based on docs/RESEARCH.md and docs/DECISIONS.md (ADR-001…006).
Extended 2026-07-26 (v1.6.0) with "Two strip hosts": from that version on, the
main host is not a Custom Task Pane but our own window inside the Word window.
Updated 2026-09-26 (v1.7.3): the event flows name the shared `ITabPaneSync`
calls instead of CTP-only steps, and a tab switch is posted and verified
(ADR-018).

## Two strip hosts (since v1.6.0)

The tab model, the Word events and the `TabStripControl` itself do not depend on
the host. `DocumentWindowManager` works through the thin `ITabPaneSync` interface
(`SyncPanes` / `PushTabs`), and the concrete host is chosen once at startup in
`Connect.TryInitialize()`:

| Mode | `ITabPaneSync` implementation | When it is used |
|---|---|---|
| **Native** (main) | `NativeHost.NativeTabHostManager` - our own WinForms window, `SetParent` into the Word window, anchored to `_WwF` | by default (ADR-015) |
| Classic | `TabPaneManager` - a Custom Task Pane, DockPosition=Top | `mode=ctp` in native-host.cfg, `WORDTABS_NATIVE_HOST=0`, or the emergency fallback |

`TabPaneManager` is created ALWAYS (in native mode as the fallback contour inside
`NativeTabHostManager`), so falling back to CTP is possible at any moment of a
Word session and needs no restart. The in-window host has its own document,
docs/NATIVE_TAB_HOST.md; the diagram below describes the shared pipeline and, in
detail, the classic host.

## The big picture

One build artefact: `TabsForWord.dll` (net48, AnyCPU, WinForms).
Word loads it as a classic COM add-in (mscoree -> CLR 4 -> Connect).

```
Word Application (WINWORD.EXE, x64)
        |
        |  COM activation through HKCU\...\Word\Addins\TabsForWord.Connect (LoadBehavior=3)
        v
   Connect  (IDTExtensibility2 + ICustomTaskPaneConsumer)   [entry point, everything in try/catch]
        |
        |  OnConnection: receives Word.Application
        |  CTPFactoryAvailable: receives ICTPFactory
        |  OnStartupComplete: starts the managers + the initial Reconcile
        v
   WordEventManager
        |   subscriptions (delegates kept in fields!):
        |   DocumentOpen, NewDocument, DocumentChange,
        |   DocumentBeforeClose(hint), DocumentBeforeSave,
        |   WindowActivate, WindowDeactivate,
        |   ProtectedViewWindowOpen/Activate/BeforeClose(+3), Quit
        |            |
        |            v      every event leads to one call
        |   DocumentWindowManager.RequestReconcile(reason)
        v
   DocumentWindowManager        [owner of the model]
        |   - builds a snapshot: Application.Windows + ProtectedViewWindows
        |   - diffs it against the list of DocumentTabModel
        |   - a debounce tick ~300 ms after BeforeClose
        |   - a background poll ~1000 ms (Save As, renames, the Saved flag)
        |   - hands the updated list to every strip
        v
   TabPaneManager               [strips per window - the classic host; the main
        |                        in-window host is NativeTabHostManager, docs/NATIVE_TAB_HOST.md]
        |   - Dictionary<hwnd, PaneInfo{CustomTaskPane ctp, TabStripControl ctrl}>
        |   - for every Word.Window without a strip: CreateCTP(progId, title, window)
        |     -> DockPosition=Top -> DockPositionRestrict=NoChange -> Height -> Visible
        |   - probes for dead CTPs (COMException) -> cleanup + ReleaseComObject
        v
   TabStripControl (a WinForms UserControl, COM-visible)  - in every Word window
        |   - draws the tabs: the name, the unsaved indicator, the active highlight, the close button, [+]
        |   - horizontal scrolling on overflow
        |   - tooltip: the full path
        |   - clicks -> callbacks into DocumentWindowManager (RequestActivate / RequestCloseWindow / CreateNewDocument)
        v
   LoggingService  (%LOCALAPPDATA%\TabsForWord\Logs\TabsForWord-yyyyMMdd.log)
```

## Components and responsibilities

| Component | File | Responsibility |
|---|---|---|
| `Connect` | src/TabsForWord/Connect.cs | The COM entry point; the lifecycle; wires everything together; no tab logic of its own |
| `ComInterop` | src/TabsForWord/ComInterop.cs | Hand-written [ComImport] interfaces: IDTExtensibility2, ICTPFactory, ICustomTaskPaneConsumer, _CustomTaskPane plus the enums |
| `WordEventManager` | src/TabsForWord/WordEventManager.cs | Subscribing to and unsubscribing from Word events; keeping the delegates alive; translating them into RequestReconcile |
| `DocumentWindowManager` | src/TabsForWord/DocumentWindowManager.cs | The tab model; Reconcile(); the Activate/Close/New actions (a switch is posted and run after the click); debounce + polling |
| `WindowActivator` + `ActivationWatch` | src/TabsForWord/WindowActivator.cs | The Win32 half of a switch (since v1.7.3): SC_RESTORE for a minimised window, SetForegroundWindow + BringWindowToTop, a log description of a window (hwnd, class, state - never the title), and the check at +250 / +1200 ms with at most one re-raise (ADR-018) |
| `DocumentTabModel` | src/TabsForWord/DocumentTabModel.cs | A POCO: Hwnd, Caption, FullPath, Saved, IsActive, IsProtectedView, TabColor, IsPinned |
| `TabPaneManager` | src/TabsForWord/TabPaneManager.cs | The classic host: one CTP per window; creation, cleanup, recreation; handing the model to the controls |
| `ITabPaneSync` | src/TabsForWord/ITabPaneSync.cs | Decouples the model from the strip host (SyncPanes/PushTabs) |
| `NativeTabHostManager` + `NativeHost/*` | src/TabsForWord/NativeHost/ | The main host: the strip window inside the Word window, the anchor locator, WinEvent-driven layout, the emergency fallback to CTP (see docs/NATIVE_TAB_HOST.md) |
| `TabStripControl` | src/TabsForWord/TabStripControl.cs | Drawing the tabs and all user interaction (in the in-window mode a press on a tab does not activate the window it lives in - WM_MOUSEACTIVATE) |
| `TabListPopup`, `SettingsForm` | src/TabsForWord/TabListPopup.cs, SettingsForm.cs | The all-tabs menu and the settings window |
| `KeyboardHookService` + `HotkeySettings` | src/TabsForWord/KeyboardHookService.cs, HotkeySettings.cs | Ctrl+Tab / Ctrl+Shift+Tab (a local hook on Word's UI thread, only while a Word document window is in front) and its on/off switch |
| `TabColorStore`, `TabOrderStore`, `TabSizeSettings` | src/TabsForWord/ | The persisted settings: tab colours, order and pinning, tab size |
| `TabTheme`, `Glyphs` | src/TabsForWord/ | The specification 2b palette and the drawn icons |
| `Strings`, `UiLanguage` | src/TabsForWord/ | The interface captions and the language choice (ADR-016) |
| `LoggingService` | src/TabsForWord/LoggingService.cs | A file log with levels, one file per day, the 7 newest files kept, never containing document content |

## Event flow (the main scenarios)

### Opening or creating a document
```
Word -> DocumentOpen/NewDocument/DocumentChange
     -> RequestReconcile
     -> snapshot of Windows: window N appeared -> create a DocumentTabModel
     -> ITabPaneSync.SyncPanes: window N has no strip -> create one
        (a NativeTabHost in the main mode, a CTP in the classic one)
     -> every control gets the new list -> repaint
```

### Clicking a tab (switching; since v1.7.3, ADR-018)
```
TabStripControl(window A) press on tab B
     -> in-window mode: WM_MOUSEACTIVATE answered "no activate" (window A is not
        activated, the strip does not take the focus, the tabs do not move)
     -> release on the same tab -> DocumentWindowManager.RequestActivate(hwnd B)
        (posted with BeginInvoke: runs after the click is over; a double click = one switch)
     -> ActivateWindow(hwnd B):
          SC_RESTORE if minimised (a maximised window comes back maximised)
          -> COM window.Activate()
          -> SetForegroundWindow + BringWindowToTop (WindowActivator)
          -> log "Window activated: hwnd=B (tab, sfw=1, in front, from ... to ..., N ms)"
     -> Reconcile -> the active tab is highlighted in every window
     -> the focus goes back to the document of window B (in-window host)
     -> ActivationWatch at +250 / +1200 ms: B not in front -> WARN "Activate check";
        re-raise once only if the foreground went back to window A with no input since
```
Enter on the strip and the rows of the all-tabs menu go through the same
RequestActivate; Ctrl+Tab uses the same posting (queued key presses add up).

### The close button on a tab
```
TabStripControl click on [x] of tab B
     -> DocumentWindowManager.RequestCloseWindow(hwnd B)
     -> document.Close(wdPromptToSaveChanges)      <- the standard Word dialog
        |- the user picks Save / Don't Save -> the window closes
        |   -> the host of window B is disposed on the next sync
        |      (classic mode: Word destroys the CTP of window B itself)
        |   -> DocumentChange / the deferred tick -> Reconcile -> the tab disappears everywhere
        \- the user picks Cancel -> COMException (Word code 4198)
            -> caught, logged at INFO, the tab stays
```

### Save / Save As
```
Word -> DocumentBeforeSave(SaveAsUI)
     -> mark "a rename is expected" + a Reconcile poll afterwards
     -> compare FullName/Saved against the model -> update the tab captions
The background timer (~1 s) additionally catches autosaves and renames that
arrive without an event.
```

### Protected View
```
ProtectedViewWindowOpen -> a tab marked as Protected View (view only; there is NO
                           strip inside the PV window itself - add-ins are not hosted there)
ProtectedViewWindowBeforeClose(CloseReason=wdProtectedViewCloseEdit)
     -> DocumentOpen for the same file follows -> the tab turns into an ordinary one
```

### Unloading (OnDisconnection / Quit)
```
unsubscribe every event -> stop the timers -> dispose the strip hosts (in-window:
remove the WinEvent hook, destroy the hosts, which restores the Word layout;
classic: Delete() the live CTPs; all in try/catch)
-> ReleaseComObject in the right order -> log "Add-in stopped"
```

## Where state lives

- The single source of truth while running is the live Word collections
  (Application.Windows, ProtectedViewWindows). The tab model is a derived cache.
- The key of a window is `Window.Hwnd` (int). Window/Document RCW objects are
  never compared and never cached beyond a single Reconcile (Application and the
  CTPs are the exceptions).
- Nothing is persisted between Word sessions (out of scope for the MVP; tab
  colours, order and pinning were added later and do persist - see the ADRs).

## Updating the tabs

Reconcile() is idempotent and singular (serialised through Word's main STA
thread; the timers are WinForms timers ticking on the same thread, so there are
no races):

1. Snapshot: `[ {hwnd, caption, fullName, saved, active, pv} ]` from the live collections.
2. Diff against the model: added / removed / changed.
3. `ITabPaneSync.SyncPanes(liveWindows, tabs, changed)` on every Reconcile: the
   host creates a strip for every window that has none (a NativeTabHost, or a
   CTP in the classic mode) and re-lays out the existing ones.
4. Strips of windows that are gone are cleaned up in the same call (dead hosts
   pruned, dead CTPs deleted).
5. If anything changed -> hand every TabStripControl the new list (one snapshot object).
6. A control repaints only if its current list actually differs (no flicker).

## Error handling

- All five IDTExtensibility2 methods, CTPFactoryAvailable, every Word event
  handler and every UI event handler are wrapped in try/catch ->
  LoggingService.Error. An exception NEVER crosses the add-in boundary
  (otherwise Word flips LoadBehavior 3 -> 2 or moves us to DisabledItems).
- Every touch of a CTP/Window/Document is a potential COMException (the window
  may have died) -> targeted try/catch degrading to "skip it and rebuild on the
  next Reconcile".
- A cancelled close (4198) is a normal path, logged at INFO.
- A failure to create the strip (host or CTP) for one window does not affect
  the others (logged, and the loop continues); in the in-window mode a failure of
  every window for 5 syncs in a row switches to the CTP fallback.

## Logging

`%LOCALAPPDATA%\TabsForWord\Logs\TabsForWord-yyyyMMdd.log`, in the format
`2026-07-23 12:00:00.123 [INFO ] Tab created: hwnd=..., name=...`
Levels INFO/WARN/ERROR; the exception and its stack trace when there is one;
one file per day, and only the 7 newest files are kept (pruning is by count, not
by age). Document content is never written; the
log does contain document names and paths, which is spelled out for the user in
the installation guide and in the diagnostics output.

## Performance budget

- Reconcile is O(windows); typically under 1 ms for 10 windows; it runs on events
  plus a 1 Hz background tick (only string and flag comparisons, with no COM
  calls heavier than reading a property).
- The control repaints only when the list has really changed.
