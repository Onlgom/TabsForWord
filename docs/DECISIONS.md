# DECISIONS.md - Architecture Decision Log

---

## ADR-001. Do not build "real" tabs inside one physical window

**Decision:** implement a synchronised tab strip in every Word window; a click on
a tab activates the corresponding window. Do not merge the windows.

**Reason:** MDI was removed from Word 2013 onwards entirely (officially,
"Replacement: N/A"); the object model has no API for merging or hosting windows;
the only route left is undocumented Win32 reparenting, which Microsoft itself
describes as dangerous (DPI, input queues, window styles) and which the task
brief forbids outright in the absence of official support.

**Alternatives considered:** (a) SetParent plus hooks, the way Office Tab does it;
(b) a shell host of our own wrapped around Word.

**Why they were rejected:** (a) unsupported, fragile against C2R updates, a risk
of data loss; (b) out of scope for the MVP and even more fragile.

**Date:** 2026-07-23

---

## ADR-002. A classic COM add-in (IDTExtensibility2) instead of VSTO

**Decision:** a C# class library targeting net48; the Connect class implements
hand-declared [ComImport] interfaces IDTExtensibility2 + ICustomTaskPaneConsumer.
No VSTO.

**Reason:** the machine has no VS, no MSBuild, no VSTO SDK and no VSTO Runtime -
there is nothing to build a VSTO project with, and installing Visual Studio is
heavy and risky. A COM add-in, by contrast: the mechanism is fully supported by
modern M365 Word (confirmed by an adversarial check), it is already used by other
add-ins on this machine, it builds with the dotnet SDK that is present, and it
installs per-user without administrator rights. The VSTO Runtime is not needed on
the target machine at all.

**Alternatives considered:** (a) a VSTO Add-in (what the task brief preferred);
(b) Office.js; (c) a separate application driving Word through UI Automation.

**Why they were rejected:** (a) technically cannot be built in this environment;
the brief explicitly allows changing the stack after research; (b) Office.js
gives no control over desktop Word's windows (and the brief forbids it);
(c) fragile and unofficial.

**Date:** 2026-07-23

---

## ADR-003. The UI host: a Custom Task Pane (top-docked) + a WinForms UserControl

**Decision:** the tab strip is a WinForms UserControl registered as a COM class
and hosted through ICTPFactory.CreateCTP with msoCTPDockPositionTop; one CTP per
Word window (keyed by Window.Hwnd), DockPositionRestrict=NoChange.

**Reason:** this is the only official way to put UI of our own at the top inside
the Word window. Top docking in Word is confirmed by the documentation. WinForms
is used because CreateCTP requires a COM content class (ActiveX hosting), and a
WinForms UserControl is registered by the same mscoree scheme without any extra
ActiveX keys.

**Alternatives considered:** (a) WPF (through ElementHost); (b) a floating window
over the Word window with position tracking; (c) a Ribbon menu with a list of
documents and nothing else.

**Why they were rejected:** (a) an extra layer on top of the very same WinForms
host; (b) an unofficial overlay - fragile, with Z-order and DPI problems; kept as
a documented fallback in case the minimum CTP height turns out to be
unacceptable; (c) does not give "permanently visible tabs" - it does not meet the
requirements.

**Known consequence:** Word enforces a minimum CTP height (roughly 70-120 px
instead of the desired 32-50 px) and draws its own title strip with a close
button on the left. The actual minimum is measured in stages 4-5 and recorded in
KNOWN_ISSUES.md.

**Date:** 2026-07-23

---

## ADR-004. An idempotent Reconcile instead of trusting individual events

**Decision:** every Word event (DocumentOpen, NewDocument, DocumentChange,
DocumentBeforeClose, DocumentBeforeSave, WindowActivate/Deactivate, the Protected
View events) only triggers one shared recomputation of the model: a diff of the
current Application.Windows + ProtectedViewWindows against the tab model. Plus a
deferred tick (about 300 ms) after BeforeClose and an infrequent background poll
(about 1 s) for Save As and renames.

**Reason:** Word has no after-events (no AfterClose, no AfterSave); BeforeClose
can be cancelled by the user after the event has fired; the order of events on
File -> Exit is different again. An idempotent Reconcile absorbs every timing
quirk and cannot lose data (the add-in itself only ever calls Activate() and
Close(wdPromptToSaveChanges)).

**Alternatives considered:** adding and removing individual tabs in response to
each event.

**Why it was rejected:** races and desynchronisation on cancelled closes, Save As
and multiple windows; more complex and less reliable.

**Date:** 2026-07-23

---

## ADR-005. Building with the dotnet 8 SDK (user scope), interop through the NuGet PIA plus hand-written interfaces

**Decision:** an SDK-style csproj (net48, UseWindowsForms, AnyCPU,
LangVersion 7.3); the Word object model comes from the
Microsoft.Office.Interop.Word 15.0.4797.1004 package with embedded interop types
(the official workaround through Target/ReferencePath); IDTExtensibility2,
ICTPFactory and ICustomTaskPaneConsumer are hand-written [ComImport]
declarations. The backup interop route is the genuine PIA from the GAC
(HintPath).

**Reason:** everything was verified by an actual build on this machine, and the
resulting DLL does not depend on interop assemblies (NoPIA). There is no official
office.dll on NuGet.

**Date:** 2026-07-23

---

## ADR-006. Installation: a PowerShell script, HKCU, a 64-bit import

**Decision:** Install/Uninstall/Repair are PowerShell scripts (plus .cmd wrappers
with ExecutionPolicy Bypass): the DLL goes to %LOCALAPPDATA%\TabsForWord, is
Unblock-File'd, and the registry gets HKCU\Software\Classes (CLSID/ProgId -
strictly from a 64-bit process, because CLSID is redirected under WOW64) and
HKCU\...\Word\Addins\<ProgId> (LoadBehavior=3), plus a cleanup of
Resiliency\DisabledItems. Installing and uninstalling happen with Word closed
(never kill the process).

**Refinement (2026-07-25):** the Resiliency\DisabledItems cleanup is targeted
only. The entries are binary, with the DLL path and the add-in name stored inside
as UTF-16; install.ps1 and repair.ps1 delete only the entries that mention
TabsForWord. The earlier code deleted the whole key, which silently re-enabled
OTHER people's add-ins that Word had disabled after they crashed - an
unacceptable side effect (found by an external audit).

**Reason:** entirely per-user, no UAC; the path was verified end to end on this
machine. MSI/ClickOnce are unavailable (no toolchain) and unnecessary for an MVP.

**Date:** 2026-07-23

---

## ADR-007. Implementing "tab specification 2b" (stage 12)

**Decision:** the strip is drawn entirely by custom paint using the fixed palette
of the specification (TabTheme; not system colours). The outline of the active tab
is a single GraphicsPath (R8 arcs blended into the border line plus R7 corners),
with the stroke offset by half a pixel exactly as in the reference SVG. The tab
order belongs to DocumentWindowManager (drag-and-drop goes through ReorderTab ->
PushTabs to every strip; Reconcile preserves the positions). Ctrl+Tab /
Ctrl+Shift+Tab use a local WH_KEYBOARD hook on the Word UI thread only (not
global, not low-level); exceptions inside the hook are swallowed. The CTP title is
left empty (" ") because the strip cannot be removed through any official API; the
collapse chevron changes the CTP height by the delta between the actual and the
desired control height. The all-tabs menu is an ordinary borderless WinForms form
(DWM rounding, CS_DROPSHADOW) that closes on Esc or Deactivate. Every size scales
from DeviceDpi (the specification is written in pixels at 100 %).

**Alternatives considered:** Word KeyBindings for Ctrl+Tab (they modify
Normal.dotm and survive uninstalling the add-in - rejected); hiding the CTP title
with Win32 (a hack against someone else's Word window - forbidden by the project
rules); a global WH_KEYBOARD_LL (broader than necessary, a risk to the whole
system - rejected).

**Date:** 2026-07-23

## ADR-008. Tab colour: stored by file path, a classic owner-drawn menu (stage 13)

**Decision:** the user's tab colour lives in TabColorStore: for saved documents it
is keyed by the full file path (case-insensitively) in
`%LOCALAPPDATA%\TabsForWord\tab-colors.cfg` (the format is `RRGGBB|path`, colour
first, because a path may contain any character), so the colour survives closing
the document and restarting Word; for unsaved documents it is keyed by the window
hwnd in memory only (until it is saved or closed; when the window closes the entry
is forgotten). The tab model is shared (TabColor is part of DocumentTabModel and
is included in SameAs), so the colour is visible in every window at once and in
both modes (CTP and native). What is drawn on the tab is not the colour itself but
a pastel derivative of it (TabTheme.Tint*: 62-82 % white) - the amber edge and the
unsaved dot stay legible over any tone. The "Tab colour" submenu was added to the
classic Win32 menu (ContextMenu): the presets are owner-drawn MenuItems
(WM_MEASUREITEM/WM_DRAWITEM travel through the control's already-guarded
try/catch WndProc), and "More colours…" is the system ColorDialog (the native
modal ChooseColor, which needs no WinForms message loop; the preset palette is
preloaded into CustomColors).

**Alternatives considered:** storing in the HKCU registry (harder to clean up and
to migrate - rejected); keying an unsaved document by its name ("Document1" -
collisions between sessions, rejected in favour of the hwnd); writing the colour
into the document's own properties (that modifies the user's file - forbidden by
the project rules); a ContextMenuStrip for prettier items (it brought Word down -
see the history of crash fix 1c26907; banned).

**Date:** 2026-07-24

## ADR-009. Diagnosing "it doesn't work on my friend's machine", and support for 32-bit Word (stage 14)

**Decision:** three levels of data for remote diagnosis. (1) The installer writes
a full log to Logs\install-*.log: the Windows and .NET versions, the path and
bitness of winword.exe (from the PE header), the Click-to-Run configuration, the
SHA256 and assembly name of the DLL, and the result of reading every registry key
back. (2) At startup the add-in leaves a beacon in the constructor of the COM
class ("Connect instantiated") - it appears before any call into Word and so
separates "Word never reached the DLL" from "a failure inside IDTExtensibility2" -
plus an Environment line (the Word version and build, the OS, the bitness, the
language). (3) Diagnostics.cmd assembles a ZIP on the desktop: the system, Word,
.NET, our own files (including the Mark-of-the-Web), the registry in BOTH views,
the Resiliency bans, the WINWORD and .NET errors from the event log for the last
14 days, all the logs, and an automatic summary of the typical causes. The script
only reads the system.

A critical fix that came with it: HKCU CLSID keys are redirected by WOW64, so
registration now writes into both the 64-bit view and WOW6432Node - otherwise
32-bit Word (a common Office installation) silently fails to find
InprocServer32. The assembly name in the registry is now read from the DLL rather
than hard-coded as 1.0.0.0.

**Alternatives considered:** telemetry over the network (unacceptable: privacy,
complexity, and an MVP has no server - rejected); asking people to send
screenshots (does not reproduce the cause - rejected); Start-Transcript instead of
a hand-written log (it loses the structure, and Cyrillic under PowerShell 5.1 -
rejected).

**Date:** 2026-07-24

## ADR-010. Per-monitor DPI: draw in your own space, convert at the boundary (stage 15)

**Decision:** on multi-DPI configurations Windows virtualises windows whose DPI
context does not match, and WinForms on .NET 4.8 inside Word creates its windows
system-aware regardless of the thread context (verified on a live Word:
GetWindowDpiAwarenessContext of both the host window and the strip reports
system-aware on any monitor). Fighting that is pointless; instead every window
draws in the DPI space it actually lives in: the scale of the strip is
GetDpiForWindow of its own window (TabStripControl.SetDpiOverride; DeviceDpi is
not used - it drifts), and a conversion happens at the boundary between spaces:
the physical height of the pane inside the Word window = ScaleHeight(strip height,
strip DPI, Word window DPI). The formula is universal: if the contexts ever do
match, the conversion becomes an identity. The layout of the native host runs
inside a PMv2 thread context (SetThreadDpiAwarenessContext for the duration of the
calculations): Word switches the context of its own UI thread, and our coordinates
used to come back virtualised. ApplyRect compares the desired rect against the
actual window geometry (not against a cache), so the strip repairs itself after
any outside interference. The WM_DPICHANGED_BEFOREPARENT/AFTERPARENT messages are
swallowed: the WinForms automatic response resized the host window at the wrong
moment. In CTP mode the same rule applies without any conversion: the control, the
CTP host and ctp.Height all live in one space.

**Alternatives considered:** a PMv2 host window (a top-level Form + WS_CHILD +
SetParent) - the strip window inside it is still created system-aware by WinForms,
so the virtualisation remains (verified live); scaling from the Word window DPI
directly - the content was then drawn outside the window's actual space
(squeezed and soft); the control's DeviceDpi - not synchronised with the monitor
and changes at the wrong moments.

**Known quirk:** on a monitor whose scale differs from the system one, the picture
of the strip is stretched by the system (slightly soft text) - that is the price
of system-aware WinForms windows; the sizes themselves are correct on every
monitor.

**Date:** 2026-07-25

## ADR-011. Manual tab size: a multiplier on top of the DPI scale (stage 16)

**Decision:** the user's "Tab size" is a single multiplier (0.85 / 1.0 / 1.15 /
1.3) applied on top of the automatic per-monitor scale: the final scale of the
strip = DPI/96 x multiplier. The order matters: first the honest automatic fit to
the monitor (stage 15), then the user's taste - otherwise the very same setting
would look different on different monitors. The setting is global (not per tab and
not per window), is changed from the context menu of any tab, applies to every
window instantly (in native mode through UpdateLayout, in CTP mode by correcting
the pane height by the delta, exactly as collapsing does), is stored in
tab-size.cfg (scale=…, InvariantCulture) and is clamped to [0.7; 2.0] both when
read and when set.

**Alternatives considered:** a free slider or percentages (a dialog is extra UI
for an MVP; four presets are enough, and the clamp plus the file format leave the
road open to arbitrary values); Ctrl+mouse wheel over the strip (it conflicts with
scrolling the tabs on overflow); a per-window setting (the windows of one Word
would look mismatched).

**Date:** 2026-07-25

## ADR-012. Pinning tabs, persisting the order, and the settings window (stage 22)

**Decision:** three related features in one stage.

*Pinning* - `DocumentTabModel.IsPinned`; pinned tabs always form the first group.
The invariant "pinned first" is the same one trick
(`_tabs.OrderBy(t => t.IsPinned ? 0 : 1)`, a stable sort) applied in three places:
`Reconcile`, `ReorderTab`, and the live drag preview
(`TabStripControl.ApplyPinOrderInvariant`) - and because the sort is stable,
dragging a pinned tab towards the ordinary ones "sticks" to the group boundary
predictably, with no separate index-clamping code. "Close others" does not touch
anyone's pinned tabs (protection from accidental mass closing is the whole point
of pinning); closing a single tab (the close button, the middle mouse button) does
not obey pinning.

*Persisting the order* - it is restored from the file ONCE, on the first reconcile
after Word starts (`_startupOrderApplied`); after that it lives in memory as
before (reordering and new tabs go to the end). Pinning and order are kept in one
file, `TabOrderStore` (`tab-order.cfg`, in the same format as `TabColorStore`: the
path for saved documents, an hwnd session key for unsaved ones), rather than in
two, because the order of the lines in the file IS the order of the tabs -
separate storage would create a risk of them drifting apart. **An important detail
found by an E2E test**: `SaveOrder` MUST APPEND rather than replace paths that are
not in the current list of open tabs - otherwise closing ALL tabs before quitting
Word (the ordinary `Quit` path) would write an empty list and wipe the pinning of
documents the user is going to open next time. The first version of the code did
not do that, and E2E (`e2e-settings-pins-order.ps1`, session 2) caught it as
`Tab order loaded: 0 (pinned: 0)` after a restart even though the document had
been pinned and saved in session 1.

**An important class of bug, caught by the multi-agent review (see below) and
confirmed on a live Word**: `SetTabPinned`/`SetTabColor` changed the `IsPinned` /
`TabColor` field IN PLACE on `DocumentTabModel` objects that had already been
handed to the strips by a previous `PushTabs`. `TabStripControl.UpdateTabs`
decides whether to repaint by comparing by value (`ListsEqual`/`SameAs`) against
its own previous list - but with an in-place mutation that is the very same
object, the comparison finds it "equal to itself", and `Invalidate()` is skipped.
It is most visible when pinning or recolouring does not change the tab's position
(a single open tab is the commonest case). Fixed with `DocumentTabModel.Clone()`
plus replacing the changed tabs with fresh clones instead of mutating the field.

*The settings window* - `SettingsForm`, styled like `TabListPopup` (the same family
of popup windows as the strip: borderless, DWM rounding, the `TabTheme` palette,
closing on Esc or on losing focus). It opens from the footer of the all-tabs menu
("Settings…") rather than from a separate button on the strip itself - the
geometry of the strip is fixed by specification 2b, whereas the footer of the
popup menu is already flexible in height. The contents are only the things that
genuinely needed a home: the tab size (duplicating the context-menu presets for
discoverability) and the Ctrl+Tab switch (which lifts limitation O8 - the
interception of the ordinary tab character inside tables). The CTP/native mode
switch was deliberately NOT surfaced in the settings: it is an experimental
feature with a branch and a config file of its own, not something for an ordinary
user. Changes apply instantly through the same events as the context menu - there
is no separate OK/Cancel.

**Process:** after the build and the unit tests, a multi-agent adversarial review
was run (a workflow of 3 independent lenses - correctness, safety for Word, design
consistency - followed by a re-check of the findings by another agent). Confirmed
and fixed: the data loss in `SaveOrder` (described above; E2E caught the same bug
independently); the mutation of shared objects in `SetTabPinned`/`SetTabColor`
(above); the missing try/catch around `Close()` in the footer of `TabListPopup`
(the same pattern was already present in `ActivateRow` - fixed in both places);
the separator between the pinned and ordinary groups was not drawn in scrolling
mode (it was computed but never rendered into the offscreen overflow buffer -
fixed). Rejected (not confirmed on re-check) was the complaint about the native
`RadioButton`/`CheckBox` in `SettingsForm` - they are the first system controls in
the project (everything before them was custom paint), but Word hosts WinForms
controls in a process with a modern ComCtl32 v6 manifest, so the visual styles are
inherited without a call to `Application.EnableVisualStyles()` of our own; the
decision was confirmed by inspection on a live Word.

**Alternatives considered:** a separate gear button on the tab strip to enter the
settings (rejected - it breaks the pixel-accurate geometry of specification 2b for
no reason, and the footer of the all-tabs menu solves the same problem without the
risk); teleporting a tab back to its "historical" position when it is unpinned
(rejected - unpinning simply leaves the tab wherever the last stable sort put it,
the way Chrome and Edge behave - no separate memory of "where it used to be" is
needed); evicting old paths from `tab-order.cfg` (rejected - the same principle as
`TabColorStore`: the file is never cleaned, a deliberate trade of simplicity
against the unbounded but harmless growth of a small text file).

**Date:** 2026-07-25

## ADR-013. Do not move to a per-monitor "native" strip window now; soft text on non-system monitors is an accepted limitation

**Context:** on a monitor whose scale differs from the system one (in a multi-DPI
configuration) the tab text is slightly soft - DWM stretches the 96-DPI raster of
the system-aware WinForms strip (the mechanism is ADR-010, the limitation is O14).
Measured 2026-07-25 on three monitors SIMULTANEOUSLY (100 / 125 / 150 %): the
width of a text edge is 1 / 1 / 2 px (median; borderline noticeable), the strip
height is correct everywhere (43 -> 54 -> 65 px), and the monitor a document was
CREATED on does not affect sharpness (the strip repaints for the current monitor -
verified with identical tab contents). `GetDpiForWindow`: the Word window
(OpusApp) = 144, our host and strip windows = 96 - that is, the Word process is
per-monitor-aware while our strip is frozen system-aware when the WinForms handle
is created.

**Decision:** do NOT rewrite the strip's rendering onto a "native"
per-monitor-aware Win32 window at the MVP stage. Accept the soft text on
non-system monitors as a documented limitation (O14); defer the real fix until
after the MVP.

**Why (cost vs benefit):** the benefit is removing about 1 px of softening, almost
invisible and only on monitors whose scale differs from the system one in
multi-monitor configurations of mixed scale (on a single monitor, whichever is
primary, the text is crisp anyway). The cost is the most expensive and most
dangerous rework in the project: the strip runs inside the Word process, where an
uncaught exception out of a window procedure brings WINWORD down (0xE0434352)
along with the document; the code being rewritten is the most battle-tested we
have (37 try/catch blocks, a documented history of the ContextMenuStrip crash).
Realistically 14-20 weeks of one experienced Win32/.NET developer with a long tail
of stabilisation, and zero new features for the user. For a single non-programmer
developer steering the project towards a shippable MVP, that is a disproportionate
and project-endangering trade.

**Options considered (multi-agent assessment, see "Process"):**
- *A. Full replacement* - a raw per-monitor-aware Win32 window (our own WNDCLASS +
  WndProc) created under `SetThreadDpiAwarenessContext(PMv2)`, so that the window
  is born matching the DPI mode of the Word window and is not virtualised. The
  rendering (about 850-950 lines: tab geometry, GDI+ shapes, TextRenderer text,
  colours, layout, dragging) transfers almost as is; only the plumbing is
  rewritten (about 650-750 lines: creating the window, mouse/keyboard/focus,
  tooltips, the context menu, reacting to WM_DPICHANGED instead of swallowing it).
  It removes the blur COMPLETELY. The only sensible option if this is done at all.
  Direct2D/DirectWrite is NOT needed - it inflates the risk with no gain in
  sharpness.
- *B. A hybrid* (the old WinForms input plus a transparent per-monitor window on
  top for the picture) - a false economy: the only coherent variant (a
  click-through layered window) is temperamental inside someone else's process in
  a multi-DPI tree, gives two coordinate systems (hit-testing at 96, drawing at
  144 -> the click target drifts at the tab edges), and the "saved" input code is
  lost anyway if the trick fails - the worst case is more expensive than A.
  REJECTED as strictly dominated.
- *C. Cheap tricks* (supersampling, tweaking `SetDpiOverride`, "turning off
  stretching") - PLACEBO: the system physically hands a system-aware window a
  96-DPI surface only, and resolution cannot be added from inside it; tweaking the
  override breaks the height geometry that is already correct. Verified by reading
  the code.
- *C as the decision - ACCEPTED* - document it (this ADR + O14). Optionally half a
  day of cosmetics: a slightly heavier caption font - it reduces not the blur but
  how much it catches the eye; for a defect that is "borderline noticeable" that
  is often enough.

**If this is ever revisited, start with 1-3 days of reconnaissance** (a throwaway
prototype: `CreateWindowEx(WS_CHILD)` under a PMv2 context -> `SetParent` into
OpusApp -> a trivial WM_PAINT -> check `GetDpiForWindow(self)==144`, the sharpness
and the survival of moving the window between monitors of different scale - ON A
REAL machine and Office build of the user). While at it, compare the embedded
child window (WS_CHILD) with a separate top-level popup window over the tab zone:
the latter sidesteps the two most expensive risks of the WS_CHILD route -
delivering WM_DPICHANGED through `SetParent` into another process, and a child
window of another process stealing focus. The go/no-go gate rests on that
prototype (about 1 % of the cost of the full rework).

**Honest caveats:** "guaranteed crisp" would be too strong. The measurement was
made on one machine and one Office build; in "Optimize for compatibility" mode the
Office UI can be system-aware, and then the WM_DPICHANGED timing and the geometry
behave differently. After the fix a NEW class of bug appears that does not exist
today: while everything is drawn at 1:1 and stretched by the system, "drawing it
wrong" is impossible; when drawing at 1.5x, any mistake (a buffer in logical
rather than physical pixels; a leftover `SetDpiOverride`/`ScaleHeight` -> double
scaling, see the calls in `TabPaneManager.cs`; a two-pass ClearType trick on a
memory DC) is immediately visible - the blur can come back through the fault of
the new code. The rescale plumbing (`SetDpiOverride`,
`CurrentPanelHeight`/`ScaleHeight`) collapses into an identity after the move and
becomes removable dead code - it cannot be left in place, or the scaling doubles.

**Process:** the assessment was made by a multi-agent workflow (9 agents: a map of
how the strip's code is tied to WinForms -> an evaluation of the three options by
feasibility, effort and risk -> an adversarial re-check of the key assumption and
of the estimates -> synthesis). Confirmed against the code:
`SetThreadDpiHostingBehaviorSafe(MIXED)` is defined (`NativeWin32.cs`) but is
NEVER called; the PMv2 wrapper covers only the geometry of `UpdateLayout` and not
the creation of the window; both WndProcs swallow WM_DPICHANGED. See also ADR-010
(the strategy of drawing in your own DPI space) and limitation O14.

**Date:** 2026-07-25

## ADR-014. Do not install dependencies on the user's behalf; check in advance and explain, rather than automate

**Context:** the question came up of whether the installer could install what is
missing by itself (first of all .NET Framework 4.8). What is actually needed on
the user's machine was checked by reflection over the shipped
`TabsForWord.dll` - the list of assembly references is only `mscorlib`, `System`,
`System.Core`, `System.Drawing`, `System.Windows.Forms`:
- .NET Framework 4.8 - the only genuine dependency;
- desktop Word - obviously cannot be installed for the user (licensing);
- the Office PIAs - NOT needed at run time (NoPIA, `EmbedInteropTypes`, see
  ADR-002/005);
- the VSTO Runtime - NOT needed (a raw COM add-in, see ADR-002);
- the VC++ Redistributable - NOT needed (all the code is managed; every
  `DllImport` points at the system `user32`/`kernel32`/`gdi32`/`dwmapi`).

**Decision:** no auto-installation. Check, explain clearly, and give a link. State
the requirements UP FRONT (README, INSTALL_EN/RU, the installer banner) and not
only at the moment of refusal.

**Why (a verifiable fact for each point):**
1. **Administrator rights cannot be worked around.** Microsoft: ".NET Framework
   requires administrator privileges for installation". A per-user variant does
   not exist in principle: from Windows 8 onwards the .NET Framework is an OS
   component serviced through CBS/Windows Update; there is no `ALLUSERS=0`, no
   `/peruser`, no xcopy. Auto-installation would break the main property of the
   project - "no administrator rights" (ADR-006).
2. **The scenario is almost empty.** .NET 4.8 is built into Windows 10 from 1903
   (May 2019) and into every Windows 11. On such systems the .NET installer
   refuses to run at all ("is already a part of this operating system"). Of the
   supported Windows versions without 4.8 only LTSC 2019 / Server 2019/2016
   remain - precisely the closed managed environments where a script could not
   elevate anyway.
3. **Download an EXE and run it - that is a malware signature.** Defender's ASR
   rules and EDR products catch exactly the pattern "Invoke-WebRequest into %TEMP%
   -> Start-Process". Our DLL is already unsigned by Authenticode - that is the
   limit of acceptable "suspicious".
4. **Failing halfway is more dangerous than refusing.** .NET 4.x versions install
   in place, overwriting shared assemblies; an interrupted installation breaks
   OTHER .NET applications on the machine and can only be repaired with
   administrator rights (the Repair Tool or DISM). An installer able to create
   such a state but unable to fix it is unacceptable.
5. **A reboot is almost inevitable** in our scenario: a running WINWORD.EXE with a
   managed add-in is a ".NET 4 app in use", the documented trigger for a mandatory
   restart (codes 1641/3010).

**What was done instead (the stage after audit 3):**
- `install.ps1`: the .NET check was split into THREE cases instead of two -
  (1) the version was read and is below 4.8 -> fatal; (2) the key is missing AND
  there is no `mscorlib.dll` in `%WINDIR%\Microsoft.NET\Framework[64]\v4.0.30319`
  -> fatal (".NET 4 is not installed at all"); (3) the key is missing but the
  files are there -> a warning only (a failure to read the registry must not block
  a working machine).
- The refusal message: a human-readable version ("4.7.2 is installed") instead of
  `Release=461814`; a direct link plus copying it to the clipboard plus an offer to
  open the browser; a hint about Windows builds below 18362; an explicit note that
  administrator rights are needed for .NET but NOT for the add-in itself; a
  numbered instruction in the style of the "Word is running" block.
- Detecting Word was strengthened to THREE independent signals (the App Paths
  registry key -> a `winword.exe` file in the standard Office folders -> the
  `Word.Application` COM registration). Missing ALL THREE -> fatal: previously,
  installing on a machine without Word finished with "INSTALLATION COMPLETED
  SUCCESSFULLY". Three signals are needed so that the fatal refusal does not fire
  falsely: some Office installations have no App Paths key.
- `diagnose.ps1`: the same 528040 threshold, the same link and the same
  explanation in the problem text (so that the solution is right there in the
  report that was sent), and the same triple detection of Word (otherwise the
  report would contain a false "Word not found"). Version 1.0 -> 1.1.

**Rejected - lowering the target framework to net472/net462** to avoid the
prerequisite installation altogether: the gain would be Windows 10 builds
1803/1809 only (net472) and a few more obsolete ones (net462); on top of that
net462 is unreachable without rewriting the DPI layer (`Control.DeviceDpi` and
`OnDpiChangedAfterParent` require 4.7 or later - `TabStripControl.cs`), and
changing `TargetFrameworkAttribute` switches the set of WinForms compatibility
quirks, DPI scaling included - that is, a regression risk in the single most
fragile subsystem we have (O14, ADR-010/013), and one that cannot be verified
locally. Every supported Windows 10 (21H2/22H2) has 4.8 anyway. Conclusion: not
worth it; net472 remains a cheap fallback move.

**Honest caveats:** the fatal .NET refusal branches were checked by simulating the
logic, not on a real machine without .NET 4.8 (no such machine exists here) - see
O16. The triple detection of Word was verified on this machine: all three signals
fired independently (App Paths, the file on disk, the COM registration), so a
false refusal is unlikely; but the behaviour on a machine WITHOUT Word was not
verified live either.

**Process:** the research was done by a multi-agent workflow (3 agents: a full
breakdown of the run-time dependencies from the code and from the built DLL ->
Microsoft's facts about deploying .NET 4.8 (rights, reboots, silent-install
switches, return codes, what is built into which Windows) -> an evaluation of the
alternatives: lowering the target, bootstrapper patterns, UX options).

**Date:** 2026-07-26

## ADR-015. The native strip host becomes the main mode; the Custom Task Pane stays as the fallback (stage 28, v1.6.0)

**Context:** since 2026-07-24 the tab strip has been able to live directly inside
the Word window (a WinForms window of our own, `SetParent` into `OpusApp`,
anchored to `_WwF`) - without the Custom Task Pane title strip, which cannot be
removed through any official Office API (limitation O2). The mode was marked
experimental and had to be switched on by hand: `native-host.cfg` /
`native-tabs-on.cmd`. Over two days stages 13-27 were built on that mode (colours,
size, pinning, order, scrolling, per-monitor DPI, diagnostics, field fixes), while
the `master` branch stayed at checkpoint-12. In effect the "experiment" was the
product, and the "stable" branch was a stale build with no native host in it at
all.

**Decision:** make the native mode the default (`mode=native`, `reserve=1` when
the config is empty or missing), merge the `experiment/native-tab-host` branch
into `master`, and collapse the two packages into one. The classic CTP mode is NOT
removed: it stays as (a) an explicit user choice through `mode=ctp`
(installer\Classic-mode.cmd), (b) the automatic emergency contour inside
`NativeTabHostManager` (5 failed synchronisations in a row -> a fallback to CTP
within the current Word session), and (c) insurance against an unforeseen failure
to read the settings.

**Why now (verifiable facts, not feelings):**
1. **The field.** The package with the native mode enabled ran on TWO other
   people's machines (machines A and B, Win10 Pro 25H2, Word ProPlus2024 x64,
   150 %): both installations succeeded, the strip worked, and Word never crashed.
   The single genuine defect (a locator race against a half-built window) was found
   from the logs and fixed in v1.5.1.
2. **Multi-DPI is closed.** A run on three monitors at once (100/125/150 %),
   13/13 PASS: the strip reports per-monitor DPI, the height scales 43 -> 54 ->
   65 px, moving back restores the size, and the exit is clean (section 11 item 6
   of the experiment).
3. **The test contour.** 53/53 automated tests and 11/11 E2E on a real Word,
   including the dedicated runs for reserve mode, side panes, scrolling, colours,
   sizes, settings and pinning - and all of them run in the native mode
   specifically.
4. **The emergency contour exists and is verified by code and tests,** which means
   the worst outcome of the native host failing is not "a broken Word" but the
   look of the pre-v1.6.0 versions.
5. **This mode is why the project was built:** O2 (the roughly 35 px CTP strip
   above the tabs) was the main complaint about the appearance, and in the native
   mode it is gone.

**What changes in behaviour:**
- A clean installation with no config: the strip is embedded in the Word window
  and the document area moves down by exactly the height of the strip
  (`reserve=1`).
- The `reserve` default changed from 0 to 1: the overlay (the strip on top of the
  document, covering the ruler) can no longer switch itself on - only an explicit
  `reserve=0` does that (the internal `scripts\mode-overlay.cmd`).
- The mode is always written to the log, the classic one included ("set by
  config/env"), so that a field report shows that CTP was enabled deliberately.
- The name of the mode in the code: `TabHostMode.NativeExperimental` ->
  `TabHostMode.Native`; the string `mode=nativeexperimental` in old configs is
  still read.

**Risks and how they are covered:**
- *The Word window class names (`_WwF`/`_WwG`/`_WwB`) are undocumented.* If they
  change after an Office update, the fallback heuristic of the locator kicks in,
  then the emergency fallback to CTP; Word and the add-in keep working. This is a
  permanent risk - it is not eliminated but insured against (O17).
- *The locator race fix (v1.5.1) is not yet confirmed in the field.* Confirmation
  will come with the Diagnostics reports for v1.6.0; the automatic summary in
  Diagnostics already contains a direct check for "the strip sits at the very top
  of the window" and for "fell back to the classic mode".
- *The Office dark theme:* the strip background will pick up a dark tone while the
  tabs themselves stay light (the specification 2b palette). Limitation O18, out
  of scope for this stage.

**Rejected - removing the CTP mode entirely** (it would simplify the code by about
one interface and one branch in `Connect`): it is cheap to keep (the same
`TabStripControl`, and the separate manager is already written and covered by
tests), while it serves both as a fallback contour if the internals of Word change
and as the only working appearance on a machine where the native host for some
reason fails to attach. The balance of cost against insurance favours keeping it.

**Rejected - waiting for field reports on v1.5.1 first and switching afterwards:**
distributing v1.6.0 gives the same confirmation, but in one step and without the
instruction "run native-tabs-on.cmd", on which users had already got lost.

**Date:** 2026-07-26

## ADR-016. The interface language: a table in the code, detection from Word's language (stage 29, v1.7.0)

**Context:** the decision was made to make the add-in usable by people from
different language regions; the first language besides Russian is English.
Measuring the volume showed that the interface itself has only about 30 captions
(the tab menu, the tooltips, the settings window, the all-tabs popup), whereas the
bulk of the Russian text lives in the installer (about 280 lines), the
instructions (about 180) and the code comments (980).

**Decision (three parts):**

1. **Strings are stored in a compiled table, `Strings.cs` - not .resx and not
   satellite assemblies.** The whole installation rests on "one DLL, per-user, no
   administrator rights": `install.ps1`, `repair.ps1`, `diagnose.ps1` and
   `package.ps1` all know exactly one file. Satellites would add a
   `ru\TabsForWord.resources.dll` folder that all four scripts would have to be
   told about, and losing it would mean the interface silently moving to another
   language. At this volume of strings a table is cheaper; the captions are static
   properties, so a missing translation is technically impossible (there are no
   keys that can drift apart), and an automated test additionally catches empty
   and untranslated ones.

2. **The language is detected from WORD's interface language, not Windows'**
   (`LanguageSettings.LanguageID[msoLanguageIDUI]`). A person may well have an
   English Windows and a Russian Word - the tabs then have to be Russian, whereas
   the system locale would say the opposite. The order is: `lang=` in language.cfg
   -> the Word language -> the Windows locale. Anything that is not Russian
   resolves to English: English is the neutral language and the fallback for
   anything untranslated. Verified on a live Word: LCID 1049 is detected and the
   menu reads in Russian; with `lang=en` the same menu reads in English
   (e2e-language.ps1).

3. **The manual choice lives in the settings window** (three options: "Same as
   Word", English, Russian). The language names are written in their own
   languages: someone who has accidentally switched Word to a language they do not
   know must be able to find the way back. Switching applies at once in every
   window: the context menu is built once and cached, so the cache is dropped
   (`RefreshLocalizedUi`), and the remaining captions are computed while drawing.

**What deliberately stays in its original language:**
- **The logs and the Diagnostics report are always English.** A report from
  someone else's machine is read by the developer; keeping the log strings in one
  language is what makes searching and grouping in `diagnose.ps1` possible at all
  (the parser currently looks for English substrings such as "falling back to
  CustomTaskPane").
- **The name in the "COM Add-ins" list is written by the installer** from the
  Windows locale: Word is closed at the moment of installation, so there is nobody
  to ask. "MVP" was dropped from the name.

**Rejected - translating strings on the fly by calling the Word API each time**
(reading Word's language before every repaint): the Office interface language does
not change without restarting the application, and extra COM calls inside a
painting loop are a direct route to lag and to exceptions in WndProc.

**Rejected - a separate language pack or translation file next to the DLL:** the
same reason as for satellites (an extra entity in the installation and in the
diagnostics), plus the user editing the text would become a source of
incomprehensible reports.

**How to add another language:** an entry in `UiLang`, a branch in
`UiLanguage.FromLcid`, an optional parameter in `Strings.T` (untranslated strings
stay English by themselves) and an option in the settings window. Nothing else in
the project knows anything about languages.

**Date:** 2026-07-27

---

## ADR-017. "No document area" is split into two answers: a failure and the File menu (stage 34, v1.7.1)

**Context:** a field report (machine C, 2026-07-28, Word 16.0.20228, Windows 10
22H2, 100 % scaling) came back with a clean install, zero ERROR lines and a user
who said everything worked - yet the Diagnostics summary raised a problem, because
the add-in had saved a window-tree dump. The dump explained itself: a visible
window of class `FullpageUIHost` covering the whole client area, and the entire
document subtree (`_WwF`, `_WwB`, `_WwG`, every `MsoCommandBarDock` and
`MsoWorkPane`) hidden. That is what Word looks like while the File menu
(Backstage) is open - the user had spent about 84 seconds there.

Everything the add-in did was correct: the locator found no document area, the
strip hid itself (had it stayed, it would have floated over the File screen), and
it came back about 1.4 s after the menu was closed. The cost was in the
bookkeeping around it:

1. Ten WARN lines and a `WindowTree-*.txt` dump for an event that is not a
   failure. The dump is our one reliable signal of a REAL breakage in the field
   ("the classes changed, the locator is lost"); if an ordinary trip to the File
   menu produces one, the signal is worth nothing.
2. Worse, `UpdateLayout` returned false, which the manager counts as a
   native-mechanics failure. In this report two windows were open, so the counter
   never moved (it only counts syncs where EVERY window failed) - but with a
   single Word window open, five polls inside the File menu would have dropped the
   add-in into the CTP fallback, turning the strip into the classic task pane in
   the middle of normal work. Nobody had hit it yet; it was found by reading the
   log, not from a complaint.

**Decision:** "the document area was not found" stops being one answer and becomes
two. `WordNativeWindowLocator.IsFullPageUiActive` recognises the state by the
window that causes it (a visible `FullpageUIHost` covering at least 90 % of the
client area), and the host distinguishes `ZoneState.NotFound` from
`ZoneState.FullPageUi`. In the full-page case: one INFO line per episode instead
of a stream of WARNs, no failure counted, no dump, and `UpdateLayout` returns true
- the same "a benign reason to do nothing" as a minimised window. The strip is
still hidden: that part was right all along.

**Why a window class and not "Word says Backstage is open":** the object model has
no such property. The Ribbon `onShow`/`onHide` callbacks belong to custom-UI XML,
which this add-in deliberately does not have (ADR-002: no ribbon XML, no VSTO), and
polling something like `Application.Visible` says nothing about Backstage. The
window that covers the document IS the state, and it is the same signal the
locator already works in - one mechanism, not two.

**Why 90 % and not "exactly the client area":** the coverage is compared against
the intersection with the client rect, so a host window that Word leaves slightly
larger than the client area, or a one-pixel rounding at a fractional scale, still
counts. A leftover 100x100 `FullpageUIHost` (Word keeps such windows around
between uses) does not.

**Deliberately NOT excused:** a half-built window (the machine A race, ADR fixed in
v1.5.1) has no full-page host, so it stays a real "not found" with its counters,
its retries and its dump. A unit test asserts exactly that - otherwise this change
would have silenced the very failure the dump exists for.

**In the Diagnostics (1.5):** dumps containing `"FullpageUIHost" visible` are
counted separately and do not raise a problem. Version 1.7.1 does not write them
at all, but builds 1.6.0-1.7.0 are already in people's hands and their reports
have to read correctly too.

**Date:** 2026-07-28
