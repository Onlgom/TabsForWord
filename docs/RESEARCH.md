# RESEARCH.md - the research behind the design (stage 1)

Written 2026-07-23. The research was done against the official Microsoft
documentation (learn.microsoft.com) with an adversarial re-check of the key
claims; the steps critical for building and registration were additionally
**verified in practice on the development machine**.

## The main question: are "real" tabs inside one physical Word window possible?

**Answer: NO - not by any officially supported means.**

The chain of evidence:

1. Word 2013 and later (Microsoft 365 included) is strictly SDI: every document
   lives in its own top-level window (window class OpusApp) with its own ribbon.
   MDI mode was "removed entirely" as a feature (the official Office 2013 list of
   changes: *"Users will no longer be able to enter MDI Mode… Replacement: N/A"*).
2. `Application.ShowWindowsInTaskbar` - the historical switch - is officially
   deprecated ("not intended to be used in your code") and does nothing in a
   modern Word.
3. The Word object model has no member that could merge windows, make a document
   window a child, or provide a tab container.
4. The commercial Office Tab achieves the effect with undocumented Win32
   manipulation (hooks plus hiding and moving OpusApp windows and injecting its
   own tab bar). That is not an official API route: no suitable API exists.
5. Microsoft itself documents the dangers of reparenting foreign top-level
   windows: SetParent does not change the WS_CHILD/WS_POPUP styles, it breaks
   across DPI modes (ERROR_INVALID_STATE) and it ties the input queues of the
   threads together (Raymond Chen, devblogs). Word additionally relies on "one
   window = one ribbon + one chain of modal dialogs"; any monthly Click-to-Run
   update could break the whole thing.

**Conclusion:** build the safe MVP - a tab strip shown in sync in every Word
window, where clicking a tab activates the corresponding window. The limitation
is documented honestly (see KNOWN_ISSUES.md).

Note from later stages: the tab strip did eventually move INSIDE the Word window
(v1.6.0), but that is a strip of our own hosted in Word's window - the document
windows themselves are still separate. The conclusion above still holds.

## The stack: VSTO versus a classic COM add-in

The machine has no Visual Studio, no MSBuild, no VSTO SDK and no VSTO Runtime
(docs/ENVIRONMENT.md). A classic VSTO project **cannot be built here at all**,
and installing the full VS means gigabytes and manual confirmations.

**Choice: a classic COM add-in (IDTExtensibility2) without VSTO.** Verified:

- COM add-ins are fully supported in a modern desktop Word M365 (2025/2026); the
  deprecation announced applies to the new Outlook only. The documentation of the
  registry mechanism was updated in 2026. Add-ins of exactly this kind already
  work on this machine (OneNote.WordAddinTakeNotesService, Adobe PDFMaker).
  **Re-check verdict: CONFIRMED.**
- The `IDTExtensibility2` interface (GUID B65AD801-ABAF-11D0-BB8B-00A0C90F2744,
  five methods OnConnection/OnDisconnection/OnAddInsUpdate/OnStartupComplete/
  OnBeginShutdown) can be declared by hand with `[ComImport]` - no reference to
  Extensibility.dll is needed.
- Word finds the add-in through the key
  `HKCU\Software\Microsoft\Office\Word\Addins\<ProgId>` (a path WITHOUT a version
  number; the values FriendlyName, Description, LoadBehavior=3). No administrator
  rights are needed. `Manifest` is for VSTO only and must not be created.
- Word loads the .NET class through standard COM activation:
  `InprocServer32 = mscoree.dll` plus Class/Assembly/RuntimeVersion=v4.0.30319/
  CodeBase. .NET Framework 4.8.1 is an in-place update of CLR 4.0, so it works.
- If OnConnection throws, Word flips LoadBehavior 3 -> 2 ("Not loaded…") - a
  diagnostic sign of a crash during loading; every entry point must be inside a
  try/catch.
- A non-shimmed managed add-in loads into the default AppDomain (no VSTO
  isolation) - acceptable for the MVP; the DLL is file-locked while Word runs, so
  reinstalling requires Word to be closed, background preload included.
- The Startup Boost trap (since mid-2025): the scheduler can keep a preloaded
  WINWORD.EXE paused - check for background Word processes before replacing the DLL.

## A Custom Task Pane from a COM add-in (without VSTO)

- The Connect class additionally implements `ICustomTaskPaneConsumer`
  (GUID 000C033E-0000-0000-C000-000000000046, the single method
  CTPFactoryAvailable); Office calls it after connecting and passes an
  `ICTPFactory` (GUID 000C033D-…) - the factory is cached for the whole lifetime.
- `ICTPFactory.CreateCTP(progId, title, window)` takes the **ProgID of a
  registered** COM class for the content. An ordinary WinForms UserControl with
  [ComVisible(true)]/[Guid]/[ProgId], registered the same way (mscoree +
  CodeBase), is enough - the special ActiveX keys (Control/MiscStatus) are not
  required (community-confirmed; adding them does no harm). The error "Unable to
  create specified ActiveX control" means the wrong registration bitness, a
  missing CodeBase or a missing ComVisible.
- **Word is SDI: a CTP belongs to exactly one document window.** A strip in every
  window therefore needs a manager: create a CTP for every Word.Window (passing
  the Window to CreateCTP), keyed by `Window.Hwnd`. **Re-check verdict: CONFIRMED.**
- `msoCTPDockPositionTop` (=1) is supported in Word; the order is: create the CTP
  -> DockPosition=Top -> DockPositionRestrict=NoChange -> Height -> Visible=true.
  Do not change the geometry inside CTP event handlers (COMException).
  **Re-check verdict: CONFIRMED.**
- **The minimum CTP height is enforced** (documented as "depends on several
  factors"); values below the minimum are silently increased. A thin 32-50 px
  strip may be unreachable through a CTP (realistically 70-120 px) - measure it
  on the target machine and write the actual value down.
- When a window closes, Word destroys its CTP itself; our RCW goes "dead" (a
  COMException on any access) and there is NO event about it - we need our own
  tracking and cleanup (a probe inside try/catch plus Marshal.ReleaseComObject).
- Protected View windows get no CTP (add-ins do not run there). The add-in still
  lists a PV document as a grey tab in the strips of ordinary windows; after
  "Enable Editing" it becomes an ordinary tab (KNOWN_ISSUES O3).

## Word object model events

The conclusions that shaped the tab model:

- Drive the strip from `Application.Windows` (one window = one tab; a document
  can have 2+ windows through "New Window" - the captions are "Name:1",
  "Name:2") plus `Application.ProtectedViewWindows`.
- `DocumentBeforeClose` is only a hint: the user can still cancel the close in
  the save dialog AFTER the event, and with File -> Exit the event can arrive
  AFTER the dialog. `DocumentAfterClose` does not exist. Solution: the event
  triggers a deferred (200-500 ms) idempotent Reconcile() that compares the live
  collections with the tab model.
- `DocumentOpen`/`NewDocument` do not fire for documents already open when the
  add-in loads, so an initial scan is mandatory. `DocumentChange` is a convenient
  catch-all (creation, opening, active document change, an actual close).
- Save As: Word has NO "after save" event. `DocumentBeforeSave(SaveAsUI)` reports
  the intent; the outcome is tracked by comparing FullName/Saved after the event
  plus a slow background poll (1-2 s) - which also catches renames that arrive
  without events. The change indicator is `Document.Saved` (false = unsaved
  changes).
- Closing a tab: `Document.Close(wdPromptToSaveChanges)` - the standard Word
  dialog. A cancel by the user is Word error 4198 -> a `COMException`, which we
  catch, leaving the tab in place. NEVER call wdDoNotSaveChanges.
- Switching: `Window.Activate()` plus P/Invoke `SetForegroundWindow(Window.Hwnd)`
  (the click happens in our process, so we are allowed to set the foreground); if
  the window is minimised, restore it first. (Since v1.7.3 the restore is
  WM_SYSCOMMAND/SC_RESTORE rather than `WindowState = Normal`: the COM route
  turned a window that had been maximised into an ordinary one - ADR-018.)
- Protected View: its own family of six events (ProtectedViewWindowOpen/…/
  BeforeClose with a CloseReason; wdProtectedViewCloseEdit is followed by
  DocumentOpen - convert the tab rather than removing it). PV documents are not
  part of Application.Documents (to be confirmed in stage 8).
- Subscribing from C#: hook the events on a `Word.Application` stored in a field;
  the cast to `ApplicationEvents4_Event` is needed only for the `Quit` event
  (a name collision with the method). **Keep the delegates in class fields** -
  otherwise the GC silently kills the COM subscription (the documented "events
  stop firing" problem). Unsubscribe in OnDisconnection.

## Building without Visual Studio - verified in practice on this machine

- `dotnet build` (SDK 8.0.423, user scope; 8.0.425 since 2026-09-26, see
  ENVIRONMENT.md) builds an SDK-style `net48` project
  with `UseWindowsForms=true` without VS; the reference assemblies come from the
  Microsoft.NETFramework.ReferenceAssemblies package automatically.
  **Verdict: CONFIRMED, and a local build succeeded (0 errors).**
- Word interop: the NuGet `Microsoft.Office.Interop.Word` 15.0.4797.1004 (a
  repackaging of the genuine Microsoft-signed PIA; the nuget.org owner is a third
  party). It works against Word 16.x. The fallback without a third-party package
  is the genuine PIA in the GAC:
  `C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c`.
  EmbedInteropTypes through a PackageReference does not work by default - the
  official workaround (a Target with ReferencePath/EmbedInteropTypes) was
  **verified: the resulting DLL does not depend on the interop DLL**.
- office.dll (core) is NOT on NuGet from Microsoft - the interfaces
  ICustomTaskPaneConsumer / ICTPFactory / IDTExtensibility2 / IRibbonExtensibility
  are declared by hand ([ComImport] plus the GUID; a route proven in production by
  Excel-DNA). A compiling Connect.cs was produced during the research.
- Registration without administrator rights was **verified end to end on this
  machine**: `RegAsm /regfile /codebase` -> replace `[HKEY_CLASSES_ROOT\` with
  `[HKEY_CURRENT_USER\Software\Classes\` -> import with the **64-bit** reg.exe ->
  `New-Object -ComObject` successfully creates the class. **Verdict: CONFIRMED.**
- **A critical bitness trap:** `HKCU\Software\Classes\CLSID` is
  WOW64-**redirected** (unlike HKCU\Software\Classes in general). The import must
  be done by a 64-bit process, otherwise the keys land in Wow6432Node and 64-bit
  Word never sees them (verified by querying /reg:64 versus /reg:32). The build is
  AnyCPU.
- `<COMReference>` does not work under dotnet build (MSB4803) - not used, and not
  needed.

## Deployment and removal (per user, no administrator rights)

- Installation is five steps of a single PowerShell script: (1) check that Word
  is closed, background instances included, and do NOT kill it; (2) copy the DLL
  into `%LOCALAPPDATA%\TabsForWord`; (3) `Unblock-File` on everything (the
  Mark-of-the-Web from a downloaded ZIP otherwise blocks loading through
  CodeBase, and editing winword.exe.config needs administrator rights);
  (4) the HKCU\Software\Classes keys (CLSID + ProgId, the 64-bit view); (5) the
  Word\Addins key plus a cleanup of `Resiliency\DisabledItems`.
- Word reads the add-in list at startup, so Word has to be restarted after the
  installation. Removal requires Word to be closed (the DLL is file-locked).
- Repair mode: LoadBehavior=3 plus a DisabledItems cleanup - that covers both
  ways in which Word "bans" a crashed add-in. (For a non-shimmed .NET add-in,
  mscoree.dll itself can end up in DisabledItems - which blocks EVERY .NET
  add-in.)
- A signature is not required: RequireAddinSig is off by default (no policy was
  found on this machine). Word started "as administrator" may ignore the HKCU
  registration - to be documented in the installation guide.
- The scripts are wrapped in .cmd files calling
  `powershell -NoProfile -ExecutionPolicy Bypass -File`, so that a double click
  works under any ExecutionPolicy.

## Word AutoSave and window switching (2026-09-26)

**The question:** the user reported that with AutoSave on, every edit got saved as
soon as they switched to another document (the unsaved dot disappears, and the
switch away from such a tab lags). They meant the "AutoSave" toggle in the Word
title bar, which applies to OneDrive/SharePoint files. Is the add-in causing the
save?

**The code:** no. The add-in never calls Save/SaveAs on a document and never
cancels a save; `DocumentBeforeSave` only schedules a reconcile.

**The experiment** (Word M365 x64 16.0.20326, v1.7.2): two temporary test
documents were created in the user's OneDrive with their permission and deleted
afterwards; Word opened them as cloud documents with AutoSave on. Each
repetition: type 3 characters, wait 250 ms, switch to the other document, poll
`Document.Saved`.

| How the switch was made | Saved after the switch |
|---|---|
| no switch (baseline) | 1 of 4 repetitions saved after 7.4 s; 3 of 4 not saved within 12 s |
| a tab click (the add-in) | 219-312 ms (4/4) |
| Alt+Tab | 218-281 ms (4/4); the target window comes to the front only after the save |
| SetForegroundWindow from outside Word | 0-15 ms (3/3) |
| Word's own `Window.Activate` (COM, from outside) | 125 ms (3/3) |
| Alt+Tab with the add-in DISCONNECTED (`COMAddIns(...).Connect = False`) | 218-234 ms (3/3) |

**Conclusion:** Word's AutoSave saves a changed cloud document as soon as its
window loses activation, synchronously - whatever does the switching, with or
without the add-in; the next window comes forward only after that save, which is
the lag the user notices. Left alone, the same edit was usually not saved within
12 s. The add-in cannot prevent it without turning AutoSave off, and it must not
change how the user's documents are saved behind their back. Documented as
limitation O21 with what the user can do: turn the "AutoSave" toggle off for the
document (Word remembers it per file) or for all cloud files (File > Options >
Save), and save with Ctrl+S. AutoRecover is a different feature: it keeps
recovery copies and never overwrites the document file.

## What was researched -> what question arose -> the decision (summary)

| Question | Decision | Reason |
|---|---|---|
| Real tabs in one window? | No; synchronised strips in every window | MDI removed, no API, reparenting dangerous and unsupportable |
| VSTO or a COM add-in? | A COM add-in (IDTExtensibility2) | VSTO cannot be built here and needs the Runtime; the COM route was verified on the machine |
| The UI host of the strip? | A Custom Task Pane, msoCTPDockPositionTop (superseded in v1.6.0 by our own window inside the Word window, ADR-015; the CTP stays as the fallback) | The official mechanism; top docking confirmed |
| The UI framework? | A WinForms UserControl | Required by CreateCTP (the ProgID of a COM class); WinForms is simpler for ActiveX hosting |
| The runtime? | .NET Framework 4.8 (net48) | Already on every Windows 10/11; CLR v4.0.30319 is what the mscoree registration requires |
| The build? | dotnet 8 SDK (user scope) | Verified: net48 + WinForms builds without VS |
| Interop? | The NuGet Word PIA (embedded) plus hand-written ComImport interfaces | There is no official office.dll on NuGet; NoPIA removes the runtime dependency |
| Registration / installation? | A PowerShell script, HKCU, 64-bit registry | No administrator rights; verified end to end |
| Does the add-in cause the AutoSave save on a switch? (2026-09-26) | No; Word does it on any deactivation - documented (O21), AutoSave is left alone | Measured with and without the add-in: the same ~0.2-0.3 s save |

## Sources (the main ones)

- https://learn.microsoft.com/en-us/previous-versions/office/office-2013-resource-kit/cc178954(v=office.15) - MDI removal
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._application.showwindowsintaskbar?view=word-pia
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setparent
- https://devblogs.microsoft.com/oldnewthing/20130412-00/?p=4683 - the dangers of cross-thread reparenting
- https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins - the Addins key, LoadBehavior
- https://learn.microsoft.com/en-us/dotnet/api/extensibility.idtextensibility2 - IDTExtensibility2
- https://learn.microsoft.com/en-us/office/vba/api/office.ictpfactory.createctp - CreateCTP
- https://learn.microsoft.com/en-us/office/vba/api/office.msoctpdockposition - DockPositionTop
- https://learn.microsoft.com/en-us/visualstudio/vsto/custom-task-panes - CTPs in a multi-window Word
- https://learn.microsoft.com/en-us/archive/blogs/andreww/the-anomalous-behavior-of-custom-task-panes-in-word-and-infopath
- https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforeclose
- https://wordmvp.com/FAQs/MacrosVBA/PseudoBeforeClose.htm - the unreliability of BeforeClose
- https://learn.microsoft.com/en-us/office/vba/api/word.document.close(method) - error 4198
- https://learn.microsoft.com/en-us/office/vba/api/word.application.protectedviewwindows
- https://learn.microsoft.com/en-us/archive/blogs/mstehle/oom-net-part-1-introduction-and-why-events-stop-firing
- https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/reference-assemblies - building without a targeting pack
- https://learn.microsoft.com/en-us/dotnet/framework/tools/regasm-exe-assembly-registration-tool - /regfile, /codebase
- https://learn.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys - WOW64: HKCU\Classes\CLSID is redirected
- https://learn.microsoft.com/en-us/dotnet/framework/configure-apps/file-schema/runtime/loadfromremotesources-element - MotW
- https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/admin/miscellaneous/new-startup-boost-tasks-windows-task-scheduler
- https://learn.microsoft.com/en-us/archive/blogs/vsod/troubleshooting-com-add-in-load-failures
