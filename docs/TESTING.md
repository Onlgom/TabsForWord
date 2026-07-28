# TESTING.md - the testing rules of this project

Two levels of checks. The rules are MANDATORY for every change: they replace
the intuition of "what should I run" with a fixed matrix.

## Level 1: unit tests, no Word involved

`powershell -ExecutionPolicy Bypass -File scripts\test.ps1` - seconds.
Pure logic: the tab model, parsing of the settings files (colours, size,
native-host.cfg, language), layout geometry (ComputePanelRect, DPI), the window
locator, construction and painting of the control, live SetParent hosting.
Run after EVERY build, no exceptions.

## Level 2: end-to-end on a real Word (tests/manual/)

Requirements: Word closed, and do not touch the mouse or the keyboard while a
script runs. Every script installs the fresh Release DLL itself and protects the
user's settings (it backs up and restores the cfg files, creates its own
temporary documents, closes them without dialogs).
The restore has to be exact: if the user had NO native-host.cfg, the script
deletes its own. This is not hypothetical: three scripts once failed to do it,
the add-in was left in the overlay mode with the strip lying on the ruler, and
the person using the machine found it rather than the tests.

| Script | What it checks | ~Time |
|---|---|---|
| e2e-smoke.ps1 | SMOKE: loading, the strip paints, a click, the menu, collapsing, a clean exit | ~30 s |
| e2e-default-mode.ps1 | mode defaults: NO config => native+reserve; mode=ctp => the classic host | ~1.5 min |
| e2e-backstage.ps1 | the File menu (full-page UI): the strip hides and comes back, no dump, no CTP fallback | ~1 min |
| e2e-language.ps1 | interface language: lang=en gives an English menu (read back with GetMenuStringW), lang=auto takes the language from Word | ~1.5 min |
| e2e-tab-colors.ps1 | colours: the menu, the pastel by pixels, the file, reset, restart | ~2 min |
| e2e-tab-colors-dialog.ps1 | the system "More colors..." dialog | ~1 min |
| e2e-tab-size.ps1 | tab size: the menu, the height, a restart | ~2 min |
| e2e-dpi-monitors.ps1 | moving the window between monitors with different scales | ~1.5 min |
| e2e-side-panes.ps1 | Word side panes are not covered | ~1 min |
| e2e-overflow-scroll.ps1 | scrolling with 24 tabs: clipping, the arrows, the touchpad | ~1.5 min |
| e2e-reserve-mode.ps1 | reserve: the document area is pushed down, resize, the ribbon | ~2 min |
| e2e-overlay-interaction.ps1 | basic clicks, the ribbon, Ctrl+Tab (overlay) | ~2 min |
| e2e-overlay-buttons-pv.ps1 | the buttons plus Protected View | ~2 min |
| e2e-settings-pins-order.ps1 | pinning, stored order, the settings window | ~3 min |
| probe-side-panes.ps1 | (not a test) a probe of the Word window tree | - |

## The matrix: what to run for which change

| Kind of change | Mandatory |
|---|---|
| Pure logic, no UI (parsing, stores, the model) | level 1 |
| Painting / strip geometry (paint, sizes, buttons) | level 1 + SMOKE |
| Menus and interaction | level 1 + SMOKE + the feature's own E2E |
| A particular feature (colours / size / DPI / side panes / scrolling) | level 1 + SMOKE + ITS own E2E |
| Host geometry, reserve, the anchor, resize | level 1 + SMOKE + e2e-reserve-mode |
| Multi-monitor / DPI | level 1 + SMOKE + e2e-dpi-monitors |
| Host mode selection, settings defaults, the emergency fallback | level 1 + e2e-default-mode + e2e-backstage |
| Interface captions, new menu items, language | level 1 + e2e-language |
| Installer / registration / diagnostics | run install.ps1 + diagnose.ps1 + check the registry |
| BEFORE REBUILDING THE PACKAGE FOR DISTRIBUTION | the full E2E set (a dress rehearsal) |

The key rule (born from the 0xE0434352 crash): ANY change that touches painting
code or a window procedure is checked on a live Word with at least the smoke
test - a mistake in that layer can bring WINWORD down entirely.

## Other people's machines in the documents

Diagnostic reports arrive from other people's computers. In the repository
documents such runs are recorded as **machine A**, **machine B** and so on - the
computer name from the report (and the user name, the paths and the document
names) is never carried over. The reason is simple: a person sent the report so
that the add-in would be fixed, not so that their configuration would sit in
public. The technical facts (Windows and Word builds, DPI, measurements, the
number of windows) do have to be recorded - the decisions rest on them.

## Honesty and flakes

- Never claim a test passed if it was not actually run
  (docs/TEST_RESULTS.md: PASS / FAIL / NOT TESTED / BLOCKED).
- The E2E tests use real input, so a single step can fail on timing.
  Policy: one repeat of the whole script; two failures in a row count as a bug
  to be investigated, not something to re-run until it goes green.
- Wait for menus and dialogs by polling (Wait-Menu by the Word process id), not
  with fixed pauses; write new scripts the same way.
- Test documents are created by the tests themselves; closing a user document
  without the standard Word dialog is forbidden - losing somebody's document is
  the one failure this project does not accept.
