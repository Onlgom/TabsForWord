# TabsForWord — Tabs for Microsoft Word

A COM add-in for desktop Microsoft Word on Windows that adds a tab bar for the
open documents: click a tab to switch to that document, the way tabs work in a
browser.

*Читать по-русски: [README.ru.md](README.ru.md).*

Since version 1.6.0 the tab bar lives **inside the Word window**, right below
the ribbon, with no Custom Task Pane title strip above it (decision
[ADR-015](docs/DECISIONS.md); how it works:
[docs/NATIVE_TAB_HOST.md](docs/NATIVE_TAB_HOST.md)). The classic Custom Task
Pane mode is kept as a fallback: it can be switched on with `Classic-mode.cmd`
from the installation folder, and it switches itself on if the in-window host
ever fails.

Since version 1.7.0 the interface follows **Word's own UI language** — Russian
Word gets Russian tabs, everything else gets English — and the language can be
picked by hand in Settings (ADR-016).

## Requirements

- Windows 10 version 1903 (May 2019) or newer, or Windows 11 — these already
  include .NET Framework 4.8.
- Desktop Microsoft Word for Windows (Microsoft 365, Word 2016 or newer).
  Not Word Online, not Word for macOS, not Word on mobile.
- No administrator rights: the add-in is installed for the current user only.

## What it does

- One tab per open document, in every Word window, kept in sync.
- Close, close others, pin, reorder by dragging, colour a tab, choose the tab
  size, open the file's folder.
- A dot marks unsaved changes; a tooltip shows the full path.
- Overflow: tabs shrink, then a compact mode, then ‹ › scrolling.
- A ▾ menu listing every tab with a search box, and a Settings window.
- Ctrl+Tab / Ctrl+Shift+Tab to switch tabs (can be turned off).
- Per-monitor DPI: the bar scales with the monitor the window is on.

## For users

- Install: [docs/INSTALL_EN.md](docs/INSTALL_EN.md) · по-русски:
  [docs/INSTALL_RU.md](docs/INSTALL_RU.md)
- Uninstall: [docs/UNINSTALL_EN.md](docs/UNINSTALL_EN.md) · по-русски:
  [docs/UNINSTALL_RU.md](docs/UNINSTALL_RU.md)
- Ready-to-install files: the [Releases](../../releases) page
- Known limitations, honestly written down: [docs/KNOWN_ISSUES.md](docs/KNOWN_ISSUES.md)

## For developers

- Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- Decision log (why things are the way they are): [docs/DECISIONS.md](docs/DECISIONS.md)
- The in-window tab host in detail: [docs/NATIVE_TAB_HOST.md](docs/NATIVE_TAB_HOST.md)
- Research behind the approach: [docs/RESEARCH.md](docs/RESEARCH.md)
- Testing rules and the required test matrix: [docs/TESTING.md](docs/TESTING.md)
- Build: `scripts/build-release.ps1` (needs the .NET SDK; Visual Studio is not
  required — see [docs/ENVIRONMENT.md](docs/ENVIRONMENT.md))

Everything under `docs/` is in English.

## Licence

MIT — see [LICENSE](LICENSE). Use it, change it, ship it in your own product;
just keep the copyright line. It comes with no warranty of any kind: the add-in
runs inside Word next to your documents, and you use it at your own risk.
