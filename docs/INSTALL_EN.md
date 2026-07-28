# Installing "Tabs for Microsoft Word" — a plain-language guide

This add-in gives Microsoft Word a tab bar, like the one in a browser: every
open document appears as a tab at the top of every Word window, a click switches
between them, the × closes a document and the "+" button creates a new one.

No special knowledge is needed. No administrator rights are needed.

(Русская версия этой инструкции: [INSTALL_RU.md](INSTALL_RU.md).)

## What you need

- **Windows 10 version 1903 (May 2019) or newer, or Windows 11.**
  These already include everything required (.NET Framework 4.8) — there is
  nothing extra to install.
- **Desktop Microsoft Word for Windows** (Microsoft 365, Word 2016 or newer).
  The add-in does not work with Word Online (in a browser), Word for macOS or
  Word on mobile devices.

If your Windows is older, the installer says so and tells you what to do.
Installing .NET Framework is a one-time operation and it does require
administrator rights (that is a Microsoft requirement); the add-in itself never
needs them.

## Installation (3 steps)

1. **Close Microsoft Word** (every window; save your documents first).
2. Open the program folder and **double-click `Install.cmd`**.
   A black window appears with the message "INSTALLATION COMPLETED SUCCESSFULLY".
   Press any key to close it.
3. **Start Word.** A tab bar appears right below the ribbon — one tab per open
   document.

If you received the program as a ZIP archive, extract the whole archive first
(right-click → "Extract All…"), then run `Install.cmd` from the extracted folder.

## How to use it

- **Click a tab** — switch to that document.
- **The × on a tab** — close the document. If it has unsaved changes, Word asks
  whether to save them, exactly as usual.
- **The "+" button** — create a new document.
- **A dot next to the name** — the document has unsaved changes.
- **Right-click a tab** — pin it, close it, close the others, give it a colour,
  change the tab size, or open the folder the file lives in.
- **Hover over a tab** — the full path to the file appears.
- **▾ on the right** — a list of all tabs with a search box, and "Settings…".
- **⌄ on the right** — collapse the tab bar; ⌃ brings it back.
- **Ctrl+Tab / Ctrl+Shift+Tab** — next / previous tab (can be switched off in
  Settings).
- With many documents open the tabs shrink first, then ‹ › scroll arrows appear.

## Language

The add-in follows **Word's own interface language**: a Russian Word gets
Russian tabs, anything else gets English. To choose the language yourself:
**▾ → "Settings…" → "Language"** (Same as Word / English / Русский).

The installer and the diagnostics script follow the Windows display language.
You can force a language explicitly:

```
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Lang en
```

## If the tabs do not appear

Work through this list, checking after each step:

1. **Restart Word** — close every Word window, then open it again.
2. **Check Word's add-in list:** File → Options → Add-ins → at the bottom
   "Manage: COM Add-ins" → "Go…". "Tabs for Word" must be there with a tick.
   Tick it if it is not.
3. **Check the disabled items:** in the same place, "Manage: Disabled Items" →
   "Go…". If our add-in or mscoree.dll is listed, select it and press "Enable".
4. **Look at the log:** open `%LOCALAPPDATA%\TabsForWord\Logs` (paste that path
   into the Explorer address bar). The file with today's date shows whether the
   add-in was loaded and what went wrong.
5. **Reinstall:** close Word, run `Repair.cmd`; if that does not help, run
   `Uninstall.cmd` and then `Install.cmd` again.
6. **Collect diagnostics:** run `Diagnostics.cmd` — a ZIP file with a report
   appears on your desktop. The report contains the **names and paths** of your
   documents, your computer name and your user name, but **not the content** of
   any document. Inside the archive are plain text files: you can open and read
   them before sending. Send that file to whoever gave you the program — it
   shows what exactly went wrong.

Two more common causes:

- **Word was started "as administrator"** — the add-in will not load that way
  (it is installed for the ordinary user). Start Word with a normal double-click.
- The program files were **downloaded and not extracted** (running straight from
  the ZIP) — extract the archive and install again.

## If the tab bar misbehaves — the fallback look

Normally the tab bar is built into the Word window itself: it is hard to tell it
apart from Word, and the document area moves down by exactly the height of the
bar. This works on every computer we have tested, but it relies on the internal
structure of the Word window, which Microsoft may change one day.

If the bar does not appear, flickers, or overlaps the document:

1. Run **`Classic-mode.cmd`** from the installation folder.
2. Restart Word.

The bar moves into a separate Word strip with a grey service title above it
(this is how the program looked before version 1.6.0). Every feature keeps
working. To go back: **`In-window-tabs.cmd`** and restart Word.

If you had to do this, please send a `Diagnostics.cmd` report — it shows what
did not work.

## What you should know (honest limitations)

- Every document still opens in its own Windows window — the tab bar is shown
  in each of them and switches between them. There is no "physical" merging of
  windows into one: modern Word offers no official way to do that.
- The bar takes a narrow strip at the top (about the height of one line of text;
  37 pixels at 100% scaling), and the document area moves down by exactly that
  much. The ⌄ button collapses it almost completely. In the fallback "classic"
  mode the bar is taller — around 1.5 cm: Word adds its own service title above
  it, and that cannot be removed.
- For documents in Protected View (files from the internet) there is no bar
  inside the yellow window — the tab appears once you press "Enable Editing".

## Technical details (you can skip this)

The program is installed for the current user only:
- files: `%LOCALAPPDATA%\TabsForWord`
- registry: `HKCU\Software\Classes` (the COM class) and
  `HKCU\Software\Microsoft\Office\Word\Addins\TabsForWord.Connect`
- log: `%LOCALAPPDATA%\TabsForWord\Logs`

Nothing system-wide is changed and other users of the computer are not affected.
