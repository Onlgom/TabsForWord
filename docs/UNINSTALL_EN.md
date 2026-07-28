# Removing "Tabs for Microsoft Word"

(Русская версия: [UNINSTALL_RU.md](UNINSTALL_RU.md).)

## Normal uninstall (one minute)

1. **Close Microsoft Word** (every window).
2. Double-click **`Uninstall.cmd`** (it sits next to `Install.cmd` in the folder
   you installed from).
3. The message "UNINSTALL COMPLETE" appears. That is all.

Everything goes: the program files, the registry entries and the logs.
Your Word documents are not touched in any way.

## If Uninstall.cmd is gone

You can remove it by hand:

1. Close Word.
2. Press Win+R, type `regedit`, press Enter.
3. Delete these keys (right-click → Delete):
   - `HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\TabsForWord.Connect`
   - `HKEY_CURRENT_USER\Software\Classes\TabsForWord.Connect`
   - `HKEY_CURRENT_USER\Software\Classes\TabsForWord.TabStripControl`
   - `HKEY_CURRENT_USER\Software\Classes\CLSID\{A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A}`
   - `HKEY_CURRENT_USER\Software\Classes\CLSID\{D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8}`
   - the same two CLSID keys under
     `HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID` if they are there
     (they exist so that 32-bit Word can see the add-in)
4. Delete the folder `%LOCALAPPDATA%\TabsForWord`
   (paste that path into the Explorer address bar).

## Turning it off without removing it

File → Options → Add-ins → "Manage: COM Add-ins" → "Go…" → clear the tick next
to "Tabs for Word" → OK. To switch it back on, tick it again.
