# Environment (stage 0)

Checked on: 2026-07-23. This records the development machine the add-in was
built on, and why the build does not need Visual Studio.

## System

| Item | Value |
|---|---|
| OS | Windows 11 Pro, build 10.0.26200 |
| Architecture | x64 (AMD64) |

## Development tools

| Tool | Status |
|---|---|
| Git | yes, 2.55.0.windows.3 |
| .NET Framework (runtime) | yes, 4.8.1 (Release 533509) - part of Windows 11 |
| .NET Framework 4.8 Developer Pack (reference assemblies) | no |
| dotnet SDK | no (installed later, see below) |
| Visual Studio | no (vswhere absent) |
| Visual Studio Build Tools / MSBuild | no |
| Office Developer Tools / VSTO templates | no |
| VSTO Runtime | no |
| csc.exe (the C# 5 compiler shipped with .NET Framework) | yes, C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe |
| RegAsm.exe (x64) | yes, C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe |

## Microsoft Word

| Item | Value |
|---|---|
| Product | Microsoft 365 (O365HomePremRetail), Click-to-Run |
| Version | 16.0.20131.20154 |
| Bitness | **x64** |
| Path | C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE |
| Installed | yes (and running at the time of the check) |

## What this means for the build

1. A classic VSTO project (a `.csproj` with the Office targets) cannot be built
   here: no Visual Studio, no VSTO SDK, no MSBuild. Installing the full VS means
   gigabytes and, most likely, a UAC prompt to confirm by hand.
2. The safe options available:
   - install the **.NET SDK** into a user folder with the official
     `dotnet-install.ps1` script (no administrator rights) and build the project
     as an ordinary `net48` class library;
   - or use the built-in `csc.exe` (limited to C# 5).
3. Everything needed to **run** the add-in is already on the machine:
   .NET Framework 4.8.1 and 64-bit Word.
4. The add-in can be registered per user (HKCU) through RegAsm or .reg files,
   so no administrator rights are needed.

## Filling the gaps

- [x] Established that VS/VSTO are not required once the COM add-in architecture
      is chosen (see docs/RESEARCH.md and docs/DECISIONS.md).
- [x] Installed the .NET SDK **8.0.423** (user scope, no administrator rights):
      `%LOCALAPPDATA%\Microsoft\dotnet\dotnet.exe`
- [x] Build smoke test: an SDK-style `net48` project with `UseWindowsForms=true`
      and the `Microsoft.NETFramework.ReferenceAssemblies` package builds
      successfully (0 errors).
- [x] Verified that classic COM add-ins already work in Word on this machine
      (OneNote.WordAddinTakeNotesService, PDFMaker.OfficeAddin under
      HKCU\...\Word\Addins) and that no restrictive Trust Center policy is in
      place - so the COM add-in mechanism itself is available.
