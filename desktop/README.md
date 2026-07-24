# Desktop COM add-in

This directory contains the maintained classic Outlook implementation.

The project deliberately avoids VSTO project templates and NuGet packages.
It uses the Office PIAs already installed with classic Outlook and the C#
compiler included with .NET Framework. This makes the build repeatable on the
current machine without Visual Studio.

## Build

From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\desktop\build\build.ps1
```

Useful switches:

```powershell
# Compile and test without creating a ZIP
.\desktop\build\build.ps1 -SkipPackage

# Compile only
.\desktop\build\build.ps1 -SkipPackage -SkipTests

# Debug symbols and no optimization
.\desktop\build\build.ps1 -Configuration Debug -SkipPackage
```

The `.csproj` is included for navigation in Visual Studio or Rider. The
PowerShell build script is the canonical build because it resolves the Office
PIA locations on each computer.

