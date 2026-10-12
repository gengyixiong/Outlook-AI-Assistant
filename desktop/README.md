# Desktop COM add-in

This directory contains the maintained classic Outlook implementation.

The project deliberately avoids VSTO project templates and NuGet packages.
It uses the Office PIAs already installed with classic Outlook and the C#
compiler plus reference assemblies from the .NET Framework 4.8 Developer
Pack. This makes the build repeatable on the current machine without Visual
Studio or the .NET SDK.

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

The distributed package also contains `embedding_service/`, the optional
local Python semantic search service. It is a separate process so the add-in
keeps zero NuGet dependencies and never blocks the Outlook UI thread. See
`docs/SEMANTIC-SEARCH.md`.

## Source layout

```text
src/OutlookAiAssistant/
  AddIn/            COM entry point
  AI/               OpenAI-compatible chat client
  Configuration/    settings + DPAPI API key
  Diagnostics/      local log
  EntityIndex/      contact index
  Models/           detached snapshots
  Outlook/          COM services (context, search, release)
  Search/           search planning + AQS compiler
  SemanticSearch/   local semantic search client, indexer, locator, UI facade
  Summary/          Executive Brief
  UI/               task pane + settings form
tests/OutlookAiAssistant.UnitTests/   self-contained test runner
```

