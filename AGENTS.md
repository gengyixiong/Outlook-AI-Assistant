# Repository guidance

## Canonical implementation

The maintained product is the classic Outlook COM add-in under `desktop/`.
The root Office.js prototype is historical and must not be extended unless a
future tenant deployment route is explicitly approved.

## Privacy invariants

- Never read a mail item before the user clicks the summary button.
- Search may send only the user's natural-language description to an AI API.
- Never send search candidates, message metadata, titles, bodies, attachments,
  or Outlook search results for reranking.
- Never log prompts, mail content, API keys, or AI response content.
- Execute only AQS produced by `AqsQueryCompiler` from allow-listed
  `SearchPlan` fields. Never execute raw model output.

## Compatibility

- Keep production source compatible with C# 5 and .NET Framework 4.0 APIs.
- The local compiler is the .NET Framework `csc.exe`; no .NET SDK is required.
- Avoid adding NuGet dependencies unless the maintenance and installer
  implications are documented and approved.

## Build and verification

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\desktop\build\build.ps1
```

The build must compile the COM assembly, run all unit tests, and create the ZIP
package. Update `AssemblyInfo.cs`, package name, and release documentation
together for version changes.

