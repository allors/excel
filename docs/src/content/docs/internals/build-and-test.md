---
title: "Build and test"
description: "Prerequisites, the Task targets, the five test passes, CI, versioning and the release workflow."
audience: "Maintainers"
kind: "How-to guide"
status: "Current"
sidebar:
  order: 2
---

> **Status: Current.** How-to guide for maintainers. The build and the tests were run locally on
> Windows 11 on 2026-10-19 with the tools named below; the result is recorded in
> [Local run](#local-run).

## Prerequisites

- Windows. The VSTO host needs full MSBuild with the VSTO targets, and the interop implementation
  needs the Office primary interop assemblies, which the build resolves from the GAC or from
  Visual Studio's `Visual Studio Tools for Office\PIA\Office15` folder.
- Visual Studio with the **Office/VSTO** workload. `Taskfile.yml` finds `MSBuild.exe` with
  `vswhere`.
- The .NET SDK 10. There is no `global.json`; the SDK on the machine is used.
- [Task](https://taskfile.dev): `winget install Task.Task` or `choco install go-task`.
- Microsoft Excel, to run the interop test passes. It must open without a modal dialog; see
  [Local run](#local-run).
- Node.js 24 and npm for the documentation website; see the
  [website instructions](../../../../README.md).

## Build with Task

`Taskfile.yml` replaces the earlier Nuke build. Every target runs at most once per invocation.
The configuration is `Debug` unless `CONFIGURATION=Release` is passed.

| Target | What it does |
| --- | --- |
| `task` | `test`, then `pack`. |
| `task clean` | Deletes and recreates `artifacts/`. |
| `task restore` | `dotnet tool restore`, which finds no manifest and restores nothing, then restores the solution with MSBuild, including the VSTO project's `packages.config`. |
| `task compile` | Rebuilds the four test projects, and through them everything else, with MSBuild. |
| `task test` | Runs the five test passes into `artifacts/tests/`; every pass runs even when an earlier one fails, and the target fails at the end when any pass failed. |
| `task pack` | Packs `Allors.Excel` and `Allors.Excel.Headless` with `dotnet pack` and `Allors.Excel.Interop` with MSBuild, with symbols, into `artifacts/nuget/`. |
| `task ci` | `test`. |

```sh
task compile
task test
task pack CONFIGURATION=Release
```

## Test passes

| Pass | Command | Tests | Needs Excel |
| --- | --- | --- | --- |
| `interop-dotnet` | `dotnet test` of `Allors.Excel.Interop.Tests` for `net462` and `net10.0`, to a `.trx` file | `ChunkTests`, 11 facts per framework | No |
| `interop-framework` | `xunit.console` on `Allors.Excel.Interop.Tests` for `net462` | The same 11 facts | No |
| `interop-vsto` | `xunit.console` on `Allors.Excel.Interop.Vsto.Tests` | The 46 shared facts | Yes |
| `headless` | `xunit.console` on `Allors.Excel.Headless.Tests` | The 46 shared facts | No |
| `interop-dna` | `xunit.console` on `Allors.Excel.Interop.Dna.Tests` | The 46 shared facts | Yes |

`xunit.console.exe` comes from the `xunit.runner.console` package in the NuGet cache; the test
projects reference it so that restore downloads it. The Excel passes start a visible Excel
instance for every test and open a copy of `data/blank.xlsx`; [Architecture](architecture.md#tests)
describes the fixture. Expect Excel windows to appear and disappear for about two minutes per
pass.

## Local run

On 2026-10-19, on Windows 11 with Visual Studio 2026 Enterprise, the .NET SDK 10.0.402, Task
3.54.0 and Microsoft 365 Excel 16.0.20430, `task test` built every project and ran all five
passes. The two passes without Excel and the headless pass passed: 11 chunk facts on each of the
three runs and 46 headless facts. The two Excel passes failed 34 and 36 of their 46 facts, for one
reason: every Excel instance the fixture started showed the Microsoft 365 privacy dialog, which
is modal, so the object model rejected calls with `0x800AC472` and `RPC_E_CALL_REJECTED`. The
first fixture failed in its constructor and leaked its Excel instance, which was closed by hand
afterwards. Open Excel once on the machine and dismiss its dialogs before running the Excel passes.

CI passes all 149 tests on both of its matrix legs; the run for the merge of pull request 24 on
2026-06-03 took about fifteen minutes per leg, including the Office installation.

## CI

The `CI` workflow runs on `windows-latest` for pushes and pull requests to `main`, except when
only `docs/`, the documentation workflow or `AGENTS.md` changed. Two matrix legs, `Excel64` and
`Excel32`, install Microsoft 365 Business through Chocolatey, import the VSTO signing certificate
`src/Allors.Excel.Interop.Vsto/ExcelAddIn.VSTO_TemporaryKey.pfx`, install Task and run
`task ci CONFIGURATION=Release`. The `Excel64` leg also packs and uploads the packages as an
artifact. The test results are published as check runs and uploaded as artifacts. A change to
`README.md` runs CI, because the readme is part of every package.

## Versioning

`version.json` sets the version `1.1` for Nerdbank.GitVersioning, with `main` and tags `vX.Y.Z` as
public release branches. A build on `main` is versioned `1.1.<git height>`; a build on another
branch carries the commit, such as `1.1.1-g10d15c6665` observed when packing the documentation
branch. `dotnet nbgv get-version` shows the version of a checkout; the `nbgv` tool is not in a
manifest, so install it globally or read `version.json`.

## Release

Nothing publishes automatically. To release:

1. Make sure `version.json` has the intended major and minor version and that `main` is green.
2. Tag the commit `vX.Y.Z` and push the tag. The `Release` workflow checks the tag out, imports
   the VSTO certificate, installs Task, runs `task pack CONFIGURATION=Release`, logs in to
   nuget.org through trusted publishing as `allors-nuget` in the `nuget-release` environment,
   pushes the `.nupkg` files with `--skip-duplicate` and creates a GitHub release with generated
   notes and the packages attached. The workflow can also be run by hand with an existing tag.
3. Update the note in `README.md` and in [Packages and versions](../packages.md#published-versions)
   that says the published packages predate the current API.

The workflow had not run by 2026-10-19; the repository's only tag was `1.0.1`, and the packages on
nuget.org, `1.0.73`, come from the earlier Azure Pipelines build.

## Documentation

The website in `docs/` has its own checks; the [website instructions](../../../../README.md) list
them, and the `Documentation` workflow runs them for changes to `docs/`, its workflow file,
`AGENTS.md`, `README.md` and `version.json`, for the `/excel/` and `/` deployment paths.
Publication of the website is a manual step described there.
