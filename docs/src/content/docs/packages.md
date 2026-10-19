---
title: "Packages and versions"
description: "The NuGet packages, their target frameworks and dependencies, the published versions and how they differ from main, versioning and the licence."
audience: "Users"
kind: "Reference"
status: "Current"
---

> **Status: Current.** Reference for users. The published versions were checked on nuget.org on
> 2026-10-19; the package contents describe `main` at that date.

## Packages

| Package | Contents | Target framework | Depends on |
| --- | --- | --- | --- |
| `Allors.Excel` | The abstractions `IAddIn`, `IProgram`, `IRibbon`, `IWorkbook`, `IWorksheet`, `ICell`, `IRow` and `IColumn`; `Range`, `Style`, `PageSetup`, `PageHeaderFooter`, `Hyperlink` and `CellChangedEvent`; the document property interfaces; `IValueConverter` and `DefaultValueConverter`; `Binder`, `Binding` and `IBinding`; `Utils`. | `netstandard2.0` | nothing |
| `Allors.Excel.Headless` | The in-memory implementation: `AddIn`, `Workbook`, `Worksheet`, `Cell`, `Row` and `Column`. | `netstandard2.0` | `Allors.Excel` |
| `Allors.Excel.Interop` | The implementation that runs inside Excel: `AddIn`, `Workbook`, `Worksheet`, `Cell`, `Row`, `Column`, the property classes and the chunking of cells. | `netstandard2.0` on `main`; `net461` in the published `1.0.73` | `Allors.Excel`, `Microsoft.CSharp` 4.7.0, `Microsoft.Office.Interop.Excel` 15.0.4795.1001, and the `Office` primary interop assembly, which the build resolves from the GAC or from Visual Studio's PIA folder |

The hosts `Allors.Excel.Interop.Dna` and `Allors.Excel.Interop.Vsto` and the sample `Application`
are built from the repository but not packaged; [Build an Excel add-in](build-an-add-in.md)
describes them.

All packages carry the licence expression `LGPL-3.0-or-later`, the repository's `README.md` as
their readme and `icon.png` as their icon. Their project URL points at
`https://github.com/Allors/Documents`, which is the repository of another Allors project; the
repository URL is this repository. Correcting the project URL is recorded in
[Architecture](internals/architecture.md#known-defects).

## Published versions

| Package | Latest version | Published |
| --- | --- | --- |
| `Allors.Excel` | `1.0.73` | 2024-03-11 |
| `Allors.Excel.Headless` | `1.0.73` | 2024-03-11 |
| `Allors.Excel.Interop` | `1.0.73` | 2024-03-11 |

Nineteen versions from `1.0.0` to `1.0.73` are listed for each package. `1.0.73` was built from
commit `dece84dcd2e3fe65459ad65d2ee03af2806b6ab5`, which predates the changes that `main` has
received since March 2024. The names of these members changed after that version:

| In `1.0.73` | On `main` |
| --- | --- |
| `IWorkbook.SetCustomXML`, `GetCustomXMLById`, `TryDeleteCustomXMLById` | `SetCustomXml`, `GetCustomXmlById`, `TryDeleteCustomXmlById` |
| `IWorksheet.SaveAsPDF`, `SaveAsXPS` | `SaveAsPdf`, `SaveAsXps` |
| `Task IWorksheet.Flush()`, `Task RefreshPivotTables()` | `void Flush()`, `void RefreshPivotTables()` |
| `IWorksheet.GetUsedRange(string column)`, `GetUsedRange(int row)` | Removed; `GetUsedRange()` remains |
| `ICell.Tag` | Removed |
| `Range.Name { get; }` | `Range.Name { get; set; }` |

These pages describe `main`. Until a release from `main` is published, use the repository, as
[Getting started](getting-started.md) does, or read the `1.0.73` packages with the names above.

## Versioning and release

Versions come from [Nerdbank.GitVersioning](https://github.com/dotnet/Nerdbank.GitVersioning)
and `version.json`, whose version is `1.1`. A build on `main` or on a tag `vX.Y.Z` is a public
release build, versioned `1.1.<git height>`; `dotnet nbgv get-version` shows the version of a
checkout. The `CI` workflow builds, tests and packs on every push and pull request and uploads the
packages as a workflow artifact, but publishes nothing.

Publication is the `Release` workflow: pushing a tag `vX.Y.Z`, or running the workflow by hand
with an existing tag, packs the three packages on a Windows runner with the VSTO workload, pushes
them to nuget.org through trusted publishing and creates a GitHub release with the packages
attached. On 2026-10-19 the workflow had never run and the repository's only tag was `1.0.1`, from
before the workflow existed; the `1.0.73` packages were published by the earlier build setup.
[Build and test](internals/build-and-test.md#release) has the maintainer's steps.

## Licence

The packages declare `LGPL-3.0-or-later` in their metadata and the repository's `LICENSE` file
holds the text of the GNU Lesser General Public License, version 3. The source files refer to the
same licence in their headers.
