# Allors Excel

[![Allors.Excel](https://img.shields.io/nuget/v/Allors.Excel?label=Allors.Excel)](https://www.nuget.org/packages/Allors.Excel)
[![Allors.Excel.Headless](https://img.shields.io/nuget/v/Allors.Excel.Headless?label=Allors.Excel.Headless)](https://www.nuget.org/packages/Allors.Excel.Headless)
[![Allors.Excel.Interop](https://img.shields.io/nuget/v/Allors.Excel.Interop?label=Allors.Excel.Interop)](https://www.nuget.org/packages/Allors.Excel.Interop)
[![CI](https://github.com/allors/excel/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/allors/excel/actions/workflows/ci.yml)

A virtual workbook for Excel add-ins.

Allors Excel is a .NET library for Excel add-ins. Your add-in reads and writes workbooks,
worksheets and cells as ordinary objects with zero-based coordinates. The changes are collected
and written to Excel in blocks when a worksheet is flushed, and the changes a user makes in Excel
come back as events. The same program runs against an in-memory workbook, so the logic of an
add-in can be tested without Excel.

## Documentation

The [documentation](https://github.com/allors/excel/blob/main/docs/src/content/docs/index.md)
is written for users and for maintainers. Start with
[Getting started](https://github.com/allors/excel/blob/main/docs/src/content/docs/getting-started.md),
a tutorial that fills a worksheet without Excel, then
[How the virtual workbook works](https://github.com/allors/excel/blob/main/docs/src/content/docs/virtual-workbook.md)
and [Build an Excel add-in](https://github.com/allors/excel/blob/main/docs/src/content/docs/build-an-add-in.md).
The [API reference](https://github.com/allors/excel/blob/main/docs/src/content/docs/api.md) lists
every member. Agent instructions are in `AGENTS.md`.

## Example

A program fills the first worksheet of every workbook the add-in starts to track:

```csharp
public sealed class LedgerProgram : IProgram
{
    public Task OnNew(IWorkbook workbook)
    {
        var sheet = workbook.Worksheets[0];
        sheet[0, 0].Value = "Item";
        sheet[0, 1].Value = "Amount";
        sheet[1, 0].Value = "Paper";
        sheet[1, 1].Value = 45m;
        sheet[1, 1].NumberFormat = "0.00";
        sheet[2, 1].Formula = "=SUM(B2:B2)";
        sheet.Flush();
        return Task.CompletedTask;
    }

    // The other members of IProgram return Task.CompletedTask or do nothing.
}
```

Inside Excel, an Excel-DNA or VSTO host creates `Allors.Excel.Interop.AddIn` with Excel's
`Application` object and the program. In a test, the headless add-in plays the host:

```csharp
var addIn = new Allors.Excel.Headless.AddIn(new LedgerProgram(), new NoRibbon());
var sheet = addIn.Workbooks[0].Worksheets[0];
Assert.Equal("Paper", sheet[1, 0].ValueAsString);
```

## Install

```sh
dotnet add package Allors.Excel.Headless   # the abstractions and the in-memory workbook, for tests and tools
dotnet add package Allors.Excel.Interop    # the implementation that runs inside Excel
```

The packages on nuget.org, version `1.0.73` from March 2024, predate renames made on `main`
since; [Packages and versions](https://github.com/allors/excel/blob/main/docs/src/content/docs/packages.md)
lists them. Until a release is tagged from `main`, reference the projects from a clone of the
repository, as the tutorial does.

## Building

The repository builds on Windows only: the VSTO host needs Visual Studio with the Office/VSTO
workload, and the interop tests drive a real Excel.

```sh
task compile                     # restore and build the solution
task test                        # run the five test passes into artifacts/tests
task pack CONFIGURATION=Release  # create the NuGet packages in artifacts/nuget
```

[Build and test](https://github.com/allors/excel/blob/main/docs/src/content/docs/internals/build-and-test.md)
describes the prerequisites, the targets, the tests, CI and the release workflow.

## Solution structure

| Project | Description |
|---|---|
| `src/Allors.Excel` | The abstractions: `IAddIn`, `IProgram`, `IWorkbook`, `IWorksheet`, `ICell` and their value types |
| `src/Allors.Excel.Headless` | The in-memory implementation, for tests and tools |
| `src/Allors.Excel.Interop` | The implementation that runs inside Excel through COM |
| `src/Allors.Excel.Interop.Dna` | A host built with Excel-DNA |
| `src/Allors.Excel.Interop.Vsto` | A host built with VSTO |
| `src/Application` | A sample program for the hosts |
| `src/Allors.Excel.Tests` | The shared facts, run against both implementations |
| `src/Allors.Excel.Headless.Tests` | The shared facts against the headless implementation |
| `src/Allors.Excel.Interop.Tests.Shared`, `src/Allors.Excel.Interop.Dna.Tests`, `src/Allors.Excel.Interop.Vsto.Tests` | The Excel fixture and the shared facts against a real Excel |
| `src/Allors.Excel.Interop.Tests` | Tests of the chunking of cells |
| `docs` | The documentation website |

## Licence

LGPL-3.0-or-later
