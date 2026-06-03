# Allors Excel

Allors.Excel is a c# Excel VSTO AddIn. It speeds up access to excel by using a virtual DOM to update cells.
It contains useful features to programmatically manage workbooks, worksheets, cells.

[![CI](https://github.com/allors/excel/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/allors/excel/actions/workflows/ci.yml)

# Building

This repository builds on Windows only (legacy VSTO add-in + Excel COM automation tests).

## Prerequisites
- Visual Studio with the **Office/VSTO** workload (provides `msbuild.exe` and VSTO targets)
- .NET 10 SDK
- [go-task](https://taskfile.dev): `choco install go-task` or `winget install Task.Task`
- Microsoft Excel (required only to run the tests)

## Commands (from the repository root)
| Command        | Description                                       |
| -------------- | ------------------------------------------------- |
| `task`         | Run tests, then produce NuGet packages            |
| `task clean`   | Create/clean `artifacts/`                         |
| `task restore` | Restore solution packages and dotnet tools        |
| `task compile` | Rebuild the test projects                         |
| `task test`    | Run all test passes into `artifacts/tests`        |
| `task pack`    | Produce NuGet packages into `artifacts/nuget`     |
| `task ci`      | CI target (build + test)                          |

Configuration defaults to `Debug`; override with `task <target> CONFIGURATION=Release`.
Versions come from [Nerdbank.GitVersioning](https://github.com/dotnet/Nerdbank.GitVersioning) (`version.json`); inspect with `dotnet nbgv get-version`.

# Installing via Nuget
	Install-Package Allors.Excel
 
# Features

## Workbook

### Properties
- IsActive will activate that workbook
- Worksheets contains the worksheets inside the workbook

### Methods
- GetNamedRanges(string refersToSheetName)
	
	Return a list of Excel.Ranges
- SetNamedRange(string name, Excel.Range range)

	Adds or updates the namedRange

- Copy(IWorksheet source, IWorksheet beforeWorksheet)

	Copies the source workbook to this workbook

## Worksheet

### Properties


### Methods


### Indexers


### Events