---
title: "Documentation"
description: "Documentation for developers who build Excel add-ins with Allors Excel, and for its maintainers."
audience: "Users and maintainers"
kind: "Index"
status: "Current"
---

> **Status: Current.**

Allors Excel is a .NET library for Excel add-ins. It gives an add-in a virtual workbook: workbooks,
worksheets and cells that the add-in reads and writes as ordinary objects with zero-based
coordinates. Changes are collected and written to Excel in blocks when a worksheet is flushed,
and the changes a user makes in Excel come back as events. The same program runs against an
in-memory workbook, so the logic of an add-in can be tested without Excel.

Users write an `IProgram` and host it in Excel with Excel-DNA or VSTO, or run it headless in tests
and tools. Maintainers change the library, the hosts, the build and these pages.
[README.md](../../../../README.md) introduces the project. The
[writing guide](internals/writing-documentation.md) defines the rules for these pages;
[Terminology](terminology.md) records the agreed names. These pages are for humans. Agents use
the same technical content, following the instructions in `AGENTS.md`.

Allors Excel is one of Allors's supporting open-source tools. Its packages are licensed under
`LGPL-3.0-or-later`, as their package metadata states. Ask questions and report problems through
the [repository's issues](https://github.com/allors/excel/issues).

## For users

| Page | Kind | Status | Content |
| --- | --- | --- | --- |
| [Getting started](getting-started.md) | Tutorial | Current | Write a program that fills a worksheet, run it against the headless workbook without Excel, and read the result back. |
| [How the virtual workbook works](virtual-workbook.md) | Explanation | Current | The abstractions, the two implementations, the program's lifecycle, dirty tracking and `Flush`, and how changes made in Excel come back. |
| [Build an Excel add-in](build-an-add-in.md) | How-to guide | Current | Host a program in Excel with Excel-DNA or VSTO, wire the ribbon, track only your own workbooks, and test the program headless. |
| [Bind cells to your model](bind-cells.md) | How-to guide | Current | Use `Binder` and `Binding` to push model values into cells and to take edits back, with a style for changed cells. |
| [API reference](api.md) | Reference | Current | Every public type and member of `Allors.Excel`, and what the headless and interop implementations do with it. |
| [Packages and versions](packages.md) | Reference | Current | The NuGet packages, their target frameworks and dependencies, the published versions and how they differ from `main`, versioning and the licence. |
| [Terminology](terminology.md) | Reference | Current | Agreed names for the add-in, the program, the virtual workbook, flushing and the implementations. |

## For maintainers

The pages under `internals/` describe how the library and the repository are built. A user must
not rely on them: they can change with any version.

| Page | Kind | Status | Content |
| --- | --- | --- | --- |
| [Architecture](internals/architecture.md) | Explanation | Current | The projects and their references, the interop render pipeline, the headless implementation, the shared tests and the Excel fixture, and the known defects. |
| [Build and test](internals/build-and-test.md) | How-to guide | Current | Prerequisites, the Task targets, the five test passes, CI, versioning and the release workflow. |
| [Writing documentation](internals/writing-documentation.md) | How-to guide | Current | Audience, Diátaxis, status, Oxford English, stable terminology, visual explanations and review. |
| [Documentation website](internals/documentation-website.md) | Decision record | Current | The adoption of the Allors documentation standard, what was reused, validation and where the website is published. |

The writing guide defines [page kinds](internals/writing-documentation.md#choose-the-audience-and-purpose)
and [statuses](internals/writing-documentation.md#state-what-is-current). Audience, kind and
status describe different things: who a page serves, what it helps them do and whether it
describes current behaviour.
