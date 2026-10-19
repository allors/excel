---
title: "Terminology"
description: "Agreed names for the add-in, the program, the virtual workbook, flushing and the implementations."
audience: "Users and maintainers"
kind: "Reference"
status: "Current"
---

> **Status: Current.** Reference for users and maintainers.

Use these names consistently when discussing Allors Excel, in prose, code comments, diagrams and
pull requests. Exact identifiers keep their declared spelling. The linked pages own the
definitions.

## The parts

| Term | Definition and usage |
| --- | --- |
| Program | The class that implements `IProgram`: the add-in's logic, which receives the lifecycle callbacks and works with the virtual workbook. Say *program* for the class and *the add-in* for the whole. [How the virtual workbook works](virtual-workbook.md#the-lifecycle-of-a-program) describes the callbacks. |
| Add-in | The `IAddIn`: the root object that holds the workbooks, the ribbon and the program. In Excel it is created by the host; headless it is created directly. |
| Host | The code that loads inside Excel and creates the add-in: an Excel-DNA ribbon or a VSTO `ThisAddIn`. [Build an Excel add-in](build-an-add-in.md) describes both. |
| Ribbon | The `IRibbon`: the host's ribbon, which the program asks to `Invalidate` so that Excel queries its callbacks again. |
| Virtual workbook | The `IWorkbook`, `IWorksheet`, `ICell`, `IRow` and `IColumn` objects, which the program reads and writes. Say *workbook*, *worksheet* and *cell* for them; say *Excel's workbook* or *the interop workbook* for the COM objects behind them. |
| Coordinates | A cell's row and column, both zero-based, as in `sheet[0, 0]` for `A1`. Excel's own indexes and A1 names are one-based; the `Index` of a worksheet is Excel's one-based index. |
| Range | An `Allors.Excel.Range`: a zero-based row and column, optional counts, an optional worksheet and an optional name. Not Excel's `Range` object, which the interop implementation calls an *interop range*. |
| Named range | An Excel name that refers to a range, scoped to the workbook or to a worksheet. |
| Options | A cell's `Options`: a range whose values are offered as the cell's list validation. |

## Writing and reading

| Term | Definition and usage |
| --- | --- |
| Dirty | A cell or row whose change has not been written to Excel yet. Setting a value, formula, comment, style, number format, options or the hidden state marks it dirty. |
| Flush | `IWorksheet.Flush()`: writing every dirty change of a worksheet to Excel in blocks. Say that a worksheet is flushed; the implementation's methods are called `Render…`, but *render* is not used in prose. |
| Chunk | A rectangular block of adjacent dirty cells that the interop implementation writes with one COM call. |
| Change event | `CellsChanged` and `SheetChanged`: raised by a worksheet with the cells whose values Excel reported as changed. |
| Value converter | An `IValueConverter` that turns the value Excel reports into the value the cell keeps; `DefaultValueConverter` uses the cell's previous value to choose a conversion. [Value conversion](api.md#value-conversion) describes its rules and limits. |
| Binding | A `Binding` registered with a `Binder`: how a cell's value is produced from the model and, for a two-way binding, how an edit goes back to the model. See [Bind cells to your model](bind-cells.md). |
| Existential attribute | `IAddIn.ExistentialAttribute`: the name of a custom document property; when set, the interop add-in tracks only workbooks that have that property. |

## Implementations

| Term | Definition and usage |
| --- | --- |
| Headless | The `Allors.Excel.Headless` implementation: the workbook kept in memory, without Excel, for tests and tools. |
| Interop | The `Allors.Excel.Interop` implementation: the workbook backed by Excel's COM objects through the primary interop assemblies. |
| Excel-DNA | The [Excel-DNA](https://excel-dna.net/) add-in framework, used by the `Allors.Excel.Interop.Dna` host. |
| VSTO | Visual Studio Tools for Office, used by the `Allors.Excel.Interop.Vsto` host. |
