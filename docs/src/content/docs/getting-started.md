---
title: "Getting started"
description: "Write a program that fills a worksheet, run it against the headless workbook without Excel, and read the result back."
audience: "Users"
kind: "Tutorial"
status: "Current"
---

> **Status: Current.** Tutorial for users. Every step was run with the .NET SDK 10.0.402 against
> the repository at commit `10d15c6` on 2026-10-19; the output shown is the output observed.

In this tutorial you write a program that fills a worksheet with a small ledger, run it against
the headless workbook, which needs no Excel, and read the cells back. The result is an `IProgram`
of the kind an Excel add-in hosts, exercised from a console application.

You need the [.NET SDK](https://dotnet.microsoft.com/download) 10 or later, Git and a terminal.
The packages on nuget.org predate the current API, so the tutorial references the library's
projects from a clone of the repository; [Packages and versions](packages.md#published-versions)
explains the difference.

## 1. Create the project

Create a directory for the tutorial and, inside it, clone the repository and create a console
application that references the abstractions and the headless implementation:

```sh
git clone https://github.com/allors/excel.git
dotnet new console -n Ledger -o Ledger -f net10.0
dotnet add Ledger reference excel/src/Allors.Excel/Allors.Excel.csproj excel/src/Allors.Excel.Headless/Allors.Excel.Headless.csproj
```

## 2. Write the program

Replace the content of `Ledger/Program.cs`:

```csharp
using System.Drawing;
using System.Globalization;
using Allors.Excel;
using Allors.Excel.Headless;

CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

var addIn = new AddIn(new LedgerProgram(), new NoRibbon());
var sheet = addIn.Workbooks[0].Worksheets[0];

var used = sheet.GetUsedRange();
var rows = used.Rows ?? 0;
var columns = used.Columns ?? 0;
Console.WriteLine($"{sheet.Name}: {rows} rows, {columns} columns");
for (var row = 0; row < rows; row++)
{
    var cells = Enumerable.Range(0, columns).Select(column => sheet[row, column]);
    Console.WriteLine(string.Join(" | ", cells.Select(cell => cell.Formula ?? cell.ValueAsString)));
}

public sealed class LedgerProgram : IProgram
{
    public Task OnStart(IAddIn addIn) => Task.CompletedTask;

    public Task OnStop() => Task.CompletedTask;

    public Task OnNew(IWorkbook workbook)
    {
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Ledger";

        var header = new Style(Color.LightGray, Color.Black);
        string[] titles = ["Item", "Quantity", "Price", "Amount"];
        for (var column = 0; column < titles.Length; column++)
        {
            sheet[0, column].Value = titles[column];
            sheet[0, column].Style = header;
        }

        (string Item, int Quantity, decimal Price)[] lines = [("Paper", 10, 4.5m), ("Toner", 2, 59m), ("Stapler", 1, 12.25m)];
        for (var i = 0; i < lines.Length; i++)
        {
            var row = i + 1;
            sheet[row, 0].Value = lines[i].Item;
            sheet[row, 1].Value = lines[i].Quantity;
            sheet[row, 2].Value = lines[i].Price;
            sheet[row, 2].NumberFormat = "0.00";
            sheet[row, 3].Formula = $"=B{row + 1}*C{row + 1}";
        }

        sheet[4, 2].Value = "Total";
        sheet[4, 3].Formula = "=SUM(D2:D4)";

        sheet.Flush();
        return Task.CompletedTask;
    }

    public void OnClose(IWorkbook workbook, ref bool cancel)
    {
    }

    public Task OnNew(IWorksheet worksheet) => Task.CompletedTask;

    public Task OnBeforeDelete(IWorksheet worksheet) => Task.CompletedTask;

    public Task OnHandle(string handle, params object[] argument) => Task.CompletedTask;

    public Task OnLogin() => Task.CompletedTask;

    public Task OnLogout() => Task.CompletedTask;

    public bool IsEnabled(string controlId, string controlTag) => true;
}

public sealed class NoRibbon : IRibbon
{
    public void Invalidate()
    {
    }
}
```

`LedgerProgram` is the program: the class an add-in hosts. The library calls `OnNew(IWorkbook)`
for every workbook it starts to track, and that is where the ledger is written. Cells are
addressed by zero-based row and column, so `sheet[0, 0]` is `A1`, and a cell exists as soon as
it is asked for. Setting a `Value`, a `Style`, a `NumberFormat` or a `Formula` changes the cell
object; `Flush()` writes the changes to Excel, in blocks, when the program runs in Excel. The other
members of `IProgram` are the hooks for the host's ribbon and lifecycle, which this tutorial does
not use; [How the virtual workbook works](virtual-workbook.md#the-lifecycle-of-a-program) describes
them.

The headless `AddIn` plays the host: its constructor creates a workbook with one worksheet and
calls the program's `OnNew` for both, so the sheet is filled by the time the constructor returns.
`NoRibbon` stands in for the ribbon the program could invalidate. The culture is set to the
invariant culture so that `ValueAsString` prints `4.5` on every machine.

## 3. Run it

```sh
dotnet run --project Ledger
```

```text
Ledger: 5 rows, 4 columns
Item | Quantity | Price | Amount
Paper | 10 | 4.5 | =B2*C2
Toner | 2 | 59 | =B3*C3
Stapler | 1 | 12.25 | =B4*C4
 |  | Total | =SUM(D2:D4)
```

The build also prints analyser warnings from the repository's own projects; they do not concern
the tutorial. `GetUsedRange()` returns the block from the first to the last cell that exists:
five rows and four columns. The formulas come back as the text you set, because the headless
workbook stores and calculates nothing. In Excel the same program shows `45`, `118`, `12.25` and
`175.25` in the amount column, with the `Price` cells formatted to two decimals and the header row
shaded.

## Where to go next

[Build an Excel add-in](build-an-add-in.md) hosts a program like this one in Excel with
Excel-DNA or VSTO and wires the ribbon to `OnHandle`. [How the virtual workbook works](virtual-workbook.md)
explains what `Flush` writes, in which order, and how the changes a user makes in Excel come back
as events. The [API reference](api.md) lists every member, with the differences between the
headless and the interop implementation.
