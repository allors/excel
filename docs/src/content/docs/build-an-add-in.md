---
title: "Build an Excel add-in"
description: "Host a program in Excel with Excel-DNA or VSTO, wire the ribbon, track only your own workbooks, and test the program headless."
audience: "Users"
kind: "How-to guide"
status: "Current"
---

> **Status: Current.** How-to guide for users. The hosts described are the two in the repository,
> `src/Allors.Excel.Interop.Dna` and `src/Allors.Excel.Interop.Vsto`, as they are on `main`; both
> were built on 2026-10-19.

Your program is a class that implements `IProgram`; [Getting started](getting-started.md) writes
one. To run it inside Excel you need a host: code that Excel loads, that creates an
`Allors.Excel.Interop.AddIn` with Excel's `Application` object and your program, and that owns the
ribbon. This guide shows the two hosts the repository contains and what to add to them.

## Choose a host

| Host | Project | Needs | Notes |
| --- | --- | --- | --- |
| [Excel-DNA](https://excel-dna.net/) | `src/Allors.Excel.Interop.Dna` | The `ExcelDna.AddIn` package; builds with the .NET SDK | The build produces an `.xll` that Excel opens like a workbook. The repository's host targets `net48`. |
| VSTO | `src/Allors.Excel.Interop.Vsto` | Visual Studio with the Office/VSTO workload, full MSBuild and the VSTO runtime | A classic `.csproj` with `packages.config`, a signed manifest and a `ThisAddIn` class. |

Both hosts reference `Allors.Excel.Interop`, and through it `Allors.Excel`. The interop package
targets `netstandard2.0` and uses the Office primary interop assemblies; see
[Packages and versions](packages.md).

## Host with Excel-DNA

The Excel-DNA host is a ribbon class. Excel-DNA calls `GetCustomUI` when it loads the ribbon and
`OnLoad` when the ribbon is ready; the repository's `Ribbon` in `src/Allors.Excel.Interop.Dna/Ribbon.cs`
creates the add-in in the first and starts the program in the second:

```csharp
[ComVisible(true)]
public class Ribbon : ExcelRibbon, IRibbon
{
    public IRibbonUI RibbonUI { get; private set; }

    public AddIn AddIn { get; private set; }

    public Program Program { get; private set; }

    public override string GetCustomUI(string _)
    {
        var application = ExcelDnaUtil.Application;
        this.Program = new Program(new ServiceLocator());
        this.AddIn = new AddIn((InteropApplication)application, this.Program, this);
        return RibbonResources.Ribbon;
    }

    public async void OnLoad(IRibbonUI ribbon)
    {
        this.RibbonUI = ribbon;
        await this.Program.OnStart(this.AddIn);
    }

    public override object LoadImage(string imageId) => RibbonResources.ResourceManager.GetObject(imageId);

    public void OnButtonPressed(IRibbonControl control) => MessageBox.Show("Hello!");

    public void Invalidate() => this.RibbonUI.Invalidate();
}
```

`RibbonResources.Ribbon` is the ribbon XML, kept as a resource; it declares a tab with one button
whose `onAction` is `OnButtonPressed`. `Invalidate` implements `IRibbon`, so that the program can
ask Excel to query the ribbon's callbacks again.

To build your own:

1. Create a class library, add the `ExcelDna.AddIn` package and a reference to
   `Allors.Excel.Interop` and `Allors.Excel`, and set `EnableDynamicLoading` to `true` so that the
   dependencies are copied next to the add-in, as the repository's project does.
2. Add a ribbon class like the one above, with your program instead of the sample `Program`.
3. Build. The output directory contains `<Project>-AddIn64.xll` and `<Project>-AddIn.xll`;
   open the one matching your Excel's bitness in Excel, or register it under
   **File › Options › Add-ins › Manage Excel Add-ins**.

## Host with VSTO

The VSTO host, `src/Allors.Excel.Interop.Vsto/ThisAddIn.cs`, creates the add-in at startup and
stops the program at shutdown:

```csharp
private async void ThisAddIn_Startup(object sender, EventArgs e) => await Task.Run(async () =>
{
    this.serviceLocator = new ServiceLocator();
    var program = new Program(this.serviceLocator);

    this.addIn = new AddIn(this.Application, program, this.Ribbon);

    this.Ribbon.AddIn = this.addIn;
    await program.OnStart(this.addIn);
});

private async void ThisAddIn_Shutdown(object sender, EventArgs e) => await Task.Run(async () =>
{
    await this.addIn.Program.OnStop();
});

protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
{
    SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
    this.Ribbon = new Ribbon();
    return this.Ribbon;
}
```

Its `Ribbon` implements `IRibbonExtensibility` and `IRibbon`, returns the embedded `Ribbon.xml`
from `GetCustomUI` and stores the `IRibbonUI` in `Ribbon_Load`. The project signs its manifest with
`ExcelAddIn.VSTO_TemporaryKey.pfx`, which CI imports before building; Visual Studio registers the
add-in for debugging when you run the project. The ribbon XML of this host names the callbacks
`GetDoSomethingLabel` and `OnClick`, which the class does not implement; add them before relying
on its buttons, as [Architecture](internals/architecture.md#known-defects) records.

## Wire the ribbon

The library does not call `OnHandle`, `IsEnabled`, `OnLogin` or `OnLogout`; they are yours to
call from ribbon callbacks. The pattern is to give every button an `onAction` that forwards the
control's id to the program, and a `getEnabled` that asks the program:

```xml
<button id="hide" label="Hide rows" onAction="OnAction" getEnabled="GetEnabled" />
```

```csharp
public async void OnAction(IRibbonControl control) => await this.Program.OnHandle(control.Id);

public bool GetEnabled(IRibbonControl control) => this.Program.IsEnabled(control.Id, control.Tag);
```

The sample program in `src/Application/Program.cs` handles the ids in its `Actions` class: `hide`
hides rows 5 to 9 of the active worksheet and flushes it. When the program changes what a button
should show, it calls `IRibbon.Invalidate()` and Excel queries the callbacks again.

## Work with workbooks

The add-in tracks the workbooks Excel reports after it started: a new workbook, an opened one or
one the user activates. `IAddIn.Workbooks` lists them; `IsActive` tells which one is active. A
workbook that was already open when the add-in loaded joins `Workbooks` when it is activated,
without an `OnNew`.

Set `ExistentialAttribute` to the name of a custom document property before workbooks are
opened, and the add-in tracks only workbooks that have that property. Write the property with
`workbook.CustomProperties.SetString(name, value)` into the workbooks your add-in creates, and
other people's workbooks are left alone.

Inside a tracked workbook, write through cells and call `Flush` on the worksheet, as
[How the virtual workbook works](virtual-workbook.md#writing-dirty-cells-and-flush) explains.
To react to edits, subscribe to `IWorksheet.CellsChanged`, or let a [`Binder`](bind-cells.md)
do it.

Excel's object model must be used from the thread that owns it. The library's `Flush` writes in
parallel and retries calls that Excel rejects, but a program that uses `async` methods should
continue on the host's synchronization context, which the VSTO host sets to a
`WindowsFormsSynchronizationContext`, rather than on a thread-pool thread.

## Test the program without Excel

`Allors.Excel.Headless.AddIn` runs the same program against an in-memory workbook:

```csharp
var addIn = new Allors.Excel.Headless.AddIn(new MyProgram(), new NoRibbon());
var sheet = addIn.Workbooks[0].Worksheets[0];
Assert.Equal("Item", sheet[0, 0].ValueAsString);
```

The headless add-in creates one workbook with one worksheet and calls `OnNew` for both and then
`OnStart`; `AddWorkbook()` adds another. It stores values and formulas without calculating, raises
no change events, and the operations that only Excel can perform do nothing or throw; the
[API reference](api.md#headless-and-interop-differences) lists them. The repository's own
`Allors.Excel.Tests` are written against the interfaces and run against both the headless add-in
and a real Excel; copy that arrangement for a program that must behave the same in both.
