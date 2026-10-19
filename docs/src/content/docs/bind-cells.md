---
title: "Bind cells to your model"
description: "Use Binder and Binding to push model values into cells and to take edits back, with a style for changed cells."
audience: "Users"
kind: "How-to guide"
status: "Current"
---

> **Status: Current.** How-to guide for users. It describes `Binder`, `Binding` and `IBinding` in
> `src/Allors.Excel/Binding` as they are on `main`. No test in the repository covers them; the
> behaviour below follows the source, read on 2026-10-19.

A `Binder` connects the cells of one worksheet to your model. For each cell you register a
`Binding` that says how to fill the cell from the model and, optionally, how to take an edit
back. The binder fills the cells on request and listens to the worksheet's `CellsChanged` event
for edits.

## Register bindings and fill the cells

Create one binder per worksheet. Register a binding for each cell, then call `ToCells()` to fill
them, then flush the worksheet:

```csharp
var binder = new Binder(sheet, changedStyle: new Style(Color.LightYellow, Color.Black));

binder.Set(1, 0, new Binding(toCell: cell => cell.Value = person.Name,
                             toDomain: cell => person.Name = cell.ValueAsString));
binder.Set(1, 1, new Binding(toCell: cell => cell.Value = person.BirthDate));

foreach (var obsolete in binder.ToCells())
{
    obsolete.Clear();
}

sheet.Flush();
```

`Set(row, column, binding)` and `Set(cell, binding)` register a binding for the next `ToCells()`.
`ToCells()` calls every binding's `ToCell` and returns the cells that were bound before but were
not registered again, so that you can clear them; a binding that is not registered again before
the next `ToCells()` is dropped. Register every binding again on every refresh, as the example
does, or keep the registrations in a method you call each time the model changes.

A `Binding` with only `toCell` is one-way; a `Binding` with `toDomain` is two-way. `OneWayBinding`
and `TwoWayBinding` on `IBinding` tell which, and `ExistBinding(row, column)` tells whether a cell
has been registered since the last `ToCells()`.

## Take edits back

When the user edits cells, the worksheet raises `CellsChanged` and the binder handles the bound
cells among them:

- For a two-way binding it calls `ToDomain(cell)`, so your model receives the edit. If the binder
  was created with a `changedStyle`, it remembers the cell's previous style the first time and
  applies the changed style, so that edited cells stand out.
- For a one-way binding it calls `ToCell(cell)` again, which puts the model's value back into the
  cell. The cell is dirty afterwards; call `Flush` to show the model's value in Excel again.

After each batch of changes the binder raises `ToDomained`. Handle it to validate the model, to
recompute dependent cells with `ToCells()`, and to flush:

```csharp
binder.ToDomained += (sender, args) =>
{
    binder.Set(1, 0, nameBinding);
    binder.Set(1, 1, birthDateBinding);
    binder.ToCells();
    sheet.Flush();
};
```

`ResetChangedCells()` restores the remembered styles of the edited cells and forgets them; call it
after the model has been saved, then flush.

The worksheet raises `CellsChanged` only for values that Excel reports as changed, converted by
each cell's `IValueConverter`; [How the virtual workbook works](virtual-workbook.md#reading-what-excel-reports)
explains the conversion. The headless worksheet never raises the event, so a two-way binding can be
tested only against Excel, while `ToCells()` can be tested headless.
