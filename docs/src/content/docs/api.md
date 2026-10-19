---
title: "API reference"
description: "Every public type and member of Allors.Excel, and what the headless and interop implementations do with it."
audience: "Users"
kind: "Reference"
status: "Current"
---

> **Status: Current.** Reference for users. It describes the three packages as they are on
> `main`, read on 2026-10-19; the shared tests named in
> [How the virtual workbook works](virtual-workbook.md#where-the-facts-live) cover the members
> they name, and the rest follows the source.

The abstractions live in the namespace `Allors.Excel`, the implementations in
`Allors.Excel.Headless` and `Allors.Excel.Interop`. Rows and columns are zero-based throughout; a
worksheet's `Index` is Excel's one-based index. Where the two implementations behave differently,
the tables say so, and [Headless and interop differences](#headless-and-interop-differences)
collects the differences.

## IAddIn

| Member | Meaning |
| --- | --- |
| `IRibbon Ribbon { get; }` | The host's ribbon. |
| `IWorkbook[] Workbooks { get; }` | The tracked workbooks. Interop: the workbooks Excel reported as new, opened or activated since the add-in started. Headless: every workbook created. |
| `string ExistentialAttribute { get; set; }` | The name of a custom document property. When set, the interop add-in tracks only workbooks that have that property. Headless: stored, not used. |
| `void DisplayAlerts(bool displayAlerts)` | Interop: sets Excel's `Application.DisplayAlerts`. Headless: throws `NotImplementedException`. |

`Allors.Excel.Interop.AddIn(InteropApplication application, IProgram program, IRibbon ribbon)`
subscribes to Excel's `NewWorkbook`, `WorkbookOpen`, `WorkbookActivate`, `WorkbookDeactivate` and
`WorkbookBeforeClose` events; it also exposes `Application`, `Program`,
`WorkbookByInteropWorkbook`, `New(InteropWorkbook)` and `Close(InteropWorkbook)`.

`Allors.Excel.Headless.AddIn(IProgram program, IRibbon ribbon)` creates one workbook with one
worksheet and calls `OnNew` for the worksheet, `OnNew` for the workbook and then `OnStart`; it
also exposes `Program`, a settable `Ribbon`, `WorkbookList`, `AddWorkbook()` and
`Remove(Workbook)`.

## IProgram

| Member | Called by |
| --- | --- |
| `Task OnStart(IAddIn addIn)` | The host, after it has created the add-in; the headless add-in, last in its constructor. |
| `Task OnStop()` | The host, at shutdown. |
| `Task OnNew(IWorkbook workbook)` | The add-in, for a workbook Excel reports as new or opened; the headless add-in, for every workbook it creates. |
| `void OnClose(IWorkbook workbook, ref bool cancel)` | The interop add-in, before Excel closes a tracked workbook. |
| `Task OnNew(IWorksheet worksheet)` | The add-in, for every worksheet it starts to track. |
| `Task OnBeforeDelete(IWorksheet worksheet)` | The interop add-in, before Excel deletes a sheet the user added after tracking began. |
| `Task OnHandle(string handle, params object[] argument)` | The host's ribbon callbacks. |
| `Task OnLogin()`, `Task OnLogout()` | The host. |
| `bool IsEnabled(string controlId, string controlTag)` | The host's `getEnabled` callbacks. |

Some callers await the returned tasks and others discard them;
[The lifecycle of a program](virtual-workbook.md#the-lifecycle-of-a-program) identifies each path
and its ordering and exception behaviour.

## IRibbon

| Member | Meaning |
| --- | --- |
| `void Invalidate()` | Asks Excel to query the ribbon's callbacks again. The hosts in the repository forward it to `IRibbonUI.Invalidate()`. |

## IWorkbook

| Member | Meaning |
| --- | --- |
| `event EventHandler<Hyperlink> OnHyperlinkClicked` | Raised by `HyperlinkClicked`. |
| `bool IsActive { get; }` | Interop: follows Excel's `WorkbookActivate` and `WorkbookDeactivate`. Headless: set by `Activate()`, which deactivates the other workbooks. |
| `IWorksheet[] Worksheets { get; }` | The tracked worksheets, in tracking order. |
| `IWorksheet[] WorksheetsByIndex { get; }` | The same worksheets ordered by `Index`. |
| `IBuiltinProperties BuiltinProperties { get; }` | The workbook's built-in document properties. |
| `ICustomProperties CustomProperties { get; }` | The workbook's custom document properties. |
| `void HyperlinkClicked(Hyperlink hyperlink)` | Interop: raises `OnHyperlinkClicked`. Headless: throws `NotImplementedException`. |
| `void Close(bool? saveChanges = null, string fileName = null)` | Interop: closes Excel's workbook with those arguments. Headless: removes the workbook from the add-in. |
| `IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null)` | Adds a worksheet and calls `OnNew` for it. `index` is the one-based index the new sheet gets: `0` inserts before the active sheet, an index beyond the last sheet appends; `before` and `after` place it next to another sheet; with no argument it is appended. |
| `IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet)` | Interop: copies `source` before `beforeWorksheet` and tracks the copy. Headless: throws `NotImplementedException`. |
| `Range[] GetNamedRanges(string refersToSheetName = null)` | Interop: every name of the workbook whose reference is on a tracked worksheet, as a `Range` with zero-based coordinates, the worksheet and the name; a worksheet-scoped name appears as `Sheet!Name`. With `refersToSheetName`, only the names referring to that sheet. Headless: every name set through `SetNamedRange`, ignoring the argument. |
| `void SetNamedRange(string name, Range range)` | Adds or updates a workbook-scoped name for `range`, whose `Worksheet` must be set. Interop: errors from Excel are swallowed, so a name that could not be created is simply absent afterwards. Headless: stores the range and sets `range.Name`. |
| `string SetCustomXml(XmlDocument xmlDocument)` | Adds a custom XML part and returns its id. |
| `XmlDocument GetCustomXmlById(string id)` | The part with that id, or `null`. |
| `bool TryDeleteCustomXmlById(string id)` | Deletes the part. Interop: `false` when Excel reports an error; an id that does not exist throws `NullReferenceException`. Headless: `false` when the id is unknown. |

## IWorksheet

| Member | Meaning |
| --- | --- |
| `event EventHandler<CellChangedEvent> CellsChanged`, `SheetChanged` | Both raised, with the same cells, when Excel reports changed values; see [Reading: what Excel reports](virtual-workbook.md#reading-what-excel-reports). Never raised headless. |
| `event EventHandler<string> SheetActivated` | Interop: raised with the sheet's name when Excel activates it. |
| `event EventHandler<Hyperlink> HyperlinkClicked` | Interop: raised when the user follows a hyperlink on the sheet, with `Address`, `DisplayName` and `SubAddress` filled from Excel's hyperlink. |
| `IWorkbook Workbook { get; }` | The owning workbook. |
| `bool IsVisible { get; set; }` | Interop: Excel's sheet visibility, visible or hidden. |
| `bool HasFreezePanes { get; }` | `true` between `FreezePanes` and `UnfreezePanes`; the library's record, not read from Excel. |
| `ICustomProperties CustomProperties { get; }` | Interop: the worksheet's custom properties. Headless: `null`. |
| `string Name { get; set; }` | The sheet's name. |
| `int Index { get; }` | The one-based position of the sheet in the workbook. |
| `bool IsActive { get; set; }` | Interop: setting `true` activates the sheet in Excel; Excel's `Activate` and `Deactivate` events keep it current. Headless: setting `true` deactivates the other sheets; setting `false` is ignored. |
| `ICell this[int row, int column]`, `ICell this[(int, int) coordinates]` | The cell at those zero-based coordinates, created on first use. |
| `IRow Row(int index)`, `IColumn Column(int index)` | The row or column, created on first use. A negative index throws `ArgumentException`. |
| `void Flush()` | Interop: writes every dirty change to Excel; see [Writing: dirty cells and Flush](virtual-workbook.md#writing-dirty-cells-and-flush). Headless: nothing. |
| `void RefreshPivotTables()` | Interop: refreshes every pivot table of the sheet. Headless: nothing. |
| `void AddPicture(string uri, Rectangle rectangle)` | Interop: inserts the picture file `uri` at the rectangle's position and size, in points, and then deletes the file. Headless: nothing. |
| `void AddHyperLink(string uri, ICell cell)` | Interop: at the next `Flush`, adds a hyperlink without an address to the cell, displaying `uri` as its text; clicking it raises `HyperlinkClicked`. Headless: nothing. |
| `Rectangle GetRectangle(string namedRange)` | The position and size of a worksheet-scoped named range. Interop: `Left`, `Top`, `Width` and `Height` in points, of the merge area when the range is merged. Headless: `X` and `Y` are the column and row, `Width` and `Height` the counts, in cells. Both throw `ArgumentException` for an unknown name. |
| `Range[] GetNamedRanges()` | The worksheet-scoped names, as for the workbook. |
| `void SetNamedRange(string name, Range range)` | Adds or updates a worksheet-scoped name; interop swallows Excel's errors. |
| `void InsertRows(int startRowIndex, int numberOfRows)` | Inserts rows below row `startRowIndex`; the rows after it move down, and the cells with them. Interop ignores a negative index or a count of zero. |
| `void DeleteRows(int startRowIndex, int numberOfRows)` | Deletes rows `startRowIndex` to `startRowIndex + numberOfRows - 1`; the rows below move up. Headless throws `ArgumentException` for a negative argument. |
| `void InsertColumns(int startColumnIndex, int numberOfColumns)` | Inserts columns to the right of column `startColumnIndex`. |
| `void DeleteColumns(int startColumnIndex, int numberOfColumns)` | Deletes columns `startColumnIndex` to `startColumnIndex + numberOfColumns - 1`. |
| `Range GetRange(string cell1, string cell2 = null)` | A zero-based range from A1-style references, with `Worksheet` set. Interop: whatever Excel resolves, such as `A1`, `A1:C5`, `A:A` or two references; `null` when Excel rejects it or `cell1` is blank. Headless: a limited parser that can return `null` or throw; see [Headless range parsing](#headless-range-parsing). |
| `Range GetUsedRange()` | Interop: Excel's used range. Headless: the block from the first to the last cell that exists; throws `InvalidOperationException` when no cell exists. |
| `void AutoFit()` | Interop: fits the columns of the used range. Headless: nothing. |
| `void SetChartObjectDataLabels(object chartObject, int seriesIndex, object seriesXValues, object seriesValues, bool showValues, bool showRange, string chartFieldRange, float baselineOffset, bool bold, bool italic, Color foreColor, object size)` | Interop: sets the values and data labels of a series of the chart object named or numbered by `chartObject`, then refreshes the chart. Headless: nothing. |
| `void SetChartObjectSourceData(object chartObject, object pivotTable)` | Interop: sets the chart's source to the pivot table's range and refreshes. Headless: nothing. |
| `void FreezePanes(Range range)` | Interop: activates the sheet and freezes panes. When `Row` and `Column` are both greater than zero, `Row` rows and `Column` columns are frozen, so the cell at `(Row, Column)` is the first that scrolls. Otherwise a `Row` of `-1` freezes no rows and a `Row` of *n* freezes the rows up to and including *n*; the same for `Column`. Headless: records the range. |
| `void UnfreezePanes()` | Removes the frozen panes. |
| `void SaveAsPdf(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)` | Interop: exports the sheet with Excel's fixed-format export at standard quality; the file's extension becomes `.pdf`, its directory is created, and an existing file throws `IOException` unless `overwriteExistingFile`. Headless: writes a one-byte placeholder file with the same rules. Both throw `ArgumentNullException` for a `null` file. |
| `void SaveAsXps(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)` | The same for `.xps`. |
| `void SetPrintArea(Range range = null)` | Interop: sets the print area to the range in A1 style, or clears it for `null`. Headless: nothing. |
| `void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true)` | Interop: replaces the cell's validation with an input-only validation that shows the message, which removes a list validation set through `Options`. Headless: nothing. |
| `void HideInputMessage(ICell cell, bool clearInputMessage = false)` | Interop: stops showing the input message; with `clearInputMessage`, clears the validation's error message and title, not the input message. Headless: nothing. |
| `void SetPageSetup(PageSetup pageSetup)` | Interop: sets the orientation, the paper size when it is 1 to 41, and the header and footer margins and texts. Headless: nothing. |

### Headless range parsing

The headless `GetRange` parser accepts single references such as `A1`, colon ranges such as
`A1:C5` with one letter and one digit on each side, whole-column ranges such as `A:A` and `A:C`,
and two separate references. It does not reliably return `null` for input it cannot parse.
With `cell2` omitted, three-character references such as `A10` and `AA1` throw
`NullReferenceException`; pass the reference twice, such as `GetRange("A10", "A10")`, to avoid
that defect. A row number too large for an `int` can also throw `OverflowException`.

## ICell, IRow and IColumn

| Member | Meaning |
| --- | --- |
| `IWorksheet Worksheet`, `IRow Row`, `IColumn Column` | Where the cell is. Interop: after rows or columns are inserted or deleted, a moved cell's `Row` and `Column` reflect its new position. Headless: the worksheet lookup moves the cell, but its `Row` and `Column` keep their original positions. |
| `object Value { get; set; }` | The value. Interop: marks the cell dirty when first set, `null` included, and afterwards when it changes; at `Flush`, written as the cell's value. |
| `string ValueAsString { get; }` | `Value` converted to text in the current culture. |
| `string Formula { get; set; }` | The formula text, such as `=SUM(A1:A3)`. Interop: dirty when first set or changed; written after the values, so it wins over a value set on the same cell. Headless: stored, not calculated. |
| `Range Options { get; set; }` | The range whose values are the cell's list validation, by its `Name` or by its coordinates on its `Worksheet`; one row or one column. Interop: dirty when the instance changes; `null` removes the validation. |
| `bool IsRequired { get; set; }` | Interop: when the options are written, an empty cell is invalid. Changing it alone does not mark the cell dirty; assign a new `Options` instance to apply it. |
| `bool HideInCellDropdown { get; set; }` | Interop: when the options are written, the drop-down arrow is hidden. Same remark. |
| `string Comment { get; set; }` | Interop: dirty when changed; written as the cell's comment, created or updated. |
| `Style Style { get; set; }` | Interop: dirty when the colours change; the background colour is written, `null` restores the automatic colour. The text colour is not written. |
| `string NumberFormat { get; set; }` | An Excel number format such as `0.00`. Interop: dirty when changed. |
| `IValueConverter ValueConverter { get; set; }` | Interop: converts the values Excel reports; `DefaultValueConverter` unless set. Headless: stored. |
| `void Clear()` | Interop: sets `Value` and `Formula` to the empty string and `Style` and `NumberFormat` to `null`, all dirty. Headless: sets `Value` to the empty string and `Style` and `NumberFormat` to `null`, and keeps `Formula`. |
| `IRow`: `IWorksheet Worksheet`, `int Index`, `bool Hidden { get; set; }` | Interop: setting `Hidden` marks the row dirty; at `Flush`, runs of rows get `EntireRow.Hidden`. |
| `IColumn`: `IWorksheet Worksheet`, `int Index` | |

## Range, Style and other values

| Type | Members |
| --- | --- |
| `Range(int row, int column, int? rows = null, int? columns = null, IWorksheet worksheet = null, string name = null)` | Zero-based `Row` and `Column`, the counts `Rows` and `Columns`, the `Worksheet` and a settable `Name`. The constructor throws `ArgumentException` unless `rows`, `columns` or `name` is given. `ToString()` prints `Sheet!Name (row:…, col:…, rows:…, cols:…)`. |
| `Style(Color backgroundColor, Color textColor)` | `BackgroundColor` and `TextColor`; two styles are equal when both colours are equal. |
| `PageSetup` | `Orientation` (1 portrait, the default, or 2 landscape), `PaperSize` (Excel's `XlPaperSize` number, default 9 for A4), `Header` and `Footer`. |
| `PageHeaderFooter` | `Margin` in points and the `Left`, `Center` and `Right` texts, in Excel's header and footer code. |
| `Hyperlink` | `Address`, `DisplayName`, `Subject` and `SubAddress`. |
| `CellChangedEvent(ICell[] cells)` | `Cells`: the cells whose values changed. |

## Document properties

`IBuiltinProperties` exposes Excel's built-in document properties as typed properties: `Title`,
`Subject`, `Author`, `Keywords`, `Comments`, `Template`, `LastAuthor`, `RevisionNumber` (read
only), `ApplicationName`, `LastPrintDate`, `CreationDate`, `LastSaveTime`, `TotalEditingTime`,
`NumberOfPages`, `NumberOfWords`, `NumberOfCharacters`, `Security`, `Category`, `Format`,
`Manager`, `Company`, `NumberOfBytes`, `NumberOfLines`, `NumberOfParagraphs`, `NumberOfSlides`,
`NumberOfNotes`, `NumberOfHiddenSlides`, `NumberOfMultimediaClips`, `HyperlinkBase`,
`NumberOfCharactersWithSpaces`, `ContentType`, `ContentStatus`, `Language` and
`DocumentVersion`. Interop reads and writes Excel's properties; setting `null` deletes the
property. Headless stores the values.

`ICustomProperties` has `GetBoolean`, `GetDate`, `GetFloat`, `GetNumber` and `GetString`, the
matching `Set…` methods, which take a nullable value, and `Remove(string name)`. Interop creates a
property of the matching type when it does not exist, updates it when it does, and removes it for
`null`. Headless stores the values; `SetNumber` throws `ArgumentOutOfRangeException` at the limits
of `long`.

## Value conversion

`IValueConverter.Convert(ICell cell, object excelValue)` returns the value to store for a value
Excel reported. `DefaultValueConverter` chooses a conversion from the cell's previous value.
For a cell that held an `int`, a reported `double` is converted to an `int`, clamped to its
minimum or maximum when outside that range. For a cell that held a `DateTime`, a `double` is
converted from an OLE automation date. For a cell that held a `string`, `null` becomes the empty
string and any other value its text. Otherwise the reported value is returned, except for the
decimal case below.

For a cell that held a `decimal`, a reported `double` below `(double)decimal.MinValue` or above
`(double)decimal.MaxValue` becomes the corresponding bound as a `double`, not a `decimal`.
Otherwise the converter calls `Convert.ToDecimal`; the rounded boundary values themselves and
`double.NaN` throw `OverflowException`. A caller cannot assume that every reported number
preserves the cell's decimal type.

## Binding

`Binder(IWorksheet worksheet, Style changedStyle = null)` with `Set(int row, int column, IBinding binding)`,
`Set(ICell cell, IBinding binding)`, `ICell[] ToCells()`, `ResetChangedCells()`,
`bool ExistBinding(int row, int column)`, the `ToDomained` event and the `Worksheet` and
`ChangedStyle` properties; `Binding(Action<ICell> toCell = null, Action<ICell> toDomain = null)`
implements `IBinding` with `ToCell`, `ToDomain`, `OneWayBinding` and `TwoWayBinding`.
[Bind cells to your model](bind-cells.md) describes their behaviour.

## Utils

| Member | Meaning |
| --- | --- |
| `string Utils.ExcelColumnFromNumber(int column)` | The letters of a one-based column number: `1` is `A`, `27` is `AA`. |
| `int Utils.ExcelColumnFromLetter(string column)` | The one-based number of column letters. |

## Interop-only members

| Member | Meaning |
| --- | --- |
| `Workbook.InteropWorkbook`, `Worksheet.InteropWorksheet`, `AddIn.Application` | The COM objects behind the virtual ones. |
| `Worksheet.RowByIndex`, `Worksheet.ColumnByIndex` | The rows and columns created so far. |
| `Worksheet.PreventChangeEvent` | `true` while the library changes the sheet's structure, so that Excel's `Change` event is ignored. |
| `static int Worksheet.ToOle(Color c)` | A `Color` as Excel's BGR colour number. |
| `Cell.UpdateValue(object rawExcelValue)` | Converts and stores a value Excel reported, without marking the cell dirty; `true` when the stored value changed. |
| `Cells.Chunks(this IEnumerable<Cell> cells, Func<Cell, Cell, bool> combine)` | Groups cells into rectangular blocks of adjacent cells for which `combine` holds between neighbours; `ChunkTests` covers it. |
| `InteropRangeExtension.Coordinates(this InteropRange range)` | The zero-based `(row, column)` of an interop range. |
| `Allors.Excel.Interop.IWorksheet` | Extends `IWorksheet` with the `AddDirtyValue`, `AddDirtyFormula`, `AddDirtyComment`, `AddDirtyStyle`, `AddDirtyNumberFormat`, `AddDirtyOptions` and `AddDirtyRow` methods that cells and rows call. |
| `Constants` | Five `Magic…` identifiers that nothing in the library uses. |

## Headless and interop differences

| Member | Headless |
| --- | --- |
| `IAddIn.DisplayAlerts`, `IWorkbook.Copy`, `IWorkbook.HyperlinkClicked` | Throw `NotImplementedException`. |
| `IWorksheet.Flush`, `RefreshPivotTables`, `AddPicture`, `AddHyperLink`, `AutoFit`, `SetPrintArea`, `SetInputMessage`, `HideInputMessage`, `SetPageSetup`, `SetChartObjectDataLabels`, `SetChartObjectSourceData` | Do nothing. |
| `IWorksheet.CellsChanged`, `SheetChanged`, `SheetActivated`, `HyperlinkClicked`, `IWorkbook.OnHyperlinkClicked` | Never raised. |
| `ICell.Formula` | Stored as text; nothing is calculated. |
| `IWorksheet.CustomProperties` | `null`. |
| `IWorksheet.GetRectangle` | Cell coordinates and counts instead of points. |
| `IWorksheet.GetRange` | A limited parser; see [IWorksheet](#iworksheet). |
| `IWorksheet.GetUsedRange` | Throws on a sheet without cells. |
| `IWorksheet.SaveAsPdf`, `SaveAsXps` | A one-byte placeholder file. |
| `IWorkbook.GetNamedRanges(refersToSheetName)` | Ignores the filter. |
| `ICell.Clear` | Keeps the formula. |
| `ICell.Row`, `ICell.Column` | Keep their original positions when row or column operations move the cell in the worksheet lookup. |
| `IAddIn.ExistentialAttribute` | Stored, without effect. |
