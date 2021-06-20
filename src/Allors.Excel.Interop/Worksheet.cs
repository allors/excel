// <copyright file="Worksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using Rectangle = System.Drawing.Rectangle;
using XlColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex;

namespace Allors.Excel.Interop
{
    using System.Threading;
    using InteropDocEvents_Event = DocEvents_Event;
    using InteropName = Name;
    using InteropPivotTable = PivotTable;
    using InteropPivotTables = PivotTables;
    using InteropRange = Microsoft.Office.Interop.Excel.Range;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropXlCalculation = XlCalculation;
    using InteropXlDeleteShiftDirection = XlDeleteShiftDirection;
    using InteropXlDVAlertStyle = XlDVAlertStyle;
    using InteropXlDVType = XlDVType;
    using InteropXlFixedFormatQuality = XlFixedFormatQuality;
    using InteropXlFixedFormatType = XlFixedFormatType;
    using InteropXlInsertShiftDirection = XlInsertShiftDirection;
    using InteropXlSheetVisibility = XlSheetVisibility;

    public class Worksheet : IWorksheet
    {
        public bool isActive;

        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;

            this.RowByIndex = new Dictionary<int, Row>();
            this.ColumnByIndex = new Dictionary<int, Column>();
            this.CellByCoordinates = new Dictionary<(int, int), Cell>();
            this.DirtyValueCells = new HashSet<Cell>();
            this.DirtyCommentCells = new HashSet<Cell>();
            this.DirtyStyleCells = new HashSet<Cell>();
            this.DirtyOptionCells = new HashSet<Cell>();
            this.DirtyNumberFormatCells = new HashSet<Cell>();
            this.DirtyFormulaCells = new HashSet<Cell>();
            this.DirtyRows = new HashSet<Row>();

            interopWorksheet.Change += this.InteropWorksheet_Change;

            ((InteropDocEvents_Event)interopWorksheet).Activate += () =>
            {
                this.isActive = true;
                this.SheetActivated?.Invoke(this, this.Name);
            };

            interopWorksheet.Deactivate += () => this.isActive = false;

            this.CustomProperties = new CustomProperties(this.InteropWorksheet.CustomProperties);

            this.Reset();
        }

        public event EventHandler<CellChangedEvent> CellsChanged;

        public event EventHandler<string> SheetActivated;

        private Range FreezeRange { get; set; }

        public int Index => this.InteropWorksheet.Index;

        public bool IsActive
        {
            get => this.isActive;
            set
            {
                if (value)
                {
                    this.isActive = true;

                    this.InteropWorksheet.Activate();
                }
                else
                {
                    this.isActive = false;
                }
            }
        }

        public Workbook Workbook { get; set; }

        public ICustomProperties CustomProperties { get; }

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.Name; set => this.InteropWorksheet.Name = value; }

        IWorkbook Excel.IWorksheet.Workbook => this.Workbook;

        public Dictionary<int, Row> RowByIndex { get; set; }

        public Dictionary<int, Column> ColumnByIndex { get; set; }

        private Dictionary<(int, int), Cell> CellByCoordinates { get; }

        private HashSet<Cell> DirtyValueCells { get; set; }

        private HashSet<Cell> DirtyCommentCells { get; set; }

        private HashSet<Cell> DirtyStyleCells { get; set; }

        private HashSet<Cell> DirtyOptionCells { get; set; }

        private HashSet<Cell> DirtyNumberFormatCells { get; set; }

        private HashSet<Cell> DirtyFormulaCells { get; set; }

        private HashSet<Row> DirtyRows { get; set; }

        public bool PreventChangeEvent { get; private set; }

        public bool IsVisible
        {
            get => this.InteropWorksheet.Visible == InteropXlSheetVisibility.xlSheetVisible;
            set
            {
                if (value)
                {
                    this.InteropWorksheet.Visible = InteropXlSheetVisibility.xlSheetVisible;
                }
                else
                {
                    this.InteropWorksheet.Visible = InteropXlSheetVisibility.xlSheetHidden;
                }
            }
        }

        public async Task RefreshPivotTables()
        {
            var pivotTables = (InteropPivotTables)this.InteropWorksheet.PivotTables();

            foreach (InteropPivotTable pivotTable in pivotTables)
            {
                pivotTable.RefreshTable();
            }

            await Task.CompletedTask;
        }

        ICell Excel.IWorksheet.this[(int, int) coordinates] => this[coordinates];

        ICell Excel.IWorksheet.this[int row, int column] => this[(row, column)];

        public Cell this[(int, int) coordinates]
        {
            get
            {
                if (!this.CellByCoordinates.TryGetValue(coordinates, out var cell))
                {
                    cell = new Cell(this, this.Row(coordinates.Item1), this.Column(coordinates.Item2));
                    this.CellByCoordinates.Add(coordinates, cell);
                }

                return cell;
            }
        }

        public static string ExcelColumnFromNumber(int column)
        {
            var columnString = string.Empty;
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                var currentLetterNumber = (columnNumber - 1) % 26;
                var currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        public static int ExcelColumnFromLetter(string column)
        {
            var retVal = 0;
            var col = column.ToUpper();
            for (var iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                var colPiece = col[iChar];
                var colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        IRow Excel.IWorksheet.Row(int index)
        {
            return this.Row(index);
        }

        IColumn Excel.IWorksheet.Column(int index)
        {
            return this.Column(index);
        }

        public Row Row(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException("Index can not be negative", nameof(this.Row));
            }

            if (!this.RowByIndex.TryGetValue(index, out var row))
            {
                row = new Row(this, index);
                this.RowByIndex.Add(index, row);
            }

            return row;
        }

        public Column Column(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException(nameof(this.Column));
            }

            if (!this.ColumnByIndex.TryGetValue(index, out var column))
            {
                column = new Column(this, index);
                this.ColumnByIndex.Add(index, column);
            }

            return column;
        }

        private Tuple<InteropXlCalculation, bool> DisableExcel()
        {
            var calculation = this.Workbook.InteropWorkbook.Application.Calculation;
            if (calculation != InteropXlCalculation.xlCalculationManual)
            {
                this.Workbook.InteropWorkbook.Application.Calculation = InteropXlCalculation.xlCalculationManual;
            }

            this.Workbook.InteropWorkbook.Application.ScreenUpdating = false;
            this.Workbook.InteropWorkbook.Application.EnableEvents = false;
            this.Workbook.InteropWorkbook.Application.DisplayStatusBar = false;
            this.Workbook.InteropWorkbook.Application.PrintCommunication = false;

            var enableFormatConditionsCalculation = this.InteropWorksheet.EnableFormatConditionsCalculation;

            if (enableFormatConditionsCalculation)
            {
                this.InteropWorksheet.EnableFormatConditionsCalculation = false;
            }

            return Tuple.Create(calculation, enableFormatConditionsCalculation);
        }

        private void EnableExcel(Tuple<InteropXlCalculation, bool> tuple)
        {
            this.Workbook.InteropWorkbook.Application.Calculation = tuple.Item1;
            this.Workbook.InteropWorkbook.Application.ScreenUpdating = true;
            this.Workbook.InteropWorkbook.Application.EnableEvents = true;
            this.Workbook.InteropWorkbook.Application.DisplayStatusBar = true;
            this.Workbook.InteropWorkbook.Application.PrintCommunication = true;

            this.InteropWorksheet.EnableFormatConditionsCalculation = tuple.Item2;

            try
            {
                // Recalculate when required. Formulas need to be resolved.
                if (tuple.Item1 == InteropXlCalculation.xlCalculationAutomatic)
                {
                    this.InteropWorksheet.Calculate();
                }
            }
            catch
            {
            }
        }
        public async Task Flush()
        {
            var tuple = this.DisableExcel();

            try
            {
                this.RenderNumberFormat(this.DirtyNumberFormatCells);
                this.DirtyNumberFormatCells = new HashSet<Cell>();

                this.RenderValue(this.DirtyValueCells);
                this.DirtyValueCells = new HashSet<Cell>();

                this.RenderFormula(this.DirtyFormulaCells);
                this.DirtyFormulaCells = new HashSet<Cell>();

                this.RenderComments(this.DirtyCommentCells);
                this.DirtyCommentCells = new HashSet<Cell>();

                this.RenderStyle(this.DirtyStyleCells);
                this.DirtyStyleCells = new HashSet<Cell>();

                this.SetOptions(this.DirtyOptionCells);
                this.DirtyOptionCells = new HashSet<Cell>();

                this.UpdateRows(this.DirtyRows);
                this.DirtyRows = new HashSet<Row>();
            }
            finally
            {
                this.EnableExcel(tuple);
            }

            await Task.CompletedTask;
        }

        public void AddDirtyNumberFormat(Cell cell)
        {
            this.DirtyNumberFormatCells.Add(cell);
        }

        public void AddDirtyValue(Cell cell)
        {
            this.DirtyValueCells.Add(cell);
        }

        public void AddDirtyFormula(Cell cell)
        {
            this.DirtyFormulaCells.Add(cell);
        }

        public void AddDirtyComment(Cell cell)
        {
            this.DirtyCommentCells.Add(cell);
        }

        public void AddDirtyStyle(Cell cell)
        {
            this.DirtyStyleCells.Add(cell);
        }

        public void AddDirtyOptions(Cell cell)
        {
            this.DirtyOptionCells.Add(cell);
        }

        public void AddDirtyRow(Row row)
        {
            this.DirtyRows.Add(row);
        }

        private void InteropWorksheet_Change(InteropRange target)
        {
            if (this.PreventChangeEvent)
            {
                return;
            }

            if (target.Cells.Count >= this.InteropWorksheet.Columns.Count)
            {
                // (probably) row or rolumn insert(s)
                this.Reset();
            }
            else
            {
                List<Cell> cells = null;
                foreach (InteropRange targetCell in target.Cells)
                {
                    var cell = this[targetCell.Coordinates()];
                    if (cell.UpdateValue(targetCell.Value2))
                    {
                        if (cells == null)
                        {
                            cells = new List<Cell>();
                        }

                        cells.Add(cell);
                    }
                }

                if (cells != null)
                {
                    this.CellsChanged?.Invoke(this, new CellChangedEvent(cells.Cast<ICell>().ToArray()));
                }
            }
        }

        private void RenderValue(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => true);

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var values = new object[chunk.Count, chunk[0].Count];
                    for (var i = 0; i < chunk.Count; i++)
                    {
                        for (var j = 0; j < chunk[0].Count; j++)
                        {
                            values[i, j] = chunk[i][j].Value;
                        }
                    }

                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var range = this.WaitAndRetry(() =>
                    {
                        var from = (InteropRange)this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = (InteropRange)this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        range.Value2 = values;
                    });
                });
        }

        private void RenderFormula(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => true);

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var formulas = new object[chunk.Count, chunk[0].Count];
                    for (var i = 0; i < chunk.Count; i++)
                    {
                        for (var j = 0; j < chunk[0].Count; j++)
                        {
                            formulas[i, j] = chunk[i][j].Formula;
                        }
                    }

                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var range = this.WaitAndRetry(() =>
                    {
                        var from = (InteropRange)this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = (InteropRange)this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        range.Formula = formulas;
                    });
                });
        }

        private void RenderComments(IEnumerable<Cell> cells)
        {
            Parallel.ForEach(
                cells,
                cell =>
                {
                    var range = this.WaitAndRetry(() =>
                    {
                        return (InteropRange)this.InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];
                    });

                    this.WaitAndRetry(() =>
                    {
                        if (range.Comment == null)
                        {
                            var comment = range.AddComment(cell.Comment);
                            comment.Shape.TextFrame.AutoSize = true;
                        }
                        else
                        {
                            range.Comment.Text(cell.Comment);
                        }
                    });
                });
        }

        private void RenderStyle(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.Style, w.Style));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var range = this.WaitAndRetry(() =>
                    {
                        var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        var cc = chunk[0][0];
                        if (cc.Style != null)
                        {
                            range.Interior.Color = ColorTranslator.ToOle(chunk[0][0].Style.BackgroundColor);
                        }
                        else
                        {
                            range.Interior.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                        }
                    });
                });
        }

        private void RenderNumberFormat(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var range = this.WaitAndRetry(() =>
                    {
                        var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        range.NumberFormat = chunk[0][0].NumberFormat;
                    });
                });
        }

        private void SetOptions(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.Options, w.Options));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var range = this.WaitAndRetry(() =>
                    {
                        var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        var cc = chunk[0][0];
                        if (cc.Options != null)
                        {
                            var validationRange = cc.Options.Name;
                            if (string.IsNullOrEmpty(validationRange))
                            {
                                if (cc.Options.Columns.HasValue)
                                {
                                    validationRange = $"{cc.Options.Worksheet.Name}!${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + 1}:${ExcelColumnFromNumber(cc.Options.Column + cc.Options.Columns.Value)}${cc.Options.Row + 1}";
                                }
                                else if (cc.Options.Rows.HasValue)
                                {
                                    validationRange = $"{cc.Options.Worksheet.Name}!${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + 1}:${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + cc.Options.Rows}";
                                }
                            }

                            try
                            {
                                range.Validation.Delete();
                            }
                            catch (Exception)
                            {
                            }

                            range.Validation.Add(InteropXlDVType.xlValidateList, InteropXlDVAlertStyle.xlValidAlertStop, Type.Missing, $"={validationRange}", Type.Missing);
                            range.Validation.IgnoreBlank = !cc.IsRequired;
                            range.Validation.InCellDropdown = !cc.HideInCellDropdown;
                        }
                        else
                        {
                            try
                            {
                                range.Validation.Delete();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    });
                });
        }

        private void UpdateRows(HashSet<Row> dirtyRows)
        {
            var chunks = dirtyRows.OrderBy(w => w.Index).Aggregate(
                        new List<IList<Row>> { new List<Row>() },
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || (list[list.Count - 1].Hidden == w.Hidden))
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<Row> { w };
                                acc.Add(list);
                            }

                            return acc;
                        });

            var updateChunks = chunks.Where(v => v.Count > 0);

            Parallel.ForEach(
                updateChunks,
                chunk =>
                {
                    var fromChunk = chunk.First();
                    var toChunk = chunk.Last();
                    var hidden = fromChunk.Hidden;

                    var from = $"$A${fromChunk.Index + 1}";
                    var to = $"$A${toChunk.Index + 1}";

                    var range = this.WaitAndRetry(() =>
                    {
                        return this.InteropWorksheet.Range[from, to];
                    });

                    this.WaitAndRetry(() =>
                    {
                        range.EntireRow.Hidden = hidden;
                    });
                });
        }

        /// <summary>
        /// Adds a Picture on the specified rectangle. <seealso cref="GetRectangle(string)"/>
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="location"></param>
        /// <param name="size"></param>
        public void AddPicture(string fileName, Rectangle rectangle)
        {
            this.InteropWorksheet.Shapes.AddPicture(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);

            try
            {
                File.Delete(fileName);
            }
            catch
            {
                // left blank: delete temp file may fail.
            }
        }

        /// <summary>
        /// Gets the Rectangle of a namedRange (uses the Range.MergeArea as reference). 
        /// NamedRange must exist on Workbook.
        /// </summary>
        /// <param name="namedRange"></param>
        /// <returns></returns>
        public Rectangle GetRectangle(string namedRange)
        {
            InteropName name = null;

            try
            {
                // Get the Namedrange that is scoped to this worksheet
                name = this.InteropWorksheet.Names.Item(namedRange);
            }
            catch
            {
                throw new ArgumentException("Name not found for namedRange", nameof(namedRange));
            }


            // when the range is a mergedrange, then take the values from the area
            var range = name.RefersToRange;

            if ((bool)range.MergeCells == false)
            {
                var left = Convert.ToInt32(range.Left);
                var top = Convert.ToInt32(range.Top);
                var width = Convert.ToInt32(range.Width);
                var height = Convert.ToInt32(range.Height);

                return new Rectangle(left, top, width, height);
            }
            else
            {
                var area = range.MergeArea;

                var left = Convert.ToInt32(area.Left);
                var top = Convert.ToInt32(area.Top);
                var width = Convert.ToInt32(area.Width);
                var height = Convert.ToInt32(area.Height);

                return new Rectangle(left, top, width, height);
            }

        }

        public Range[] GetNamedRanges()
        {
            var ranges = new List<Range>();

            foreach (Name namedRange in this.InteropWorksheet.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {
                        ranges.Add(new Range(refersToRange.Row - 1, refersToRange.Column - 1, refersToRange.Rows.Count, refersToRange.Columns.Count, worksheet: this, name: namedRange.Name));
                    }
                }
                catch
                {
                    // RefersToRange can throw exception
                }
            }

            return ranges.ToArray();
        }

        /// <summary>
        /// Adds a NamedRange that has its scope on the Worksheet.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        public void SetNamedRange(string name, Range range)
        {
            if (!string.IsNullOrWhiteSpace(name) && range != null)
            {
                try
                {
                    var topLeft = this.InteropWorksheet.Cells[range.Row + 1, range.Column + 1];
                    var bottomRight = this.InteropWorksheet.Cells[range.Row + range.Rows, range.Column + range.Columns];

                    var refersTo = this.InteropWorksheet.Range[topLeft, bottomRight];

                    // When it does not exist, add it, else we update the range.
                    if (this.InteropWorksheet.Names
                        .Cast<Name>()
                        .Any(v => string.Equals(v.Name, name)))
                    {
                        this.InteropWorksheet.Names.Item(name).RefersTo = refersTo;
                    }
                    else
                    {
                        this.InteropWorksheet.Names.Add(name, refersTo);
                    }
                }
                catch
                {
                    // can throw exception, we dont care.
                }
            }
        }

        public void InsertRows(int startRowIndex, int numberOfRows)
        {
            if (startRowIndex >= 0 && numberOfRows > 0)
            {
                this.PreventChangeEvent = true;

                try
                {
                    var rows = this.InteropWorksheet.Range[$"{startRowIndex + 2}:{startRowIndex + numberOfRows + 1}"];

                    this.WaitAndRetry(() =>
                    {
                        var tuple = this.DisableExcel();
                        rows.Insert(InteropXlInsertShiftDirection.xlShiftDown);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByCoordinates.Any())
                    {
                        // Shift all cell rows down with the numberOfRows
                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (var item in this.CellByCoordinates
                        .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                        .OrderByDescending(kvp => kvp.Value.Row.Index)
                        .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByCoordinates.Remove(item.Key);

                            // Make new or use existing row
                            var row = this.Row(item.Value.Row.Index + numberOfRows);
                            cell.Row = row;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            this.CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    this.PreventChangeEvent = false;
                }
            }
        }

        public void DeleteRows(int startRowIndex, int numberOfRows)
        {
            if (startRowIndex >= 0 && numberOfRows > 0)
            {
                this.PreventChangeEvent = true;

                try
                {
                    var rows = this.InteropWorksheet.Range[$"{startRowIndex + 1}:{startRowIndex + numberOfRows}"];

                    this.WaitAndRetry(() =>
                    {
                        var tuple = this.DisableExcel();
                        rows.Delete(InteropXlDeleteShiftDirection.xlShiftUp);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByCoordinates.Any())
                    {
                        // Delete all cells in the deleted rows.
                        foreach (var rowIndex in Enumerable.Range(startRowIndex, numberOfRows))
                        {
                            foreach (var item in this.CellByCoordinates
                              .Where(kvp => kvp.Value.Row.Index == rowIndex)
                              .ToList())
                            {
                                this.CellByCoordinates.Remove(item.Key);
                            }
                        }

                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (var item in this.CellByCoordinates
                            .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                            .OrderBy(kvp => kvp.Value.Row.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByCoordinates.Remove(item.Key);

                            var rowIndex = cell.Row.Index - numberOfRows;

                            // Link the cell to the new Row that already exists
                            cell.Row = this.Row(rowIndex);

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            this.CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    this.PreventChangeEvent = false;
                }
            }
        }

        public void InsertColumns(int startColumnIndex, int numberOfColumns)
        {
            if (startColumnIndex >= 0 && numberOfColumns > 0)
            {
                this.PreventChangeEvent = true;

                try
                {
                    var startColumnName = ExcelColumnFromNumber(startColumnIndex + 2);
                    var endColumnName = ExcelColumnFromNumber(startColumnIndex + 1 + numberOfColumns);

                    var rows = this.InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    this.WaitAndRetry(() =>
                    {
                        var tuple = this.DisableExcel();
                        rows.EntireColumn.Insert(InteropXlInsertShiftDirection.xlShiftToRight);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByCoordinates.Any())
                    {
                        // Shift all cell columns to the right with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (var item in this.CellByCoordinates
                                    .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                                    .OrderByDescending(kvp => kvp.Value.Column.Index)
                                    .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByCoordinates.Remove(item.Key);

                            var column = this.Column(item.Value.Column.Index + numberOfColumns);

                            // Shift rows up with the numberofrows that were deleted.
                            cell.Column = column;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            this.CellByCoordinates[coordinates] = cell;
                        }
                    }
                }
                finally
                {
                    this.PreventChangeEvent = false;
                }
            }
        }

        public void DeleteColumns(int startColumnIndex, int numberOfColumns)
        {
            if (startColumnIndex >= 0 && numberOfColumns > 0)
            {
                this.PreventChangeEvent = true;

                try
                {
                    var startColumnName = ExcelColumnFromNumber(startColumnIndex + 1);
                    var endColumnName = ExcelColumnFromNumber(startColumnIndex + numberOfColumns);

                    var range = this.InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    this.WaitAndRetry(() =>
                    {
                        var tuple = this.DisableExcel();
                        range.EntireColumn.Delete(InteropXlDeleteShiftDirection.xlShiftToLeft);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByCoordinates.Any())
                    {
                        // Delete all cells in the deleted columns.
                        foreach (var columnIndex in Enumerable.Range(startColumnIndex, numberOfColumns))
                        {
                            foreach (var item in this.CellByCoordinates
                                .Where(kvp => kvp.Value.Column.Index == columnIndex)
                                .ToList())
                            {
                                this.CellByCoordinates.Remove(item.Key);
                            }
                        }

                        // Shift all cell Columns to the left with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (var item in this.CellByCoordinates
                            .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                            .OrderBy(kvp => kvp.Value.Column.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByCoordinates.Remove(item.Key);

                            var columnIndex = cell.Column.Index - numberOfColumns;

                            var column = this.Column(columnIndex);

                            // Link to the correct column that already exists.
                            cell.Column = column;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            this.CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    this.PreventChangeEvent = false;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell1">The name of the range in A1-style notation - "A1" - "A1:C5", "A", "A:C"</param>
        /// <param name="cell2">The cell in the lower-right corner of the range</param>
        /// <returns></returns>
        public Range GetRange(string cell1, string cell2 = null)
        {
            if (string.IsNullOrWhiteSpace(cell1) && cell2 == null)
            {
                return null;
            }

            try
            {
                InteropRange interopRange;
                if (cell2 == null)
                {
                    interopRange = this.InteropWorksheet.Range[cell1];
                }
                else
                {
                    interopRange = this.InteropWorksheet.Range[cell1, cell2];
                }


                return new Range(interopRange.Row - 1, interopRange.Column - 1, interopRange.Rows.Count, interopRange.Columns.Count, this);
            }
            catch
            {
                return null;
            }


        }

        public Range GetUsedRange()
        {
            var range = this.InteropWorksheet.UsedRange;

            return new Range(range.Row - 1, range.Column - 1, range.Rows.Count, range.Columns.Count, this);
        }

        public Range GetUsedRange(int row)
        {
            if (row < 0 || row >= this.InteropWorksheet.UsedRange.Row + this.InteropWorksheet.UsedRange.Rows.Count)
            {
                return null;
            }

            var rowRange = (InteropRange)this.InteropWorksheet.Rows[row + 1];

            var endColumnIndex = this.InteropWorksheet.UsedRange.Column + this.InteropWorksheet.UsedRange.Columns.Count - 1;
            var quit = false;

            do
            {
                var cell = (InteropRange)this.InteropWorksheet.Cells[rowRange.Row, endColumnIndex];

                if (cell.Value2 == null)
                {
                    endColumnIndex--;
                }
                else
                {
                    quit = true;
                }
            }
            while (!quit && endColumnIndex >= rowRange.Column); // do not read passed the left of the rowRange.Column

            var beginColumnIndex = rowRange.Column;
            quit = false;

            do
            {
                var cell = (InteropRange)this.InteropWorksheet.Cells[rowRange.Row, beginColumnIndex];

                if (cell.Value2 == null)
                {
                    beginColumnIndex++;
                }
                else
                {
                    quit = true;
                }
            }
            while (!quit && beginColumnIndex <= endColumnIndex);

            // B..D => 3 Columns B,C,D
            // 3 = 1 + 4 - 2
            var columnCount = 1 + endColumnIndex - beginColumnIndex;

            return new Range(rowRange.Row - 1, beginColumnIndex - 1, rowRange.Rows.Count, columnCount, this);
        }

        public Range GetUsedRange(string column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return null;
            }

            var columnIndex = ExcelColumnFromLetter(column);
            var columnRange = (InteropRange)this.InteropWorksheet.Columns[columnIndex];

            var beginRowIndex = columnRange.Row;
            var maxRows = this.InteropWorksheet.UsedRange.Rows.Count;
            var quit = false;

            do
            {
                var cell = (InteropRange)this.InteropWorksheet.Cells[beginRowIndex, columnRange.Column];

                if (cell.Value2 == null)
                {
                    beginRowIndex++;
                }
                else
                {
                    quit = true;
                }
            }
            while (!quit || beginRowIndex >= maxRows);


            var endRowIndex = this.InteropWorksheet.UsedRange.Row + this.InteropWorksheet.UsedRange.Rows.Count - 1;
            quit = false;

            do
            {
                var cell = (InteropRange)this.InteropWorksheet.Cells[endRowIndex, columnRange.Column];

                if (cell.Value2 == null)
                {
                    endRowIndex--;
                }
                else
                {
                    quit = true;
                }
            }
            while (!quit && endRowIndex >= columnRange.Row); // do not read passed the top of the columnRange.Row

            var rowCount = 1 + endRowIndex - beginRowIndex;

            return new Range(beginRowIndex - 1, columnRange.Column - 1, rowCount, columnRange.Columns.Count, this);
        }

        public void AutoFit()
        {
            var usedRange = this.InteropWorksheet.UsedRange;
            var columns = usedRange.Columns;
            columns.AutoFit();
        }

        /// <inheritdoc/>
        /// <summary>
        /// When range.Row = 0 and range.Column = -1, then topRow in frozen
        /// When range.Row = -1 and range.Column = 0 then leftColumn in  frozen
        /// When range.Row > 0 and range.Column > 0, then that cell is the topleft position for the freezepanes
        /// </summary>
        public void FreezePanes(Range range)
        {
            this.InteropWorksheet.Application.ScreenUpdating = true;
            this.InteropWorksheet.Activate();

            this.InteropWorksheet.Application.ActiveWindow.FreezePanes = false;

            if (range.Row > 0 && range.Column > 0)
            {
                this.InteropWorksheet.Application.ActiveWindow.SplitRow = range.Row;
                this.InteropWorksheet.Application.ActiveWindow.SplitColumn = range.Column;
            }
            else
            {
                var row = 0;
                if (range.Row > -1)
                {
                    row = range.Row + 1;
                }

                this.InteropWorksheet.Application.ActiveWindow.SplitRow = row;

                var column = 0;
                if (range.Column > -1)
                {
                    column = range.Column + 1;
                }

                this.InteropWorksheet.Application.ActiveWindow.SplitColumn = column;
            }

            this.InteropWorksheet.Application.ActiveWindow.FreezePanes = true;

            this.FreezeRange = range;
        }

        public void UnfreezePanes()
        {
            this.InteropWorksheet.Application.ScreenUpdating = true;
            this.InteropWorksheet.Activate();

            this.InteropWorksheet.Application.ActiveWindow.SplitRow = 0;
            this.InteropWorksheet.Application.ActiveWindow.SplitColumn = 0;
            this.InteropWorksheet.Application.ActiveWindow.FreezePanes = false;

            this.FreezeRange = null;
        }

        public bool HasFreezePanes => this.FreezeRange != null;

        public void SaveAsXPS(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            this.SaveAs(file, InteropXlFixedFormatType.xlTypeXPS, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
        }

        public void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            this.SaveAs(file, InteropXlFixedFormatType.xlTypePDF, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
        }

        /// <summary>
        /// Save the sheet in the given formattype (0=PDF, 1=XPS) 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="formatType"></param>
        /// <param name="overwriteExistingFile"></param>
        /// <param name="openAfterPublish"></param>
        /// <param name="ignorePrintAreas"></param>
        private void SaveAs(FileInfo fileInfo, InteropXlFixedFormatType formatType, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            if (fileInfo == null)
            {
                throw new ArgumentNullException(nameof(fileInfo));
            }

            fileInfo.Refresh();

            // In case we would overwrite an existing file
            if (fileInfo.Exists && !overwriteExistingFile)
            {
                throw new IOException($"File {fileInfo.FullName} already exists and should not be overwritten.");
            }

            if (formatType == InteropXlFixedFormatType.xlTypePDF && !string.Equals(fileInfo.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                fileInfo = new FileInfo(Path.ChangeExtension(fileInfo.FullName, ".pdf"));
            }

            if (formatType == InteropXlFixedFormatType.xlTypeXPS && !string.Equals(fileInfo.Extension, ".xps", StringComparison.OrdinalIgnoreCase))
            {
                fileInfo = new FileInfo(Path.ChangeExtension(fileInfo.FullName, ".xps"));
            }

            if (!Directory.Exists(fileInfo.DirectoryName))
            {
                Directory.CreateDirectory(fileInfo.DirectoryName);
            }

            this.InteropWorksheet
                .ExportAsFixedFormat
                (
                       Type: formatType,
                       Filename: fileInfo.FullName,
                       Quality: InteropXlFixedFormatQuality.xlQualityStandard,
                       IgnorePrintAreas: ignorePrintAreas,
                       OpenAfterPublish: openAfterPublish
               );
        }

        /// <inheritdoc/>
        public void SetPrintArea(Range range = null)
        {
            // Use A1-Style reference for the printarea

            var printArea = "";

            if (range != null)
            {
                // row 3, column 5, rows 6 column 2
                // => A1 Style = F4:H10
                var startColumn = ExcelColumnFromNumber(range.Column + 1); // !zero-based
                var startRow = range.Row + 1;

                var endColumn = ExcelColumnFromNumber(range.Column + range.Columns.GetValueOrDefault());
                var endRow = range.Row + range.Rows.GetValueOrDefault();

                printArea = $"{startColumn}{startRow}:{endColumn}{endRow}";
            }

            this.InteropWorksheet.PageSetup.PrintArea = printArea;
        }

        public void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true)
        {
            var inputCell = (InteropRange)this.InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];

            inputCell.Validation.Delete();
            inputCell.Validation.Add(InteropXlDVType.xlValidateInputOnly);
            inputCell.Validation.ShowInput = showInputMessage;

            inputCell.Validation.InputMessage = message;
            inputCell.Validation.InputTitle = title;
        }

        public void HideInputMessage(ICell cell, bool clearInputMessage = false)
        {
            var inputCell = (InteropRange)this.InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];

            if (clearInputMessage)
            {
                inputCell.Validation.ErrorMessage = null;
                inputCell.Validation.ErrorTitle = null;
            }

            inputCell.Validation.ShowInput = false;
        }

        private void Reset()
        {
            var deletedCells = new HashSet<Cell>(this.CellByCoordinates.Values);
            var changedCells = new List<ICell>();
            foreach (InteropRange interopCell in this.InteropWorksheet.UsedRange)
            {
                var coordinates = interopCell.Coordinates();
                if (!this.CellByCoordinates.TryGetValue(coordinates, out var cell))
                {
                    if (interopCell.Value2 == null)
                    {
                        // There was no cell, so no need to create one
                        continue;
                    }

                    cell = this[coordinates];
                }

                deletedCells.Remove(cell);

                if (cell.UpdateValue(interopCell.Value2))
                {
                    changedCells.Add(cell);
                }
            }

            changedCells.AddRange(deletedCells);
            foreach (var deletedCell in deletedCells)
            {
                deletedCell.Clear();
            }

            if (changedCells.Count > 0)
            {
                this.CellsChanged?.Invoke(this, new CellChangedEvent(changedCells.ToArray()));
            }
        }

        public void SetPageSetup(PageSetup pageSetup)
        {
            if (pageSetup == null)
            {
                return;
            }

            this.InteropWorksheet.PageSetup.Orientation = pageSetup.Orientation == 1
                    ? XlPageOrientation.xlPortrait
                    : XlPageOrientation.xlLandscape;

            if (pageSetup.PaperSize >= 1 && pageSetup.PaperSize <= 41)
            {
                this.InteropWorksheet.PageSetup.PaperSize = (XlPaperSize)pageSetup.PaperSize;
            }

            if (pageSetup.Header?.Margin >= 0)
            {
                this.InteropWorksheet.PageSetup.HeaderMargin = pageSetup.Header?.Margin ?? 0;
            }

            this.InteropWorksheet.PageSetup.LeftHeader = pageSetup.Header?.Left;
            this.InteropWorksheet.PageSetup.CenterHeader = pageSetup.Header?.Center;
            this.InteropWorksheet.PageSetup.RightHeader = pageSetup.Header?.Right;

            if (pageSetup.Footer?.Margin >= 0)
            {
                this.InteropWorksheet.PageSetup.FooterMargin = pageSetup.Footer?.Margin ?? 0;
            }

            this.InteropWorksheet.PageSetup.LeftFooter = pageSetup.Footer?.Left;
            this.InteropWorksheet.PageSetup.CenterFooter = pageSetup.Footer?.Center;
            this.InteropWorksheet.PageSetup.RightFooter = pageSetup.Footer?.Right;
        }

        public void SetChartObjectSourceData(object chartObject, object pivotTable)
        {
            var chart = ((ChartObject)this.InteropWorksheet.ChartObjects(chartObject)).Chart;
            var interopPivotTable = (InteropPivotTable)this.InteropWorksheet.PivotTables(pivotTable);

            chart.SetSourceData(interopPivotTable.TableRange1);
            chart.Refresh();
        }

        private void WaitAndRetry(Action action, int waitTime = 100, int maxRetries = 10)
        {
            Exception exception = null;

            for (var i = 0; i < maxRetries; i++)
            {
                try
                {
                    action();
                    exception = null;
                }
                catch (COMException e)
                {
                    exception = e;
                    Thread.Sleep(waitTime);
                }
            }


            if (exception != null)
            {
                throw exception;
            }
        }

        private T WaitAndRetry<T>(Func<T> func, int waitTime = 100, int maxRetries = 10)
        {
            T result = default;
            Exception exception = null;

            for (var i = 0; i < maxRetries; i++)
            {
                try
                {
                    result = func();
                    exception = null;
                }
                catch (COMException e)
                {
                    exception = e;
                    Thread.Sleep(waitTime);
                }
            }

            if (exception != null)
            {
                throw exception;
            }

            return result;
        }
    }
}
