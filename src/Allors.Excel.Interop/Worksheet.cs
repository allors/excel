// <copyright file="Worksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Embedded
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Threading.Tasks;
    using Allors.Excel;
    using Microsoft.Office.Interop.Excel;
    using Polly;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public interface IEmbeddedWorksheet : IWorksheet
    {
        void AddDirtyValue(Cell cell);

        void AddDirtyFormula(Cell cell);

        void AddDirtyComment(Cell cell);

        void AddDirtyStyle(Cell cell);

        void AddDirtyNumberFormat(Cell cell);

        void AddDirtyOptions(Cell cell);

        void AddDirtyRow(Row row);
    }

    public class Worksheet : IEmbeddedWorksheet
    {
        private Dictionary<int, Row> rowByIndex;

        private Dictionary<int, Column> columnByIndex;
        private bool isActive;
        private Excel.Range FreezeRange { get; set; }

        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;
            this.rowByIndex = new Dictionary<int, Row>();
            this.columnByIndex = new Dictionary<int, Column>();
            this.CellByRowColumn = new Dictionary<string, Cell>();
            this.DirtyValueCells = new HashSet<Cell>();
            this.DirtyCommentCells = new HashSet<Cell>();
            this.DirtyStyleCells = new HashSet<Cell>();
            this.DirtyOptionCells = new HashSet<Cell>();
            this.DirtyNumberFormatCells = new HashSet<Cell>();
            this.DirtyFormulaCells = new HashSet<Cell>();
            this.DirtyRows = new HashSet<Row>();

            interopWorksheet.Change += this.InteropWorksheet_Change;

            ((DocEvents_Event)interopWorksheet).Activate += () =>
            {
                this.isActive = true;
                this.SheetActivated?.Invoke(this, this.Name);
            };

            ((DocEvents_Event)interopWorksheet).Deactivate += () => this.isActive = false;
        }

        public event EventHandler<CellChangedEvent> CellsChanged;

        public event EventHandler<string> SheetActivated;

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

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.Name; set => this.InteropWorksheet.Name = value; }

        IWorkbook IWorksheet.Workbook => this.Workbook;

        private Dictionary<string, Cell> CellByRowColumn { get; }

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
            get => this.InteropWorksheet.Visible == XlSheetVisibility.xlSheetVisible;
            set
            {
                if (value)
                {
                    this.InteropWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
                else
                {
                    this.InteropWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
                }
            }
        }

        public async Task RefreshPivotTables(string sourceDataRange = null)
        {
            var pivotTables = (PivotTables)this.InteropWorksheet.PivotTables();

            foreach (PivotTable pivotTable in pivotTables)
            {
                if (!string.IsNullOrWhiteSpace(sourceDataRange))
                {
                    pivotTable.SourceData = sourceDataRange;
                }

                pivotTable.RefreshTable();
            }

            await Task.CompletedTask;
        }

        public ICell this[int row, int column]
        {
            get
            {
                var key = $"{row}:{column}";
                if (!this.CellByRowColumn.TryGetValue(key, out var cell))
                {
                    cell = new Cell(this, this.Row(row), this.Column(column));
                    this.CellByRowColumn.Add(key, cell);
                }

                return cell;
            }
        }

        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = string.Empty;
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        public static int ExcelColumnFromLetter(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        IRow IWorksheet.Row(int index) => this.Row(index);

        IColumn IWorksheet.Column(int index) => this.Column(index);

        public Row Row(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException("Index can not be negative", nameof(this.Row));
            }

            if (!this.rowByIndex.TryGetValue(index, out var row))
            {
                row = new Row(this, index);
                this.rowByIndex.Add(index, row);
            }

            return row;
        }

        public Column Column(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException(nameof(this.Column));
            }

            if (!this.columnByIndex.TryGetValue(index, out var column))
            {
                column = new Column(this, index);
                this.columnByIndex.Add(index, column);
            }

            return column;
        }

        private Tuple<XlCalculation, bool> DisableExcel()
        {
            var calculation = this.Workbook.InteropWorkbook.Application.Calculation;
            if (calculation != XlCalculation.xlCalculationManual)
            {
                this.Workbook.InteropWorkbook.Application.Calculation = XlCalculation.xlCalculationManual;
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

            return Tuple.Create(calculation,  enableFormatConditionsCalculation);
        }

        private void EnableExcel(Tuple<XlCalculation, bool> tuple)
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
                if (tuple.Item1 == XlCalculation.xlCalculationAutomatic)
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

        private void InteropWorksheet_Change(Microsoft.Office.Interop.Excel.Range target)
        {
            if (this.PreventChangeEvent)
            {
                return;
            }

            List<Cell> cells = null;
            foreach (Microsoft.Office.Interop.Excel.Range targetCell in target.Cells)
            {
                var row = targetCell.Row - 1;
                var column = targetCell.Column - 1;
                var cell = (Cell)this[row, column];

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
                        var from = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
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
                        var from = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        var to = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
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
                        return (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];
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
                            range.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;
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

                            range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, Type.Missing, $"={validationRange}", Type.Missing);
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

                    string from = $"$A${fromChunk.Index + 1}";
                    string to = $"$A${toChunk.Index + 1}";

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

        private void WaitAndRetry(System.Action method, int waitTime = 100, int maxRetries = 10)
        {
            Policy
            .Handle<System.Runtime.InteropServices.COMException>()
            .WaitAndRetry(
                maxRetries,
                (retryCount) =>
                {
                    // returns the waitTime for the onRetry
                    return TimeSpan.FromMilliseconds(waitTime * retryCount);
                })
                .Execute(method);
        }

        private T WaitAndRetry<T>(System.Func<T> method, int waitTime = 100, int maxRetries = 10)
        {
            return Policy
             .Handle<System.Runtime.InteropServices.COMException>()
             .WaitAndRetry(
                 maxRetries,
                 (retryCount) =>
                 {
                     // returns the waitTime for the onRetry
                     return TimeSpan.FromMilliseconds(waitTime * retryCount);
                 })
             .Execute(method);
        }

        /// <summary>
        /// Adds a Picture on the specified rectangle. <seealso cref="GetRectangle(string)"/>
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="location"></param>
        /// <param name="size"></param>
        public void AddPicture(string fileName, System.Drawing.Rectangle rectangle)
        {
            this.Workbook.AddIn.Office.AddPicture(this.InteropWorksheet, fileName, rectangle);

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
        public System.Drawing.Rectangle GetRectangle(string namedRange)
        {
            Name name = null;

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
                int left = Convert.ToInt32(range.Left);
                int top = Convert.ToInt32(range.Top);
                int width = Convert.ToInt32(range.Width);
                int height = Convert.ToInt32(range.Height);

                return new System.Drawing.Rectangle(left, top, width, height);              
            }
            else
            {
                var area = range.MergeArea;

                int left = Convert.ToInt32(area.Left);
                int top = Convert.ToInt32(area.Top);
                int width = Convert.ToInt32(area.Width);
                int height = Convert.ToInt32(area.Height);

                return new System.Drawing.Rectangle(left, top, width, height);
            }          

        }

        public Excel.Range[] GetNamedRanges()
        {
            var ranges = new List<Excel.Range>();

            foreach (Microsoft.Office.Interop.Excel.Name namedRange in this.InteropWorksheet.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {
                        ranges.Add(new Excel.Range(refersToRange.Row - 1, refersToRange.Column - 1, refersToRange.Rows.Count, refersToRange.Columns.Count, worksheet: this, name: namedRange.Name));
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
        public void SetNamedRange(string name, Excel.Range range)
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
                        .Cast<Microsoft.Office.Interop.Excel.Name>()
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
                    Microsoft.Office.Interop.Excel.Range rows = this.InteropWorksheet.Range[$"{startRowIndex + 2}:{startRowIndex + numberOfRows + 1}"];

                    this.WaitAndRetry(() => {
                        var tuple = this.DisableExcel();
                        rows.Insert(XlInsertShiftDirection.xlShiftDown);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByRowColumn.Any()) 
                    {
                        // Shift all cell rows down with the numberOfRows
                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (var item in this.CellByRowColumn
                        .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                        .OrderByDescending(kvp => kvp.Value.Row.Index)
                        .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByRowColumn.Remove(item.Key);

                            // Shift rows up with the numberofrows that were deleted.
                            cell.Row.Index += numberOfRows;

                            // Add the existing cell with its new key
                            var key = $"{cell.Row.Index}:{cell.Column.Index}";
                            this.CellByRowColumn.Add(key, cell);
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
                    Microsoft.Office.Interop.Excel.Range rows = this.InteropWorksheet.Range[$"{startRowIndex + 1}:{startRowIndex + numberOfRows}"];

                    this.WaitAndRetry(() => 
                    {
                        var tuple = this.DisableExcel();
                        rows.Delete(XlDeleteShiftDirection.xlShiftUp);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByRowColumn.Any()) 
                    {
                        // Delete all cells in the deleted rows.
                        foreach (int rowIndex in Enumerable.Range(startRowIndex, numberOfRows))
                        {
                            foreach (var item in this.CellByRowColumn
                              .Where(kvp => kvp.Value.Row.Index == rowIndex)
                              .ToList())
                            {
                                this.CellByRowColumn.Remove(item.Key);
                            }
                        }

                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (var item in this.CellByRowColumn
                            .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                            .OrderBy(kvp => kvp.Value.Row.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByRowColumn.Remove(item.Key);

                            var rowIndex = cell.Row.Index - numberOfRows;

                            // Link the cell to the new Row that already exists
                            cell.Row = this.rowByIndex[rowIndex];

                            // Add the existing cell with its new key
                            var key = $"{cell.Row.Index}:{cell.Column.Index}";
                            this.CellByRowColumn.Add(key, cell);
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

                    Microsoft.Office.Interop.Excel.Range rows = this.InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    this.WaitAndRetry(() => {
                        var tuple = this.DisableExcel();
                        rows.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByRowColumn.Any())
                    {
                        // Shift all cell columns to the right with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (var item in this.CellByRowColumn
                                    .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                                    .OrderByDescending(kvp => kvp.Value.Column.Index)
                                    .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByRowColumn.Remove(item.Key);

                            // Shift rows up with the numberofrows that were deleted.
                            cell.Column.Index += numberOfColumns;

                            // Add the existing cell with its new key
                            var key = $"{cell.Row.Index}:{cell.Column.Index}";
                            this.CellByRowColumn.Add(key, cell);
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

                    Microsoft.Office.Interop.Excel.Range range = this.InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    this.WaitAndRetry(() => {
                        var tuple = this.DisableExcel();
                        range.EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                        this.EnableExcel(tuple);
                    });

                    if (this.CellByRowColumn.Any())
                    {
                        // Delete all cells in the deleted columns.
                        foreach (int columnIndex in Enumerable.Range(startColumnIndex, numberOfColumns))
                        {
                            foreach (var item in this.CellByRowColumn
                                .Where(kvp => kvp.Value.Column.Index == columnIndex)
                                .ToList())
                            {
                                this.CellByRowColumn.Remove(item.Key);
                            }
                        }

                        // Shift all cell Columns to the left with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (var item in this.CellByRowColumn
                            .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                            .OrderBy(kvp => kvp.Value.Column.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            var cell = item.Value;
                            this.CellByRowColumn.Remove(item.Key);

                            var columnIndex = cell.Column.Index - numberOfColumns;

                            var column = this.columnByIndex[columnIndex];

                            // Link to the correct column that already exists.
                            cell.Column = column;

                            // Add the existing cell with its new key
                            var key = $"{cell.Row.Index}:{cell.Column.Index}";
                            this.CellByRowColumn.Add(key, cell);
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
        public Excel.Range GetRange(string cell1, string cell2 = null)
        {
            if (string.IsNullOrWhiteSpace(cell1) && cell2 == null)
            {
                return null;
            }

            try
            {
                Microsoft.Office.Interop.Excel.Range interopRange;
                if (cell2 == null)
                {
                    interopRange = this.InteropWorksheet.Range[cell1];
                }
                else
                {
                    interopRange = this.InteropWorksheet.Range[cell1, cell2];
                }


                return new Excel.Range(interopRange.Row - 1, interopRange.Column - 1, interopRange.Rows.Count, interopRange.Columns.Count, this);
            }
            catch
            {
                return null;
            }

                    
        }

        public Excel.Range GetUsedRange()
        {
            var range = this.InteropWorksheet.UsedRange;

            return new Excel.Range(range.Row - 1, range.Column - 1, range.Rows.Count, range.Columns.Count, this);
        }

        public Excel.Range GetUsedRange(int row)
        {
            if (row < 0 || row >= this.InteropWorksheet.UsedRange.Row + this.InteropWorksheet.UsedRange.Rows.Count)
            {
                return null;
            }

            var rowRange = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Rows[row+1];

            var endColumnIndex = this.InteropWorksheet.UsedRange.Column + this.InteropWorksheet.UsedRange.Columns.Count - 1;
            var quit = false;

            do
            {
                var cell = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[rowRange.Row, endColumnIndex];

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
                var cell = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[rowRange.Row, beginColumnIndex];

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

            return new Excel.Range(rowRange.Row - 1, beginColumnIndex - 1, rowRange.Rows.Count, columnCount, this);
        }

        public Excel.Range GetUsedRange(string column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return null;
            }

            var columnIndex = ExcelColumnFromLetter(column);
            var columnRange = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Columns[columnIndex];

            var beginRowIndex = columnRange.Row;
            var maxRows = this.InteropWorksheet.UsedRange.Rows.Count;
            var quit = false;

            do
            {
                var cell = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[beginRowIndex, columnRange.Column];

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
                var cell = (Microsoft.Office.Interop.Excel.Range)this.InteropWorksheet.Cells[endRowIndex, columnRange.Column];

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

            return new Excel.Range(beginRowIndex - 1, columnRange.Column - 1, rowCount, columnRange.Columns.Count, this);         
        }

        /// <inheritdoc/>
        /// <summary>
        /// When range.Row = 0 and range.Column = -1, then topRow in frozen
        /// When range.Row = -1 and range.Column = 0 then leftColumn in  frozen
        /// When range.Row > 0 and range.Column > 0, then that cell is the topleft position for the freezepanes
        /// </summary>
        public void FreezePanes(Excel.Range range)
        {
            this.InteropWorksheet.Application.ScreenUpdating = true;
            this.InteropWorksheet.Activate();

            this.InteropWorksheet.Application.ActiveWindow.FreezePanes = false;

            if(range.Row > 0 && range.Column > 0)
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
            this.SaveAs(file, XlFixedFormatType.xlTypeXPS, overwriteExistingFile, openAfterPublish, ignorePrintAreas);                  
        }

        public void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            this.SaveAs(file, XlFixedFormatType.xlTypePDF, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
        }

        /// <summary>
        /// Save the sheet in the given formattype (0=PDF, 1=XPS) 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="formatType"></param>
        /// <param name="overwriteExistingFile"></param>
        /// <param name="openAfterPublish"></param>
        /// <param name="ignorePrintAreas"></param>
        private void SaveAs(FileInfo file, XlFixedFormatType formatType, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            var fi = new FileInfo(file.FullName);

            // In case we would overwrite an existing file
            if (fi.Exists && !overwriteExistingFile)
            {
                throw new IOException($"File {file.FullName} already exists and should not be overwritten.");
            }

            if (!Directory.Exists(fi.DirectoryName))
            {
                Directory.CreateDirectory(fi.DirectoryName);
            }

            if (formatType == XlFixedFormatType.xlTypePDF && !string.Equals(fi.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
               fi = new FileInfo(Path.ChangeExtension(fi.FullName, ".pdf"));
            }

            if (formatType == XlFixedFormatType.xlTypeXPS && !string.Equals(fi.Extension, ".xps", StringComparison.OrdinalIgnoreCase))
            {
                fi = new FileInfo(Path.ChangeExtension(fi.FullName, ".xps"));
            }

            this.InteropWorksheet
                .ExportAsFixedFormat
                (
                       Type: formatType,
                       Filename: fi.FullName,
                       Quality: XlFixedFormatQuality.xlQualityStandard,
                       IgnorePrintAreas: ignorePrintAreas,
                       OpenAfterPublish: openAfterPublish
               );
        }


        /// <inheritdoc/>
        public void SetPrintArea(Excel.Range range = null)
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
    }
}
