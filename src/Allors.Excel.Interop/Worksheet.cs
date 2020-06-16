// <copyright file="Worksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using Allors.Excel;
    using Polly;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using InteropCustomProperty = Microsoft.Office.Interop.Excel.CustomProperty;
    using InteropDocEvents_Event = Microsoft.Office.Interop.Excel.DocEvents_Event;
    using InteropName = Microsoft.Office.Interop.Excel.Name;
    using InteropPivotTable = Microsoft.Office.Interop.Excel.PivotTable;
    using InteropPivotTables = Microsoft.Office.Interop.Excel.PivotTables;
    using InteropRange = Microsoft.Office.Interop.Excel.Range;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropXlCalculation = Microsoft.Office.Interop.Excel.XlCalculation;
    using InteropXlDeleteShiftDirection = Microsoft.Office.Interop.Excel.XlDeleteShiftDirection;
    using InteropXlDVAlertStyle = Microsoft.Office.Interop.Excel.XlDVAlertStyle;
    using InteropXlDVType = Microsoft.Office.Interop.Excel.XlDVType;
    using InteropXlFixedFormatQuality = Microsoft.Office.Interop.Excel.XlFixedFormatQuality;
    using InteropXlFixedFormatType = Microsoft.Office.Interop.Excel.XlFixedFormatType;
    using InteropXlInsertShiftDirection = Microsoft.Office.Interop.Excel.XlInsertShiftDirection;
    using InteropXlSheetVisibility = Microsoft.Office.Interop.Excel.XlSheetVisibility;

    public class Worksheet : IWorksheet
    {
        public bool isActive;

        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;

            RowByIndex = new Dictionary<int, Row>();
            ColumnByIndex = new Dictionary<int, Column>();
            CellByCoordinates = new Dictionary<(int, int), Cell>();
            DirtyValueCells = new HashSet<Cell>();
            DirtyCommentCells = new HashSet<Cell>();
            DirtyStyleCells = new HashSet<Cell>();
            DirtyOptionCells = new HashSet<Cell>();
            DirtyNumberFormatCells = new HashSet<Cell>();
            DirtyFormulaCells = new HashSet<Cell>();
            DirtyRows = new HashSet<Row>();

            interopWorksheet.Change += InteropWorksheet_Change;

            ((InteropDocEvents_Event)interopWorksheet).Activate += () =>
            {
                isActive = true;
                SheetActivated?.Invoke(this, Name);
            };

            interopWorksheet.Deactivate += () => isActive = false;

            this.Reset();
        }

        public event EventHandler<CellChangedEvent> CellsChanged;

        public event EventHandler<string> SheetActivated;

        private Range FreezeRange { get; set; }

        public int Index => InteropWorksheet.Index;

        public bool IsActive
        {
            get => isActive;
            set
            {
                if (value)
                {
                    isActive = true;

                    InteropWorksheet.Activate();
                }
                else
                {
                    isActive = false;
                }
            }
        }

        public Workbook Workbook { get; set; }

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => InteropWorksheet.Name; set => InteropWorksheet.Name = value; }

        IWorkbook Excel.IWorksheet.Workbook => Workbook;

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
            get => InteropWorksheet.Visible == InteropXlSheetVisibility.xlSheetVisible;
            set
            {
                if (value)
                {
                    InteropWorksheet.Visible = InteropXlSheetVisibility.xlSheetVisible;
                }
                else
                {
                    InteropWorksheet.Visible = InteropXlSheetVisibility.xlSheetHidden;
                }
            }
        }

        public async Task RefreshPivotTables(string sourceDataRange = null)
        {
            InteropPivotTables pivotTables = (InteropPivotTables)InteropWorksheet.PivotTables();

            foreach (InteropPivotTable pivotTable in pivotTables)
            {
                if (!string.IsNullOrWhiteSpace(sourceDataRange))
                {
                    pivotTable.SourceData = sourceDataRange;
                }

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
                if (!CellByCoordinates.TryGetValue(coordinates, out Cell cell))
                {
                    cell = new Cell(this, Row(coordinates.Item1), Column(coordinates.Item2));
                    CellByCoordinates.Add(coordinates, cell);
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

        IRow Excel.IWorksheet.Row(int index)
        {
            return Row(index);
        }

        IColumn Excel.IWorksheet.Column(int index)
        {
            return Column(index);
        }

        public Row Row(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException("Index can not be negative", nameof(this.Row));
            }

            if (!RowByIndex.TryGetValue(index, out Row row))
            {
                row = new Row(this, index);
                RowByIndex.Add(index, row);
            }

            return row;
        }

        public Column Column(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException(nameof(this.Column));
            }

            if (!ColumnByIndex.TryGetValue(index, out Column column))
            {
                column = new Column(this, index);
                ColumnByIndex.Add(index, column);
            }

            return column;
        }

        private Tuple<InteropXlCalculation, bool> DisableExcel()
        {
            InteropXlCalculation calculation = Workbook.InteropWorkbook.Application.Calculation;
            if (calculation != InteropXlCalculation.xlCalculationManual)
            {
                Workbook.InteropWorkbook.Application.Calculation = InteropXlCalculation.xlCalculationManual;
            }

            Workbook.InteropWorkbook.Application.ScreenUpdating = false;
            Workbook.InteropWorkbook.Application.EnableEvents = false;
            Workbook.InteropWorkbook.Application.DisplayStatusBar = false;
            Workbook.InteropWorkbook.Application.PrintCommunication = false;

            bool enableFormatConditionsCalculation = InteropWorksheet.EnableFormatConditionsCalculation;

            if (enableFormatConditionsCalculation)
            {
                InteropWorksheet.EnableFormatConditionsCalculation = false;
            }

            return Tuple.Create(calculation, enableFormatConditionsCalculation);
        }

        private void EnableExcel(Tuple<InteropXlCalculation, bool> tuple)
        {
            Workbook.InteropWorkbook.Application.Calculation = tuple.Item1;
            Workbook.InteropWorkbook.Application.ScreenUpdating = true;
            Workbook.InteropWorkbook.Application.EnableEvents = true;
            Workbook.InteropWorkbook.Application.DisplayStatusBar = true;
            Workbook.InteropWorkbook.Application.PrintCommunication = true;

            InteropWorksheet.EnableFormatConditionsCalculation = tuple.Item2;

            try
            {
                // Recalculate when required. Formulas need to be resolved.
                if (tuple.Item1 == InteropXlCalculation.xlCalculationAutomatic)
                {
                    InteropWorksheet.Calculate();
                }
            }
            catch
            {
            }
        }
        public async Task Flush()
        {
            Tuple<InteropXlCalculation, bool> tuple = DisableExcel();

            try
            {
                RenderNumberFormat(DirtyNumberFormatCells);
                DirtyNumberFormatCells = new HashSet<Cell>();

                RenderValue(DirtyValueCells);
                DirtyValueCells = new HashSet<Cell>();

                RenderFormula(DirtyFormulaCells);
                DirtyFormulaCells = new HashSet<Cell>();

                RenderComments(DirtyCommentCells);
                DirtyCommentCells = new HashSet<Cell>();

                RenderStyle(DirtyStyleCells);
                DirtyStyleCells = new HashSet<Cell>();

                SetOptions(DirtyOptionCells);
                DirtyOptionCells = new HashSet<Cell>();

                UpdateRows(DirtyRows);
                DirtyRows = new HashSet<Row>();
            }
            finally
            {
                EnableExcel(tuple);
            }

            await Task.CompletedTask;
        }

        public void AddDirtyNumberFormat(Cell cell)
        {
            DirtyNumberFormatCells.Add(cell);
        }

        public void AddDirtyValue(Cell cell)
        {
            DirtyValueCells.Add(cell);
        }

        public void AddDirtyFormula(Cell cell)
        {
            DirtyFormulaCells.Add(cell);
        }

        public void AddDirtyComment(Cell cell)
        {
            DirtyCommentCells.Add(cell);
        }

        public void AddDirtyStyle(Cell cell)
        {
            DirtyStyleCells.Add(cell);
        }

        public void AddDirtyOptions(Cell cell)
        {
            DirtyOptionCells.Add(cell);
        }

        public void AddDirtyRow(Row row)
        {
            DirtyRows.Add(row);
        }

        private void InteropWorksheet_Change(InteropRange target)
        {
            if (PreventChangeEvent)
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
                    CellsChanged?.Invoke(this, new CellChangedEvent(cells.Cast<ICell>().ToArray()));
                }
            }
        }

        private void RenderValue(IEnumerable<Cell> cells)
        {
            IEnumerable<IList<IList<Cell>>> chunks = cells.Chunks((v, w) => true);

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    object[,] values = new object[chunk.Count, chunk[0].Count];
                    for (int i = 0; i < chunk.Count; i++)
                    {
                        for (int j = 0; j < chunk[0].Count; j++)
                        {
                            values[i, j] = chunk[i][j].Value;
                        }
                    }

                    Row fromRow = chunk.First().First().Row;
                    Column fromColumn = chunk.First().First().Column;

                    Row toRow = chunk.Last().Last().Row;
                    Column toColumn = chunk.Last().Last().Column;

                    InteropRange range = WaitAndRetry(() =>
                    {
                        InteropRange from = (InteropRange)InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        InteropRange to = (InteropRange)InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
                    {
                        range.Value2 = values;
                    });
                });
        }

        private void RenderFormula(IEnumerable<Cell> cells)
        {
            IEnumerable<IList<IList<Cell>>> chunks = cells.Chunks((v, w) => true);

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    object[,] formulas = new object[chunk.Count, chunk[0].Count];
                    for (int i = 0; i < chunk.Count; i++)
                    {
                        for (int j = 0; j < chunk[0].Count; j++)
                        {
                            formulas[i, j] = chunk[i][j].Formula;
                        }
                    }

                    Row fromRow = chunk.First().First().Row;
                    Column fromColumn = chunk.First().First().Column;

                    Row toRow = chunk.Last().Last().Row;
                    Column toColumn = chunk.Last().Last().Column;

                    InteropRange range = WaitAndRetry(() =>
                    {
                        InteropRange from = (InteropRange)InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        InteropRange to = (InteropRange)InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
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
                    InteropRange range = WaitAndRetry(() =>
                    {
                        return (InteropRange)InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];
                    });

                    WaitAndRetry(() =>
                    {
                        if (range.Comment == null)
                        {
                            Microsoft.Office.Interop.Excel.Comment comment = range.AddComment(cell.Comment);
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
            IEnumerable<IList<IList<Cell>>> chunks = cells.Chunks((v, w) => Equals(v.Style, w.Style));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    Row fromRow = chunk.First().First().Row;
                    Column fromColumn = chunk.First().First().Column;

                    Row toRow = chunk.Last().Last().Row;
                    Column toColumn = chunk.Last().Last().Column;

                    InteropRange range = WaitAndRetry(() =>
                    {
                        object from = InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        object to = InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
                    {
                        Cell cc = chunk[0][0];
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
            IEnumerable<IList<IList<Cell>>> chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    Row fromRow = chunk.First().First().Row;
                    Column fromColumn = chunk.First().First().Column;

                    Row toRow = chunk.Last().Last().Row;
                    Column toColumn = chunk.Last().Last().Column;

                    InteropRange range = WaitAndRetry(() =>
                    {
                        object from = InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        object to = InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
                    {
                        range.NumberFormat = chunk[0][0].NumberFormat;
                    });
                });
        }

        private void SetOptions(IEnumerable<Cell> cells)
        {
            IEnumerable<IList<IList<Cell>>> chunks = cells.Chunks((v, w) => Equals(v.Options, w.Options));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    Row fromRow = chunk.First().First().Row;
                    Column fromColumn = chunk.First().First().Column;

                    Row toRow = chunk.Last().Last().Row;
                    Column toColumn = chunk.Last().Last().Column;

                    InteropRange range = WaitAndRetry(() =>
                    {
                        object from = InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                        object to = InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
                    {
                        Cell cc = chunk[0][0];
                        if (cc.Options != null)
                        {
                            string validationRange = cc.Options.Name;
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
            List<IList<Row>> chunks = dirtyRows.OrderBy(w => w.Index).Aggregate(
                        new List<IList<Row>> { new List<Row>() },
                        (acc, w) =>
                        {
                            IList<Row> list = acc[acc.Count - 1];
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

            IEnumerable<IList<Row>> updateChunks = chunks.Where(v => v.Count > 0);

            Parallel.ForEach(
                updateChunks,
                chunk =>
                {
                    Row fromChunk = chunk.First();
                    Row toChunk = chunk.Last();
                    bool hidden = fromChunk.Hidden;

                    string from = $"$A${fromChunk.Index + 1}";
                    string to = $"$A${toChunk.Index + 1}";

                    InteropRange range = WaitAndRetry(() =>
                    {
                        return InteropWorksheet.Range[from, to];
                    });

                    WaitAndRetry(() =>
                    {
                        range.EntireRow.Hidden = hidden;
                    });
                });
        }

        private void WaitAndRetry(Action method, int waitTime = 100, int maxRetries = 10)
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

        private T WaitAndRetry<T>(Func<T> method, int waitTime = 100, int maxRetries = 10)
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
            Workbook.AddIn.Office.AddPicture(InteropWorksheet, fileName, rectangle);

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
                name = InteropWorksheet.Names.Item(namedRange);
            }
            catch
            {
                throw new ArgumentException("Name not found for namedRange", nameof(namedRange));
            }


            // when the range is a mergedrange, then take the values from the area
            InteropRange range = name.RefersToRange;

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
                InteropRange area = range.MergeArea;

                int left = Convert.ToInt32(area.Left);
                int top = Convert.ToInt32(area.Top);
                int width = Convert.ToInt32(area.Width);
                int height = Convert.ToInt32(area.Height);

                return new System.Drawing.Rectangle(left, top, width, height);
            }

        }

        public Range[] GetNamedRanges()
        {
            List<Range> ranges = new List<Excel.Range>();

            foreach (Microsoft.Office.Interop.Excel.Name namedRange in InteropWorksheet.Names)
            {
                try
                {
                    InteropRange refersToRange = namedRange.RefersToRange;
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
        public void SetNamedRange(string name, Range range)
        {
            if (!string.IsNullOrWhiteSpace(name) && range != null)
            {
                try
                {
                    object topLeft = InteropWorksheet.Cells[range.Row + 1, range.Column + 1];
                    object bottomRight = InteropWorksheet.Cells[range.Row + range.Rows, range.Column + range.Columns];

                    InteropRange refersTo = InteropWorksheet.Range[topLeft, bottomRight];

                    // When it does not exist, add it, else we update the range.
                    if (InteropWorksheet.Names
                        .Cast<Microsoft.Office.Interop.Excel.Name>()
                        .Any(v => string.Equals(v.Name, name)))
                    {
                        InteropWorksheet.Names.Item(name).RefersTo = refersTo;
                    }
                    else
                    {
                        InteropWorksheet.Names.Add(name, refersTo);
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
                PreventChangeEvent = true;

                try
                {
                    InteropRange rows = InteropWorksheet.Range[$"{startRowIndex + 2}:{startRowIndex + numberOfRows + 1}"];

                    WaitAndRetry(() =>
                    {
                        Tuple<InteropXlCalculation, bool> tuple = DisableExcel();
                        rows.Insert(InteropXlInsertShiftDirection.xlShiftDown);
                        EnableExcel(tuple);
                    });

                    if (CellByCoordinates.Any())
                    {
                        // Shift all cell rows down with the numberOfRows
                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                        .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                        .OrderByDescending(kvp => kvp.Value.Row.Index)
                        .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            Cell cell = item.Value;
                            CellByCoordinates.Remove(item.Key);

                            // Shift rows up with the numberofrows that were deleted.
                            cell.Row.Index += numberOfRows;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    PreventChangeEvent = false;
                }
            }
        }

        public void DeleteRows(int startRowIndex, int numberOfRows)
        {
            if (startRowIndex >= 0 && numberOfRows > 0)
            {
                PreventChangeEvent = true;

                try
                {
                    InteropRange rows = InteropWorksheet.Range[$"{startRowIndex + 1}:{startRowIndex + numberOfRows}"];

                    WaitAndRetry(() =>
                    {
                        Tuple<InteropXlCalculation, bool> tuple = DisableExcel();
                        rows.Delete(InteropXlDeleteShiftDirection.xlShiftUp);
                        EnableExcel(tuple);
                    });

                    if (CellByCoordinates.Any())
                    {
                        // Delete all cells in the deleted rows.
                        foreach (int rowIndex in Enumerable.Range(startRowIndex, numberOfRows))
                        {
                            foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                              .Where(kvp => kvp.Value.Row.Index == rowIndex)
                              .ToList())
                            {
                                CellByCoordinates.Remove(item.Key);
                            }
                        }

                        // Shift all cell rows up with the numberOfRows
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                            .Where(kvp => kvp.Value.Row.Index > startRowIndex)
                            .OrderBy(kvp => kvp.Value.Row.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            Cell cell = item.Value;
                            CellByCoordinates.Remove(item.Key);

                            int rowIndex = cell.Row.Index - numberOfRows;

                            // Link the cell to the new Row that already exists
                            cell.Row = RowByIndex[rowIndex];

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    PreventChangeEvent = false;
                }
            }
        }

        public void InsertColumns(int startColumnIndex, int numberOfColumns)
        {
            if (startColumnIndex >= 0 && numberOfColumns > 0)
            {
                PreventChangeEvent = true;

                try
                {
                    string startColumnName = ExcelColumnFromNumber(startColumnIndex + 2);
                    string endColumnName = ExcelColumnFromNumber(startColumnIndex + 1 + numberOfColumns);

                    InteropRange rows = InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    WaitAndRetry(() =>
                    {
                        Tuple<InteropXlCalculation, bool> tuple = DisableExcel();
                        rows.EntireColumn.Insert(InteropXlInsertShiftDirection.xlShiftToRight);
                        EnableExcel(tuple);
                    });

                    if (CellByCoordinates.Any())
                    {
                        // Shift all cell columns to the right with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        // Order by descending so we will not have a duplicate key in the dictionary
                        foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                                    .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                                    .OrderByDescending(kvp => kvp.Value.Column.Index)
                                    .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            Cell cell = item.Value;
                            CellByCoordinates.Remove(item.Key);

                            // Shift rows up with the numberofrows that were deleted.
                            cell.Column.Index += numberOfColumns;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            CellByCoordinates[coordinates] =  cell;
                        }
                    }
                }
                finally
                {
                    PreventChangeEvent = false;
                }
            }
        }

        public void DeleteColumns(int startColumnIndex, int numberOfColumns)
        {
            if (startColumnIndex >= 0 && numberOfColumns > 0)
            {
                PreventChangeEvent = true;

                try
                {
                    string startColumnName = ExcelColumnFromNumber(startColumnIndex + 1);
                    string endColumnName = ExcelColumnFromNumber(startColumnIndex + numberOfColumns);

                    InteropRange range = InteropWorksheet.Range[$"{startColumnName}1:{endColumnName}1"];

                    WaitAndRetry(() =>
                    {
                        Tuple<InteropXlCalculation, bool> tuple = DisableExcel();
                        range.EntireColumn.Delete(InteropXlDeleteShiftDirection.xlShiftToLeft);
                        EnableExcel(tuple);
                    });

                    if (CellByCoordinates.Any())
                    {
                        // Delete all cells in the deleted columns.
                        foreach (int columnIndex in Enumerable.Range(startColumnIndex, numberOfColumns))
                        {
                            foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                                .Where(kvp => kvp.Value.Column.Index == columnIndex)
                                .ToList())
                            {
                                CellByCoordinates.Remove(item.Key);
                            }
                        }

                        // Shift all cell Columns to the left with the numberOfColumns
                        // We need to set the Key in the correct format "row:column" in order to track the rights cells.
                        foreach (KeyValuePair<(int, int), Cell> item in CellByCoordinates
                            .Where(kvp => kvp.Value.Column.Index > startColumnIndex)
                            .OrderBy(kvp => kvp.Value.Column.Index)
                            .ToList())
                        {
                            // Remove the entry in the dictionary, but keep the cell.
                            Cell cell = item.Value;
                            CellByCoordinates.Remove(item.Key);

                            int columnIndex = cell.Column.Index - numberOfColumns;

                            Column column = Column(columnIndex);

                            // Link to the correct column that already exists.
                            cell.Column = column;

                            // Add the existing cell with its new key
                            var coordinates = (cell.Row.Index, cell.Column.Index);
                            CellByCoordinates.Add(coordinates, cell);
                        }
                    }
                }
                finally
                {
                    PreventChangeEvent = false;
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
                    interopRange = InteropWorksheet.Range[cell1];
                }
                else
                {
                    interopRange = InteropWorksheet.Range[cell1, cell2];
                }


                return new Excel.Range(interopRange.Row - 1, interopRange.Column - 1, interopRange.Rows.Count, interopRange.Columns.Count, this);
            }
            catch
            {
                return null;
            }


        }

        public Range GetUsedRange()
        {
            InteropRange range = InteropWorksheet.UsedRange;

            return new Excel.Range(range.Row - 1, range.Column - 1, range.Rows.Count, range.Columns.Count, this);
        }

        public Range GetUsedRange(int row)
        {
            if (row < 0 || row >= InteropWorksheet.UsedRange.Row + InteropWorksheet.UsedRange.Rows.Count)
            {
                return null;
            }

            InteropRange rowRange = (InteropRange)InteropWorksheet.Rows[row + 1];

            int endColumnIndex = InteropWorksheet.UsedRange.Column + InteropWorksheet.UsedRange.Columns.Count - 1;
            bool quit = false;

            do
            {
                InteropRange cell = (InteropRange)InteropWorksheet.Cells[rowRange.Row, endColumnIndex];

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

            int beginColumnIndex = rowRange.Column;
            quit = false;

            do
            {
                InteropRange cell = (InteropRange)InteropWorksheet.Cells[rowRange.Row, beginColumnIndex];

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
            int columnCount = 1 + endColumnIndex - beginColumnIndex;

            return new Excel.Range(rowRange.Row - 1, beginColumnIndex - 1, rowRange.Rows.Count, columnCount, this);
        }

        public Range GetUsedRange(string column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return null;
            }

            int columnIndex = ExcelColumnFromLetter(column);
            InteropRange columnRange = (InteropRange)InteropWorksheet.Columns[columnIndex];

            int beginRowIndex = columnRange.Row;
            int maxRows = InteropWorksheet.UsedRange.Rows.Count;
            bool quit = false;

            do
            {
                InteropRange cell = (InteropRange)InteropWorksheet.Cells[beginRowIndex, columnRange.Column];

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


            int endRowIndex = InteropWorksheet.UsedRange.Row + InteropWorksheet.UsedRange.Rows.Count - 1;
            quit = false;

            do
            {
                InteropRange cell = (InteropRange)InteropWorksheet.Cells[endRowIndex, columnRange.Column];

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

            int rowCount = 1 + endRowIndex - beginRowIndex;

            return new Excel.Range(beginRowIndex - 1, columnRange.Column - 1, rowCount, columnRange.Columns.Count, this);
        }

        /// <inheritdoc/>
        /// <summary>
        /// When range.Row = 0 and range.Column = -1, then topRow in frozen
        /// When range.Row = -1 and range.Column = 0 then leftColumn in  frozen
        /// When range.Row > 0 and range.Column > 0, then that cell is the topleft position for the freezepanes
        /// </summary>
        public void FreezePanes(Range range)
        {
            InteropWorksheet.Application.ScreenUpdating = true;
            InteropWorksheet.Activate();

            InteropWorksheet.Application.ActiveWindow.FreezePanes = false;

            if (range.Row > 0 && range.Column > 0)
            {
                InteropWorksheet.Application.ActiveWindow.SplitRow = range.Row;
                InteropWorksheet.Application.ActiveWindow.SplitColumn = range.Column;
            }
            else
            {
                int row = 0;
                if (range.Row > -1)
                {
                    row = range.Row + 1;
                }

                InteropWorksheet.Application.ActiveWindow.SplitRow = row;

                int column = 0;
                if (range.Column > -1)
                {
                    column = range.Column + 1;
                }
                InteropWorksheet.Application.ActiveWindow.SplitColumn = column;
            }

            InteropWorksheet.Application.ActiveWindow.FreezePanes = true;

            FreezeRange = range;
        }

        public void UnfreezePanes()
        {
            InteropWorksheet.Application.ScreenUpdating = true;
            InteropWorksheet.Activate();

            InteropWorksheet.Application.ActiveWindow.SplitRow = 0;
            InteropWorksheet.Application.ActiveWindow.SplitColumn = 0;
            InteropWorksheet.Application.ActiveWindow.FreezePanes = false;

            FreezeRange = null;
        }

        public bool HasFreezePanes => FreezeRange != null;
             
        public void SaveAsXPS(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            SaveAs(file, InteropXlFixedFormatType.xlTypeXPS, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
        }

        public void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            SaveAs(file, InteropXlFixedFormatType.xlTypePDF, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
        }

        /// <summary>
        /// Save the sheet in the given formattype (0=PDF, 1=XPS) 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="formatType"></param>
        /// <param name="overwriteExistingFile"></param>
        /// <param name="openAfterPublish"></param>
        /// <param name="ignorePrintAreas"></param>
        private void SaveAs(FileInfo file, InteropXlFixedFormatType formatType, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            FileInfo fi = new FileInfo(file.FullName);

            // In case we would overwrite an existing file
            if (fi.Exists && !overwriteExistingFile)
            {
                throw new IOException($"File {file.FullName} already exists and should not be overwritten.");
            }

            if (formatType == InteropXlFixedFormatType.xlTypePDF && !string.Equals(fi.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                fi = new FileInfo(Path.ChangeExtension(fi.FullName, ".pdf"));
            }

            if (formatType == InteropXlFixedFormatType.xlTypeXPS && !string.Equals(fi.Extension, ".xps", StringComparison.OrdinalIgnoreCase))
            {
                fi = new FileInfo(Path.ChangeExtension(fi.FullName, ".xps"));
            }

            if (!Directory.Exists(fi.DirectoryName))
            {
                Directory.CreateDirectory(fi.DirectoryName);
            }

            InteropWorksheet
                .ExportAsFixedFormat
                (
                       Type: formatType,
                       Filename: fi.FullName,
                       Quality: InteropXlFixedFormatQuality.xlQualityStandard,
                       IgnorePrintAreas: ignorePrintAreas,
                       OpenAfterPublish: openAfterPublish
               );
        }

        /// <inheritdoc/>
        public void SetPrintArea(Range range = null)
        {
            // Use A1-Style reference for the printarea

            string printArea = "";

            if (range != null)
            {
                // row 3, column 5, rows 6 column 2
                // => A1 Style = F4:H10
                string startColumn = ExcelColumnFromNumber(range.Column + 1); // !zero-based
                int startRow = range.Row + 1;

                string endColumn = ExcelColumnFromNumber(range.Column + range.Columns.GetValueOrDefault());
                int endRow = range.Row + range.Rows.GetValueOrDefault();

                printArea = $"{startColumn}{startRow}:{endColumn}{endRow}";
            }

            InteropWorksheet.PageSetup.PrintArea = printArea;
        }

        public void SetCustomProperties(CustomProperties properties)
        {
            InteropCustomProperty[] cps = InteropWorksheet.CustomProperties.Cast<InteropCustomProperty>().ToArray(); ;

            foreach (KeyValuePair<string, object> kvp in properties)
            {
                bool found = false;
                foreach (InteropCustomProperty cp in cps)
                {
                    if (string.Equals(cp.Name, kvp.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;

                        if (kvp.Value == null)
                        {
                            cp.Value = Excel.CustomProperties.MagicNull;
                        }
                        else
                        {
                            cp.Value = kvp.Value;
                        }
                    }
                }


                if (!found)
                {
                    if (kvp.Value == null)
                    {
                        InteropWorksheet.CustomProperties.Add(kvp.Key, Excel.CustomProperties.MagicNull);
                    }
                    else
                    {
                        InteropWorksheet.CustomProperties.Add(kvp.Key, kvp.Value);
                    }
                }

            }
        }

        public CustomProperties GetCustomProperties()
        {
            CustomProperties dict = new Excel.CustomProperties();

            foreach (InteropCustomProperty customProperty in InteropWorksheet.CustomProperties)
            {
                if (Excel.CustomProperties.MagicNull.Equals(customProperty.Value))
                {
                    dict.Add(customProperty.Name, null);
                }
                else
                {
                    dict.Add(customProperty.Name, customProperty.Value);
                }
            }

            return dict;
        }

        public void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true)
        {
            InteropRange inputCell = (InteropRange)InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];

            inputCell.Validation.Delete();
            inputCell.Validation.Add(InteropXlDVType.xlValidateInputOnly);
            inputCell.Validation.ShowInput = showInputMessage;

            inputCell.Validation.InputMessage = message;
            inputCell.Validation.InputTitle = title;
        }

        public void HideInputMessage(ICell cell, bool clearInputMessage = false)
        {
            InteropRange inputCell = (InteropRange)InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];

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
            List<ICell> changedCells = new List<ICell>();
            foreach (InteropRange interopCell in (InteropRange)this.InteropWorksheet.UsedRange)
            {
                (int, int) coordinates = interopCell.Coordinates();
                if (!CellByCoordinates.TryGetValue(coordinates, out var cell))
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
                CellsChanged?.Invoke(this, new CellChangedEvent(changedCells.ToArray()));
            }
        }
    }
}
