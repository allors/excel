// <copyright file="Worksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading.Tasks;
    using QuestPDF;
    using QuestPDF.Fluent;
    using QuestPDF.Helpers;
    using QuestPDF.Infrastructure;

    public class Worksheet(Workbook workbook) : IWorksheet
    {
        private readonly Dictionary<int, Row> rowByIndex = new();

        private readonly Dictionary<int, Column> columnByIndex = new();
        private bool isActive;

        IWorkbook IWorksheet.Workbook => this.Workbook;

        public ICustomProperties CustomProperties { get; }

        public Workbook Workbook { get; } = workbook;

        public event EventHandler<CellChangedEvent> CellsChanged;
        public event EventHandler<CellChangedEvent> SheetChanged;
        public event EventHandler<string> SheetActivated;
        public event EventHandler<Allors.Excel.Hyperlink> HyperlinkClicked;

        public string Name { get; set; }

        public int Index
        {
            get
            {
                return this.Workbook.WorksheetList.IndexOf(this) + 1;
            }
        }
        //public int Index => throw new NotImplementedException();

        public Dictionary<(int, int), Cell> CellByCoordinates { get; } = new();

        public bool IsVisible { get; set; } = true;

        public bool HasFreezePanes
        {
            get
            {
                return this.FrozenRange != null;
            }
        }


        public bool IsActive
        {
            get
            {
                return this.isActive;
            }
            set
            {
                if (value == false)
                {
                    return;
                }

                foreach (var worksheet in this.Workbook.WorksheetList)
                {
                    worksheet.isActive = false;
                }
                this.isActive = value;
            }
        }

        ICell IWorksheet.this[(int, int) coordinates] => this[coordinates];

        ICell IWorksheet.this[int row, int column] => this[(row, column)];

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

        IRow IWorksheet.Row(int index) => this.Row(index);

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

        IColumn IWorksheet.Column(int index) => this.Column(index);

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

        public void Flush()
        {
        }

        public void Activate()
        {
            foreach (var worksheet in this.Workbook.WorksheetList)
            {
                worksheet.IsActive = false;
            }

            this.IsActive = true;
        }

        public void RefreshPivotTables()
        {
        }

        public void AddPicture(string uri, Rectangle rectangle)
        {
        }
        public Dictionary<string, Range> NamedRangeByName { get; set; } = new Dictionary<string, Range>();

        public Range[] GetNamedRanges()
        {
            return this.NamedRangeByName.Values.ToArray();
        }

        public Rectangle GetRectangle(string namedRange)
        {
            if (string.IsNullOrEmpty(namedRange))
            {
                throw new ArgumentException("Named range cannot be null or empty.", nameof(namedRange));
            }

            if (!this.NamedRangeByName.TryGetValue(namedRange, out var range))
            {
                throw new ArgumentException($"No range found with the name {namedRange}.", nameof(namedRange));
            }

            // Assuming that the Range object has properties for the start and end rows and columns
            int startRow = range.Row;
            int startColumn = range.Column;
            int endRow = (int)(range.Row + range.Rows - 1);
            int endColumn = (int)(range.Column + range.Columns - 1);

            // Calculate the width and height of the rectangle
            int width = endColumn - startColumn + 1;
            int height = endRow - startRow + 1;

            // Create a new Rectangle object with the calculated properties
            return new Rectangle(startColumn, startRow, width, height);
        }





        public void SetNamedRange(string name, Range range)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("Name cannot be null or empty.", nameof(name));
            }

            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            // Replace the existing named range with the new range
            this.NamedRangeByName[name] = range;
            range.Name = name;
        }

        public void InsertRows(int startRowIndex, int numberOfRows) => throw new NotImplementedException();

        public void DeleteRows(int startRowIndex, int numberOfRows)
        {
            if (startRowIndex < 0 || numberOfRows < 0)
            {
                throw new ArgumentException("Start index and number of rows must be non-negative.");
            }

            // Remove the rows from the dictionary
            for (int i = startRowIndex; i < startRowIndex + numberOfRows; i++)
            {
                if (this.rowByIndex.ContainsKey(i))
                {
                    this.rowByIndex.Remove(i);
                }
            }

            // Shift the remaining rows up
            var keys = new List<int>(this.rowByIndex.Keys.Where(key => key >= startRowIndex + numberOfRows));
            foreach (var key in keys)
            {
                var value = this.rowByIndex[key];
                this.rowByIndex.Remove(key);
                this.rowByIndex[key - numberOfRows] = value;
            }

            // Shift the cells in the rows to the up
            var cellKeys = new List<(int, int)>(this.CellByCoordinates.Keys.Where(key => key.Item1 >= startRowIndex + numberOfRows));
            foreach (var key in cellKeys)
            {
                var cell = this.CellByCoordinates[key];
                this.CellByCoordinates.Remove(key);
                this.CellByCoordinates[(key.Item1 - numberOfRows, key.Item2)] = cell;
            }
        }


        public void InsertColumns(int startColumnIndex, int numberOfColumns) => throw new NotImplementedException();

        public void DeleteColumns(int startColumnIndex, int numberOfColumns)
        {
            if (startColumnIndex < 0 || numberOfColumns < 0)
            {
                throw new ArgumentException("Start index and number of columns must be non-negative.");
            }

            // Remove the columns from the dictionary
            for (int i = startColumnIndex; i < startColumnIndex + numberOfColumns; i++)
            {
                if (this.columnByIndex.ContainsKey(i))
                {
                    this.columnByIndex.Remove(i);
                }
            }

            // Shift the remaining columns to the left
            var keys = new List<int>(this.columnByIndex.Keys.Where(key => key >= startColumnIndex + numberOfColumns));
            foreach (var key in keys)
            {
                var value = this.columnByIndex[key];
                this.columnByIndex.Remove(key);
                this.columnByIndex[key - numberOfColumns] = value;
            }

            // Shift the cells in the columns to the left
            var cellKeys = new List<(int, int)>(this.CellByCoordinates.Keys.Where(key => key.Item2 >= startColumnIndex + numberOfColumns));
            foreach (var key in cellKeys)
            {
                var cell = this.CellByCoordinates[key];
                this.CellByCoordinates.Remove(key);
                this.CellByCoordinates[(key.Item1, key.Item2 - numberOfColumns)] = cell;
            }
        }

        public Range? GetRange(string? cell1, string? cell2 = null)
        {
            if (string.IsNullOrWhiteSpace(cell1))
            {
                return null;
            }

            
            if (cell1.Length == 5 && cell1[2] == ':' && cell2 == null)
            {
                cell2 = cell1.Substring(3);
                cell1 = cell1.Substring(0, 2);
            } else if (cell1.Length == 3 && cell1[1] == ':' && cell2 == null) {
                cell2 = cell1[2] + "1048576"; // Excel has 1,048,576 rows.
                cell1 = cell1[0] + "1";
            } else if (cell1.Length == 3 && cell2.Length == 3 && cell1[1] == ':' && cell2[1] == ':')
            {
                cell2 = cell2[0] + "1048576";
                cell1 = cell1[0] + "1";
            }
            else if (cell2 == null)
            {
                cell2 = cell1;
            }

            var parsedCell1 = ParseCellName(cell1);
            var parsedCell2 = ParseCellName(cell2);

            if (!parsedCell1.HasValue || !parsedCell2.HasValue)
            {
                return null;
            }

            var (row1, column1) = parsedCell1.Value;
            var (row2, column2) = parsedCell2.Value;

            if (row1 < 0 || column1 < 0 || row2 < 0 || column2 < 0)
            {
                return null;
            }

            return new Range(row1, column1, row2 - row1 + 1, column2 - column1 + 1, this);

            (int, int)? ParseCellName(string cellName)
            {
                if (string.IsNullOrWhiteSpace(cellName))
                {
                    return null;
                }

                var column = 0;
                var row = 0;

                var i = 0;
                while (i < cellName.Length && char.IsLetter(cellName[i]))
                {
                    column = column * 26 + cellName[i] - 'A' + 1;
                    i++;
                }

                if (i < cellName.Length && char.IsDigit(cellName[i]))
                {
                    try
                    {
                        row = int.Parse(cellName.Substring(i)) - 1;
                    }
                    catch (FormatException)
                    {
                        return null;  // Return null if the row number is not a valid number.
                    }
                }
                else
                {
                    return null;  // Return null if the cell name doesn't contain a row number.
                }

                return (row, column - 1);
            }
        }

        public Range GetUsedRange()
        {
            var minRow = this.CellByCoordinates.Keys.Min(key => key.Item1);
            var maxRow = this.CellByCoordinates.Keys.Max(key => key.Item1);
            var minColumn = this.CellByCoordinates.Keys.Min(key => key.Item2);
            var maxColumn = this.CellByCoordinates.Keys.Max(key => key.Item2);

            return new Range(minRow, minColumn, maxRow - minRow + 1, maxColumn - minColumn + 1, this);
        }

        public Range FrozenRange { get; private set; }

        public void FreezePanes(Range range)
        {
            // Check if the range is null
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            // Set the frozen range
            this.FrozenRange = range;
        }

        public void UnfreezePanes()
        {
            this.FrozenRange = null;
        }


        public void SaveAsPdf(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {

            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            var fullName = file.FullName;
            fullName = Path.ChangeExtension(fullName, "pdf");
            if (File.Exists(fullName) && !overwriteExistingFile)
            {
                throw new IOException("File already exists");
            }
            if (this.CellByCoordinates.Count == 0)
            {
                throw new COMException("Cannot save an empty file as PDF.");
            }
            using (FileStream fs = new FileStream(fullName, FileMode.Create))
            {
                using (BinaryWriter bw = new BinaryWriter(fs))
                {
                    for (int i = 0; i <= 0; i++)
                    {
                        bw.Write((byte)0);
                    }
                }
            }
        }

        public void SaveAsXps(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
                        if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            var fullName = file.FullName;
            fullName = Path.ChangeExtension(fullName, "xps");
            if (File.Exists(fullName) && !overwriteExistingFile)
            {
                throw new IOException("File already exists");
            }
            if (this.CellByCoordinates.Count == 0)
            {
                throw new COMException("Cannot save an empty file as XPS.");
            }
            using (FileStream fs = new FileStream(fullName, FileMode.Create))
            {
                using (BinaryWriter bw = new BinaryWriter(fs))
                {
                    for (int i = 0; i <= 0; i++)
                    {
                        bw.Write((byte)0);
                    }
                }
            }
        }

        public void SetPrintArea(Range range = null)
        {
        }

        public void HideInputMessage(ICell cell, bool clearInputMessage = false) { }

        public void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true) { }

        public void SetPageSetup(PageSetup pageSetup) { }

        public void AutoFit() { }

        public void SetChartObjectDataLabels(
            object chartObject,
            int seriesIndex,
            object seriesXValues,
            object seriesValues,
            bool showValues,
            bool showRange,
            string chartFieldRange,
            float baselineOffset,
            bool bold,
            bool italic,
            Color foreColor,
            object size)
        {
        }

        public void SetChartObjectSourceData(object chartObject, object pivotTable)
        {
        }

        public void AddHyperLink(string uri, ICell cell) { }
    }
}
