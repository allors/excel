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
    using System.Threading.Tasks;

    public class Worksheet : IWorksheet
    {
        private readonly Dictionary<int, Row> rowByIndex;

        private readonly Dictionary<int, Column> columnByIndex;
        private bool isActive;

        IWorkbook IWorksheet.Workbook => this.Workbook;

        public ICustomProperties CustomProperties { get; }

        public Workbook Workbook { get; }

        public Worksheet(Workbook workbook)
        {
            this.Workbook = workbook;

            this.rowByIndex = new Dictionary<int, Row>();
            this.columnByIndex = new Dictionary<int, Column>();
            this.CellByCoordinates = new Dictionary<(int, int), Cell>();
        }

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


        public Dictionary<(int, int), Cell> CellByCoordinates { get; }

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

        public Rectangle GetRectangle(string namedRange) => Rectangle.Empty;


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

        public Range GetRange(string cell1, string cell2 = null) => throw new NotImplementedException();

        public Range GetUsedRange() => throw new NotImplementedException();

        public Range GetUsedRange(string column) => throw new NotImplementedException();

        public Range GetUsedRange(int row) => throw new NotImplementedException();

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


        public void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {

        }

        public void SaveAsXPS(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {

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
