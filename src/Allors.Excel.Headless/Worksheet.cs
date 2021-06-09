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
    using System.Threading.Tasks;

    public class Worksheet : IWorksheet
    {
        private readonly Dictionary<int, Row> rowByIndex;

        private readonly Dictionary<int, Column> columnByIndex;

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
        public event EventHandler<string> SheetActivated;

        public string Name { get; set; }

        public int Index => throw new NotImplementedException();


        public bool IsActive { get; set; }

        public Dictionary<(int, int), Cell> CellByCoordinates { get; }

        public bool IsVisible { get; set; }

        public bool HasFreezePanes => throw new NotImplementedException();


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

        public async Task Flush()
        {
            await Task.CompletedTask;
        }

        public void Activate()
        {
            foreach (var worksheet in this.Workbook.WorksheetList)
            {
                worksheet.IsActive = false;
            }

            this.IsActive = true;
        }

        public async Task RefreshPivotTables()
        {
            // strictly ui
            await Task.CompletedTask;
        }

        public void AddPicture(string uri, Rectangle rectangle)
        {
            // strictly ui
        }

        public Rectangle GetRectangle(string namedRange)
        {
            // strictly ui
            return Rectangle.Empty;
        }

        public Range[] GetNamedRanges()
        {
            throw new NotImplementedException();
        }

        public void SetNamedRange(string name, Range range)
        {
            throw new NotImplementedException();
        }

        public void InsertRows(int startRowIndex, int numberOfRows)
        {
            throw new NotImplementedException();
        }

        public void DeleteRows(int startRowIndex, int numberOfRows)
        {
            throw new NotImplementedException();
        }

        public void InsertColumns(int startColumnIndex, int numberOfColumns)
        {
            throw new NotImplementedException();
        }

        public void DeleteColumns(int startColumnIndex, int numberOfColumns)
        {
            throw new NotImplementedException();
        }

        public Range GetRange(string cell1, string cell2 = null)
        {
            throw new NotImplementedException();
        }

        public Range GetUsedRange()
        {
            throw new NotImplementedException();
        }

        public Range GetUsedRange(string column)
        {
            throw new NotImplementedException();
        }

        public Range GetUsedRange(int row)
        {
            throw new NotImplementedException();
        }

        public void FreezePanes(Range range)
        {
            throw new NotImplementedException();
        }

        public void UnfreezePanes()
        {
            throw new NotImplementedException();
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

        public void HideInputMessage(ICell cell, bool clearInputMessage = false)
        {
            throw new NotImplementedException();
        }

        public void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true)
        {
            throw new NotImplementedException();
        }

        public void SetPageSetup(PageSetup pageSetup)
        {
            throw new NotImplementedException();
        }

        public void AutoFit()
        {
        }

        public void SetChartObjectSourceData(object chartObject, object pivotTable)
        {
        }
    }
}
