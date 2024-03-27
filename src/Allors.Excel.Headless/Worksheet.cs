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

        public bool HasFreezePanes => throw new NotImplementedException();

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
        public Dictionary<string, Range> NamedRangeByName { get; } = new Dictionary<string, Range>();

        public Range[] GetNamedRanges()
        {
            return this.NamedRangeByName.Values.ToArray();
        }

        public Rectangle GetRectangle(string namedRange) => Rectangle.Empty;

        public void SetNamedRange(string name, Range range) => throw new NotImplementedException();

        public void InsertRows(int startRowIndex, int numberOfRows) => throw new NotImplementedException();

        public void DeleteRows(int startRowIndex, int numberOfRows) => throw new NotImplementedException();

        public void InsertColumns(int startColumnIndex, int numberOfColumns) => throw new NotImplementedException();

        public void DeleteColumns(int startColumnIndex, int numberOfColumns) => throw new NotImplementedException();

        public Range GetRange(string cell1, string cell2 = null) => throw new NotImplementedException();

        public Range GetUsedRange() => throw new NotImplementedException();

        public Range GetUsedRange(string column) => throw new NotImplementedException();

        public Range GetUsedRange(int row) => throw new NotImplementedException();

        public void FreezePanes(Range range) => throw new NotImplementedException();

        public void UnfreezePanes() => throw new NotImplementedException();

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
            Document.Create(container =>
                                                                                                                                                          {
                                                                                                                                                              container.Page(page =>
                                                                                                                                                              {
                                                                                                                                                                  page.Size(PageSizes.A4);
                                                                                                                                                                  page.Margin(2, Unit.Centimetre);
                                                                                                                                                                  page.PageColor(Colors.White);
                                                                                                                                                                  page.DefaultTextStyle(x => x.FontSize(20));

                                                                                                                                                                  page.Header()
                                                                                                                                                                      .Text("Hello PDF!")
                                                                                                                                                                      .SemiBold().FontSize(36).FontColor(Colors.Blue.Medium);

                                                                                                                                                                  page.Content()
                                                                                                                                                                      .PaddingVertical(1, Unit.Centimetre)
                                                                                                                                                                      .Column(x =>
                                                                                                                                                                      {
                                                                                                                                                                          x.Spacing(20);


                                                                                                                                                                          foreach (var cell in this.CellByCoordinates)
                                                                                                                                                                          {
                                                                                                                                                                              x.Item().Text(cell.Value.Value);
                                                                                                                                                                          }


                                                                                                                                                                      });

                                                                                                                                                                  page.Footer()
                                                                                                                                                                      .AlignCenter()
                                                                                                                                                                      .Text(x =>
                                                                                                                                                                      {
                                                                                                                                                                          x.Span("Page ");
                                                                                                                                                                          x.CurrentPageNumber();
                                                                                                                                                                      });
                                                                                                                                                              });
                                                                                                                                                          })
  .GeneratePdf(fullName);
        }

        public void SaveAsXps(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true)
        {
            SaveAsXPS(file, overwriteExistingFile, openAfterPublish, ignorePrintAreas);
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
