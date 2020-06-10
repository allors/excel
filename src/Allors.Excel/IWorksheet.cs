// <copyright file="IWorksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Threading.Tasks;

namespace Allors.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        event EventHandler<string> SheetActivated;

        IWorkbook Workbook { get; }

        string Name { get; set; }

        /// <summary>
        /// Gets or sets the active worksheet.
        /// </summary>
        bool IsActive { get; set; }

        IRow Row(int index);

        IColumn Column(int index);

        ICell this[int row, int column]
        {
            get;
        }

        ICell this[(int, int) coordinates]
        {
            get;
        }

        Task Flush();

        Task RefreshPivotTables(string newRange);

        void AddPicture(string uri, Rectangle rectangle);

        Rectangle GetRectangle(string namedRange);

        /// <summary>
        /// Gets all named Ranges in this worksheet scope.
        /// </summary>
        /// <returns></returns>
        Range[] GetNamedRanges();

        /// <summary>
        /// Adds or updates a NamedRange scoped to the Worksheet
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        void SetNamedRange(string name, Range range);

        /// <summary>
        /// Insert new rows in this worksheet. Rows below will be shifted down.
        /// </summary>
        /// <param name="startRowIndex"></param>
        /// <param name="numberOfRows"></param>
        void InsertRows(int startRowIndex, int numberOfRows);

        /// <summary>
        /// Delete rows in this worksheet. Rows below will be shifted up.
        /// </summary>
        /// <param name="startRowIndex"></param>
        /// <param name="numberOfRows"></param>
        void DeleteRows(int startRowIndex, int numberOfRows);

        /// <summary>
        /// Insert new columns in this worksheet. Columns to the right will be shifted to the right.
        /// </summary>
        /// <param name="startColumnIndex"></param>
        /// <param name="numberOfColumns"></param>
        void InsertColumns(int startColumnIndex, int numberOfColumns);

        /// <summary>
        /// Delete columns in this worksheet. Columns to the right will be shifted to the left.
        /// </summary>
        /// <param name="startColumnIndex"></param>
        /// <param name="numberOfColumns"></param>
        void DeleteColumns(int startColumnIndex, int numberOfColumns);

        Range GetRange(string cell1, string cell2 = null);

        Range GetUsedRange();

        /// <summary>
        /// column equals the excel columns A,B,C, ...
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        Range GetUsedRange(string column);

        /// <summary>
        /// row equals the zero-based index of excel rows (so 1 less than the excel rowindex)
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        Range GetUsedRange(int row);

        bool IsVisible { get; set; }

        /// <summary>
        /// </summary>
        /// <param name="range"></param>
        void FreezePanes(Excel.Range range);

        void UnfreezePanes();

        bool HasFreezePanes { get; }

        void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true);

        void SaveAsXPS(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true);

        /// <summary>
        /// Sets the PrintArea to the given range, or to the entire sheet when range is null.
        /// </summary>
        /// <param name="range"></param>
        void SetPrintArea(Excel.Range range = null);

        void SetCustomProperties(CustomProperties properties);

        CustomProperties GetCustomProperties();

        /// <summary>
        /// Sets the inputMessage that will be displayed when the cell is selected (aka Help text)
        /// </summary>
        void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true);

        void HideInputMessage(ICell cell, bool clearInputMessage = false);
    }
}
