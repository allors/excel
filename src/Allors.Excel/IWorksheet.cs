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
        /// <summary>
        /// Event raised when cell values change.
        /// </summary>
        event EventHandler<CellChangedEvent> CellsChanged;

        /// <summary>
        /// Event raised when the sheet is activated. It becomes the active sheet.
        /// </summary>
        /// <returns>
        /// the name of the activated sheet.
        /// </returns>
        event EventHandler<string> SheetActivated;

        /// <summary>
        /// Gets the workbook for this worksheet
        /// </summary>
        IWorkbook Workbook { get; }

        /// <summary>
        /// Gets or sets the name of this worksheet.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// Gets the index of the sheet inside the workbook
        /// </summary>
        int Index { get; }

        /// <summary>
        /// Gets or sets the active worksheet.
        /// </summary>
        bool IsActive { get; set; }

        /// <summary>
        /// Indexer for getting the Row by index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        IRow Row(int index);

        /// <summary>
        /// Indexer for getting the column by index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        IColumn Column(int index);

        /// <summary>
        /// Indexer for getting the Cell by Row and Column index. If the cell does not exist, one will be created.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>
        /// the existing or newly created cell at this coordinate
        /// </returns>
        ICell this[int row, int column]
        {
            get;
        }

        /// <summary>
        /// Indexer for getting the Cell by Row and Column index. If the cell does not exist, one will be created.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>
        /// the existing or newly created cell at this coordinate
        /// </returns>
        ICell this[(int, int) coordinates]
        {
            get;
        }

        /// <summary>
        /// Flushes all ICells properties to the underlying implementation (interop or headless).
        /// Only cells that need an update will be updated. Update are performed in block ranges.
        /// Properties:
        ///  - Numberformats
        ///  - Values
        ///  - Formulas
        ///  - Comments
        ///  - Styles
        ///  - Options
        ///  - Hidden Rows
        /// </summary>
        /// <returns></returns>
        Task Flush();

        /// <summary>
        /// Updates the data for each pivottable in the worksheet to the newRange
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        Task RefreshPivotTables();

        /// <summary>
        /// Inserts a picture (via url) in the given rectangle
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="rectangle"></param>
        void AddPicture(string uri, Rectangle rectangle);

        /// <summary>
        /// Gets the rectange from the namedRange (the location of the namedrange in the worksheet)
        /// </summary>
        /// <param name="namedRange"></param>
        /// <returns></returns>
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

        void AutoFit();

        void SetChartObjectSourceData(object chartObject, object pivotTable);

        bool IsVisible { get; set; }

        /// <summary>
        /// Freeze the pane at the given range (row, column)
        /// </summary>
        /// <param name="range"></param>
        void FreezePanes(Excel.Range range);

        /// <summary>
        /// Removes the frozen pane from the worksheet
        /// </summary>
        void UnfreezePanes();

        /// <summary>
        /// True when the sheet has pane that is frozen, otherwise, false
        /// </summary>
        bool HasFreezePanes { get; }

        /// <summary>
        /// Saves the sheet as a PDF to the file, with the given parameters.
        /// </summary>
        /// <param name="file">A file at a certain location</param>
        /// <param name="overwriteExistingFile">true when we want to overwrite an existing file. Default is false</param>
        /// <param name="openAfterPublish">true when we want to open the pdf after it has been created. Default is false.</param>
        /// <param name="ignorePrintAreas">true if we want to print the entire sheet. Default is true</param>
        void SaveAsPDF(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true);


        /// <summary>
        /// Saves the sheet as a XPS to the file, with the given parameters.
        /// </summary>
        /// <param name="file">A file at a certain location</param>
        /// <param name="overwriteExistingFile">true when we want to overwrite an existing file. Default is false</param>
        /// <param name="openAfterPublish">true when we want to open the pdf after it has been created. Default is false.</param>
        /// <param name="ignorePrintAreas">true if we want to print the entire sheet. Default is true</param>
        void SaveAsXPS(FileInfo file, bool overwriteExistingFile = false, bool openAfterPublish = false, bool ignorePrintAreas = true);

        /// <summary>
        /// Sets the PrintArea to the given range, or to the entire sheet when range is null.
        /// </summary>
        /// <param name="range"></param>
        void SetPrintArea(Excel.Range range = null);

        /// <summary>
        /// Adds or updates the worksheet's custom properties with the given customproperties keyvalue pairs
        /// </summary>
        /// <param name="properties"></param>
        void SetCustomProperties(CustomProperties properties);

        /// <summary>
        /// Gets the CustomProperties from the worksheets
        /// </summary>
        /// <returns></returns>
        CustomProperties GetCustomProperties();

        /// <summary>
        /// Sets the inputMessage that will be displayed when the cell is selected (aka Help text)
        /// </summary>
        void SetInputMessage(ICell cell, string message, string title = null, bool showInputMessage = true);

        /// <summary>
        /// Hides or optionally remove the inputtext from a cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="clearInputMessage"></param>
        void HideInputMessage(ICell cell, bool clearInputMessage = false);


        /// <summary>
        /// PageSetup for printing purpose: Orientation, PageSize, Header and Footer
        /// </summary>
        /// <param name="pageSetup"></param>
        void SetPageSetup(PageSetup pageSetup);
    }
}
