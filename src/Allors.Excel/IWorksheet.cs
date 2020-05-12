// <copyright file="IWorksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Drawing;
using System.Threading.Tasks;

namespace Allors.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        event EventHandler<string> SheetActivated;

        IWorkbook Workbook { get; }

        string Name { get; set; }

        bool IsActive { get; }

        IRow Row(int index);

        IColumn Column(int index);

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();

        Task RefreshPivotTables(string newRange);

        void AddPicture(string uri, Rectangle rectangle);       

        Rectangle GetRectangle(string namedRange);

        Range[] GetNamedRanges();

        /// <summary>
        /// Adds a NamedRange scoped to the Worksheet
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        void SetNamedRange(string name, Range range);

        void InsertRows(int startRowIndex, int numberOfRows);

        void DeleteRows(int startRowIndex, int numberOfRows);

        void InsertColumns(int startColumnIndex, int numberOfColumns);

        void DeleteColumns(int startColumnIndex, int numberOfColumns);
    }
}
