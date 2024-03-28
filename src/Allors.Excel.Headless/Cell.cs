// <copyright file="Cell.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System;
    using System.Globalization;

    public class Cell(Worksheet worksheet, Row row, Column column) : ICell
    {
        public IWorksheet Worksheet { get; } = worksheet;

        IRow ICell.Row => this.Row;

        public Row Row { get; } = row;

        IColumn ICell.Column => this.Column;

        public Column Column { get; } = column;

        public object Value { get; set; }

        public string Comment { get; set; }

        public Style Style { get; set; }

        public string NumberFormat { get; set; }

        public IValueConverter ValueConverter { get; set; }

        public Range Options { get; set; }

        public bool IsRequired { get; set; }

        public bool HideInCellDropdown { get; set; }

        public string Formula { get; set; }

        public string ValueAsString => Convert.ToString(this.Value, CultureInfo.CurrentCulture);

        public void Clear()
        {
            this.Value = string.Empty;
            this.NumberFormat = null;
            this.Style = null;
        }
    }
}
