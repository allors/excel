// <copyright file="Cell.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;
    using System.Globalization;

    public class Cell(IWorksheet worksheet, Row row, Column column) : ICell
    {
        private static readonly IValueConverter DefaultValueConverter = new DefaultValueConverter();

        // the state of this when it is created
        private bool touched;

        private object? value;
        private Style? style;
        private string? formula;
        private string? numberFormat;
        private Range? options;

        private IValueConverter? valueConverter;
        private string? comment;

        Excel.IWorksheet ICell.Worksheet => this.Worksheet;

        public IWorksheet Worksheet { get; } = worksheet;

        IRow ICell.Row => this.Row;

        public Row Row { get; internal set; } = row;

        IColumn ICell.Column => this.Column;

        public Column Column { get; internal set; } = column;

        object? ICell.Value { get => this.Value; set => this.Value = value; }

        string? ICell.Formula { get => this.Formula; set => this.Formula = value; }

        public object? Value
        {
            get => this.value;
            set
            {
                // When we init the value with Null, we still want to be involved!
                if (!this.touched || !Equals(this.value, value))
                {
                    this.Worksheet.AddDirtyValue(this);
                    this.value = value;
                    this.touched = true;
                }
            }
        }

        public string? ValueAsString => Convert.ToString(this.Value, CultureInfo.CurrentCulture);

        public string? Formula
        {
            get => this.formula;
            set
            {
                if (!this.touched || !Equals(this.formula, value))
                {
                    this.Worksheet.AddDirtyFormula(this);
                    this.formula = value;
                    this.touched = true;
                }
            }
        }

        public string? Comment
        {
            get => this.comment;
            set
            {
                if (Equals(this.comment, value))
                {
                    return;
                }

                this.Worksheet.AddDirtyComment(this);
                this.comment = value;
            }
        }

        public Style? Style
        {
            get => this.style;
            set
            {
                if (this.style == null && value == null)
                {
                    return;
                }

                if (this.style != null && value != null && this.style.Equals(value))
                {
                    return;
                }

                this.Worksheet.AddDirtyStyle(this);
                this.style = value;
            }
        }

        public string? NumberFormat
        {
            get => this.numberFormat;
            set
            {
                if (!Equals(this.numberFormat, value))
                {
                    this.Worksheet.AddDirtyNumberFormat(this);
                    this.numberFormat = value;
                }
            }
        }

        public IValueConverter ValueConverter
        {
            get => this.valueConverter ?? DefaultValueConverter;
            set => this.valueConverter = value;
        }

        public Range? Options
        {
            get => this.options;
            set
            {
                if (!Equals(this.options, value))
                {
                    this.Worksheet.AddDirtyOptions(this);
                    this.options = value;
                }
            }
        }

        public bool IsRequired { get; set; }

        public bool HideInCellDropdown { get; set; }

        public override string ToString() => $"{this.Row}:{this.Column}";

        public bool UpdateValue(object rawExcelValue)
        {
            var excelValue = this.ValueConverter.Convert(this, rawExcelValue);
            var update = !Equals(this.value, excelValue);

            if (update)
            {
                this.value = excelValue;
            }

            return update;
        }

        public void Clear()
        {
            this.Value = string.Empty;
            this.Formula = string.Empty;
            this.Style = null;
            this.NumberFormat = null;
        }

        public object? Tag { get; set; }
    }
}
