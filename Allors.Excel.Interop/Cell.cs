// <copyright file="Cell.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Embedded
{
    public class Cell : ICell
    {
        // the state of this when it is created
        private bool touched = false;

        private object value;
        private Style style;
        private string formula;
        private string numberFormat;
        private Range options;

        private IValueConverter valueConverter;
        private readonly IValueConverter defaultValueConverter = new DefaultValueConverter();
        private string comment;

        public Cell(IEmbeddedWorksheet worksheet, Row row, Column column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        IWorksheet ICell.Worksheet => this.Worksheet;

        public IEmbeddedWorksheet Worksheet { get; }

        IRow ICell.Row => this.Row;

        public Row Row { get; }

        IColumn ICell.Column => this.Column;

        public Column Column { get; }

        object ICell.Value { get => this.Value; set => this.Value = value; }

        string ICell.Formula { get => this.Formula; set => this.Formula = value; }

        public object Value
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

        public string Formula
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

        public string Comment
        {
            get => comment;
            set
            {
                if (!Equals(this.comment, value))
                {
                    this.Worksheet.AddDirtyComment(this);
                    this.comment = value;                

                }
            }
        }

        public Style Style
        {
            get => style;
            set
            {
                if (!this.style?.Equals(value) ?? value != null)
                {
                    this.Worksheet.AddDirtyStyle(this);
                    this.style = value;
                }
            }
        }

        public string NumberFormat
        {
            get => numberFormat;
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
            get => valueConverter ?? this.defaultValueConverter;
            set => valueConverter = value;
        }

        public Range Options
        {
            get => options;
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

        public override string ToString()
        {
            return $"{Row}:{Column}";
        }

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
    }
}
