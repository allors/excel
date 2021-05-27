// <copyright file="Cell.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Globalization;

namespace Allors.Excel.Interop
{
    public class Cell : ICell
    {
        // the state of this when it is created
        private bool touched;

        private object value;
        private Style style;
        private string formula;
        private string numberFormat;
        private Range options;

        private IValueConverter valueConverter;
        private readonly IValueConverter defaultValueConverter = new DefaultValueConverter();
        private string comment;

        public Cell(IWorksheet worksheet, Row row, Column column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        Excel.IWorksheet ICell.Worksheet => Worksheet;

        public IWorksheet Worksheet { get; }

        IRow ICell.Row => Row;

        public Row Row { get; internal set; }

        IColumn ICell.Column => Column;

        public Column Column { get; internal set; }

        object ICell.Value { get => Value; set => Value = value; }

        string ICell.Formula { get => Formula; set => Formula = value; }

        public object Value
        {
            get => value;
            set
            {
                // When we init the value with Null, we still want to be involved!
                if (!touched || !Equals(this.value, value))
                {
                    Worksheet.AddDirtyValue(this);
                    this.value = value;
                    touched = true;
                }
            }
        }

        public string ValueAsString => Convert.ToString(Value, CultureInfo.CurrentCulture);

        public string Formula
        {
            get => formula;
            set
            {
                if (!touched || !Equals(formula, value))
                {
                    Worksheet.AddDirtyFormula(this);
                    formula = value;
                    touched = true;
                }
            }
        }

        public string Comment
        {
            get => comment;
            set
            {
                if (!Equals(comment, value))
                {
                    Worksheet.AddDirtyComment(this);
                    comment = value;                

                }
            }
        }

        public Style Style
        {
            get => style;
            set
            {
                if (!style?.Equals(value) ?? value != null)
                {
                    Worksheet.AddDirtyStyle(this);
                    style = value;
                }
            }
        }

        public string NumberFormat
        {
            get => numberFormat;
            set
            {
                if (!Equals(numberFormat, value))
                {
                    Worksheet.AddDirtyNumberFormat(this);
                    numberFormat = value;
                }
            }
        }

        public IValueConverter ValueConverter
        {
            get => valueConverter ?? defaultValueConverter;
            set => valueConverter = value;
        }

        public Range Options
        {
            get => options;
            set
            {
                if (!Equals(options, value))
                {
                    Worksheet.AddDirtyOptions(this);
                    options = value;
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
            var excelValue = ValueConverter.Convert(this, rawExcelValue);
            var update = !Equals(value, excelValue);

            if (update)
            {
                value = excelValue;
            }

            return update;
        }

        public void Clear()
        {
            Value = string.Empty;
            Formula = string.Empty;
            Style = null;
            NumberFormat = null;
        }
        
        public object Tag { get; set; }
    }
}
