// <copyright file="Row.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;

    public class Row : IRow, IComparable<Row>
    {
        private bool hidden;

        public Row(Worksheet worksheet, int index)
        {
            this.Worksheet = worksheet;
            this.Index = index;
        }

        Excel.IWorksheet IRow.Worksheet => this.Worksheet;

        public Worksheet Worksheet { get; }

        public int Index { get; internal set; }

        bool IRow.Hidden { get => this.Hidden; set => this.Hidden = value; }

        public bool Hidden
        {
            get => this.hidden;
            set
            {
                this.hidden = value;
                this.Worksheet.AddDirtyRow(this);
            }
        }

        public int CompareTo(Row other) => this.Index.CompareTo(other.Index);
    }
}
