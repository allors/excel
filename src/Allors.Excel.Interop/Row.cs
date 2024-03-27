// <copyright file="Row.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;

    public class Row(Worksheet worksheet, int index) : IRow, IComparable<Row>
    {
        private bool hidden;

        Excel.IWorksheet IRow.Worksheet => this.Worksheet;

        public Worksheet Worksheet { get; } = worksheet;

        public int Index { get; internal set; } = index;

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
