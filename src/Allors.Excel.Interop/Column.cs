// <copyright file="Column.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;

    public class Column : IColumn, IComparable<Column>
    {
        public Column(Worksheet worksheet, int index)
        {
            this.Worksheet = worksheet;
            this.Index = index;
        }

        public Excel.IWorksheet Worksheet { get; }

        public int Index { get; internal set; }

        public int CompareTo(Column other)
        {
            return this.Index.CompareTo(other.Index);
        }
    }
}
