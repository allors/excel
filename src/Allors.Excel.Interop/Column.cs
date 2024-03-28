// <copyright file="Column.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;

    public sealed class Column(Worksheet worksheet, int index) : IColumn, IComparable<Column>
    {
        public Excel.IWorksheet Worksheet { get; } = worksheet;

        public int Index { get; internal set; } = index;

        public int CompareTo(Column other) => this.Index.CompareTo(other.Index);
    }
}
