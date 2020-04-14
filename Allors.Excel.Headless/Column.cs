// <copyright file="Column.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    public class Column : IColumn
    {
        public Column(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }
    }
}
