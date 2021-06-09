// <copyright file="Row.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    public class Row : IRow
    {
        public Row(Worksheet worksheet, int index)
        {
            this.Worksheet = worksheet;
            this.Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }

        public bool Hidden { get; set; }
    }
}
