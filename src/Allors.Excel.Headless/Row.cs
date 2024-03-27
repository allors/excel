// <copyright file="Row.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    public class Row(Worksheet worksheet, int index) : IRow
    {
        public IWorksheet Worksheet { get; } = worksheet;

        public int Index { get; } = index;

        public bool Hidden { get; set; }
    }
}
