// <copyright file="Range.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;

namespace Allors.Excel
{
    public class Range
    {
        public Range(int row, int column, int? rows = null, int? columns = null, IWorksheet worksheet = null, string name = null)
        {
            this.Row = row;
            this.Column = column;
            this.Rows = rows;
            this.Columns = columns;
            this.Worksheet = worksheet;
            this.Name = name;

            if (this.Columns == null && this.Rows == null && string.IsNullOrEmpty(this.Name))
            {
                throw new ArgumentException("Either Columns or Rows, or Name is required.");
            }
        }

        public IWorksheet Worksheet { get; }

        public string Name { get; }

        public int Row { get; }

        public int Column { get; }

        public int? Rows { get; }

        public int? Columns { get; }
    }
}
