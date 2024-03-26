// <copyright file="Range.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;

    public class Range
    {
        public override string ToString() => $"{this.Worksheet?.Name ?? "null"}!{this.Name ?? "null"} (row:{this.Row}, col:{this.Column}, rows:{this.Rows}, cols:{this.Columns})";


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

        /// <summary>
        /// Gets the name of the Range
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets the start row index of the Range
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the start column index of the Range
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the number of rows in the Range
        /// </summary>
        public int? Rows { get; }

        /// <summary>
        /// Gets the number of columns in the Range
        /// </summary>
        public int? Columns { get; }
    }
}
