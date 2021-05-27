// <copyright file="Range.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;

namespace Allors.Excel
{
    public class Range
    {
        public override string ToString()
        {
            return $"{Worksheet?.Name ?? "null"}!{Name ?? "null"} (row:{Row}, col:{Column}, rows:{Rows}, cols:{Columns})";
        }


        public Range(int row, int column, int? rows = null, int? columns = null, IWorksheet worksheet = null, string name = null)
        {
            Row = row;
            Column = column;
            Rows = rows;
            Columns = columns;
            Worksheet = worksheet;
            Name = name;

            if (Columns == null && Rows == null && string.IsNullOrEmpty(Name))
            {
                throw new ArgumentException("Either Columns or Rows, or Name is required.");
            }
        }

        public IWorksheet Worksheet { get; }

        /// <summary>
        /// Gets the name of the Range
        /// </summary>
        public string Name { get; }

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
