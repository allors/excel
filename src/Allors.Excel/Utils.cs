// <copyright file="Range.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;

    public static class Utils
    {
        public static string ExcelColumnFromNumber(int column)
        {
            var columnString = string.Empty;
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                var currentLetterNumber = (columnNumber - 1) % 26;
                var currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        public static int ExcelColumnFromLetter(string column)
        {
            var retVal = 0;
            var col = column.ToUpper();
            for (var iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                var colPiece = col[iChar];
                var colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }
    }
}
