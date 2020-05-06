// <copyright file="IWorkbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;

namespace Allors.Excel
{
    public interface IWorkbook
    {
        bool IsActive { get; }

        IWorksheet[] Worksheets { get; }

        void Close(bool? saveChanges = null, string fileName = null);

        IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null);

        IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet);

        Range[] GetNamedRanges();

        /// <summary>
        /// Adds a NamedRange scoped to the Workbook
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        void SetNamedRange(string name, Range range);
    }
}
