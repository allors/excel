// <copyright file="IRow.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    public interface IRow
    {
        IWorksheet Worksheet { get; }

        int Index { get; }

        bool Hidden { get; set; }
    }
}
