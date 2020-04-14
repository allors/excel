// <copyright file="IColumn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    public interface IColumn
    {
        IWorksheet Worksheet { get; }

        int Index { get; }
    }
}
