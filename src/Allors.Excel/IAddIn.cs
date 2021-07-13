// <copyright file="IAddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    public interface IAddIn
    {
        IRibbon Ribbon { get; }

        IWorkbook[] Workbooks { get; }

        void DisplayAlerts(bool displayAlerts);
    }
}
