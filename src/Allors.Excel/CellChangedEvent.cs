// <copyright file="CellChangedEvent.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;

namespace Allors.Excel
{
    public class CellChangedEvent : EventArgs
    {
        public ICell[] Cells { get; }

        public CellChangedEvent(ICell[] cells) => this.Cells = cells;
    }
}
