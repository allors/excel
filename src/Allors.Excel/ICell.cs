// <copyright file="ICell.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Dynamic;

namespace Allors.Excel
{
    public interface ICell
    {
        IWorksheet Worksheet { get; }

        IRow Row { get; }

        IColumn Column { get; }

        object Value { get; set; }

        string ValueAsString { get; }

        string Formula { get; set; }

        Range Options { get; set; }

        bool IsRequired { get; set; }

        bool HideInCellDropdown { get; set; }

        string Comment { get; set; }

        Style Style { get; set; }

        string NumberFormat { get; set; }

        IValueConverter ValueConverter { get; set; }

        void Clear();

        object Tag { get; set; }       
    }
}
