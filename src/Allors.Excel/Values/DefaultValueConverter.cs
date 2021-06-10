// <copyright file="DefaultValueConverter.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;

    public class DefaultValueConverter : IValueConverter
    {
        public object Convert(ICell cell, object excelValue)
        {
            {
                if (cell.Value is decimal && excelValue is double @double)
                {
                    return @double switch
                    {
                        < (double)decimal.MinValue => (double)decimal.MinValue,
                        > (double)decimal.MaxValue => (double)decimal.MaxValue,
                        _ => System.Convert.ToDecimal(excelValue)
                    };
                }
            }

            {
                if (cell.Value is int && excelValue is double @double)
                {
                    return @double switch
                    {
                        < int.MinValue => int.MinValue,
                        > int.MaxValue => int.MaxValue,
                        _ => System.Convert.ToInt32(excelValue)
                    };
                }
            }

            {
                if (cell.Value is DateTime && excelValue is double @double)
                {
                    return DateTime.FromOADate(@double);
                }
            }

            return cell.Value switch
            {
                string when excelValue == null => string.Empty,
                string when excelValue is not string => excelValue.ToString(),
                _ => excelValue
            };
        }
    }
}
