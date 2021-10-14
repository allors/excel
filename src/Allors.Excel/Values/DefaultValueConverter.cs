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
                    if (@double < (double)decimal.MinValue)
                    {
                        return (double)decimal.MinValue;
                    }

                    if (@double > (double)decimal.MaxValue)
                    {
                        return (double)decimal.MaxValue;
                    }

                    return System.Convert.ToDecimal(excelValue);
                }
            }

            {
                if (cell.Value is int && excelValue is double @double)
                {
                    if (@double < int.MinValue)
                    {
                        return int.MinValue;
                    }

                    if (@double > int.MaxValue)
                    {
                        return int.MaxValue;
                    }

                    return System.Convert.ToInt32(excelValue);
                }
            }

            {
                if (cell.Value is DateTime && excelValue is double @double)
                {
                    return DateTime.FromOADate(@double);
                }
            }

            if (cell.Value is string && excelValue == null)
            {
                return string.Empty;
            }

            if (cell.Value is string && !(excelValue is string))
            {
                return excelValue.ToString();
            }

            return excelValue;
        }
    }
}
