// <copyright file="DefaultValueConverter.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;
    using System.Globalization;

    /// <summary>
    /// Default <see cref="IValueConverter"/>. Coerces the value Excel produced for a cell into
    /// the runtime type the model already holds (<see cref="ICell.Value"/>), so the model keeps a
    /// stable type across edits.
    /// <para>
    /// It is <b>total</b>: it never throws. When an edit cannot be represented in the expected type
    /// (an unparsable string, an out-of-range number/date, NaN/Infinity) it keeps the cell's current
    /// value, i.e. the edit is rejected rather than corrupting the model. Clearing a cell yields
    /// <see langword="null"/> for value types and the empty string for text.
    /// </para>
    /// <para>
    /// Numeric/text formatting uses an injectable <see cref="IFormatProvider"/> (invariant by default,
    /// so model values do not vary with the machine's locale) and fractional-to-integer conversion uses
    /// an injectable <see cref="MidpointRounding"/> (away-from-zero by default, matching Excel's ROUND).
    /// </para>
    /// </summary>
    public class DefaultValueConverter : IValueConverter
    {
        private readonly IFormatProvider formatProvider;
        private readonly MidpointRounding midpointRounding;

        public DefaultValueConverter()
            : this(CultureInfo.InvariantCulture, MidpointRounding.AwayFromZero)
        {
        }

        public DefaultValueConverter(IFormatProvider formatProvider, MidpointRounding midpointRounding)
        {
            this.formatProvider = formatProvider ?? CultureInfo.InvariantCulture;
            this.midpointRounding = midpointRounding;
        }

        public object Convert(ICell cell, object excelValue)
        {
            // The runtime type of the current model value is the "expected" type to coerce toward.
            var current = cell.Value;

            switch (current)
            {
                case decimal _: return this.ToDecimal(excelValue, current);
                case int _: return this.ToInt(excelValue, current);
                case double _: return this.ToDouble(excelValue, current);
                case bool _: return this.ToBool(excelValue, current);
                case DateTime _: return this.ToDateTime(excelValue, current);
                case string _: return this.ToStringValue(excelValue);
                default: return excelValue; // no type hint (null/unknown) -> accept Excel's value as-is
            }
        }

        private object ToDecimal(object value, object current)
        {
            switch (value)
            {
                case null: return null;
                case decimal d: return d;
                case double d: return this.DoubleToDecimal(d, current);
                case int i: return (decimal)i;
                case long l: return (decimal)l;
                case bool b: return b ? 1m : 0m;
                case string s: return decimal.TryParse(s, NumberStyles.Any, this.formatProvider, out var r) ? r : current;
                default: return current;
            }
        }

        private object DoubleToDecimal(double d, object current)
        {
            if (double.IsNaN(d))
            {
                return current;
            }

            // (double)decimal.MaxValue rounds ABOVE decimal.MaxValue, so a plain Convert.ToDecimal
            // on it overflows; clamp with >= / <= (and absorb +/-Infinity here too).
            if (double.IsPositiveInfinity(d) || d >= (double)decimal.MaxValue)
            {
                return decimal.MaxValue;
            }

            if (double.IsNegativeInfinity(d) || d <= (double)decimal.MinValue)
            {
                return decimal.MinValue;
            }

            return (decimal)d;
        }

        private object ToInt(object value, object current)
        {
            switch (value)
            {
                case null: return null;
                case int i: return i;
                case double d: return this.DoubleToInt(d, current);
                case long l: return ClampToInt(l);
                case bool b: return b ? 1 : 0;
                case string s: return int.TryParse(s, NumberStyles.Any, this.formatProvider, out var r) ? r : current;
                default: return current;
            }
        }

        private object DoubleToInt(double d, object current)
        {
            if (double.IsNaN(d))
            {
                return current;
            }

            if (double.IsPositiveInfinity(d) || d >= int.MaxValue)
            {
                return int.MaxValue;
            }

            if (double.IsNegativeInfinity(d) || d <= int.MinValue)
            {
                return int.MinValue;
            }

            return (int)Math.Round(d, this.midpointRounding);
        }

        // Clamps a long into int range (matching the DoubleToInt overflow behaviour).
        private static int ClampToInt(long value)
        {
            if (value >= int.MaxValue)
            {
                return int.MaxValue;
            }

            if (value <= int.MinValue)
            {
                return int.MinValue;
            }

            return (int)value;
        }

        private object ToDouble(object value, object current)
        {
            switch (value)
            {
                case null: return null;
                case double d: return d;
                case int i: return (double)i;
                case long l: return (double)l;
                case decimal d: return (double)d;
                case bool b: return b ? 1d : 0d;
                case string s: return double.TryParse(s, NumberStyles.Any, this.formatProvider, out var r) ? r : current;
                default: return current;
            }
        }

        private object ToBool(object value, object current)
        {
            switch (value)
            {
                case null: return null;
                case bool b: return b;
                case double d: return d != 0d;
                case int i: return i != 0;
                case long l: return l != 0;
                case string s: return bool.TryParse(s, out var r) ? r : current;
                default: return current;
            }
        }

        private object ToDateTime(object value, object current)
        {
            switch (value)
            {
                case null: return null;
                case DateTime dt: return dt;
                case double d: return this.DoubleToDateTime(d, current);
                case string s: return DateTime.TryParse(s, this.formatProvider, DateTimeStyles.None, out var r) ? r : current;
                default: return current;
            }
        }

        private object DoubleToDateTime(double d, object current)
        {
            if (double.IsNaN(d) || double.IsInfinity(d))
            {
                return current;
            }

            try
            {
                return DateTime.FromOADate(d);
            }
            catch (ArgumentException)
            {
                // Outside the valid OLE Automation date range: reject the edit.
                return current;
            }
        }

        private object ToStringValue(object value)
        {
            switch (value)
            {
                case null: return string.Empty;
                case string s: return s;
                default: return System.Convert.ToString(value, this.formatProvider) ?? string.Empty;
            }
        }
    }
}
