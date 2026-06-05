// <copyright file="DefaultValueConverterTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless.Tests
{
    using System;
    using System.Globalization;
    using System.Threading;
    using Allors.Excel;
    using Moq;
    using Xunit;

    public class DefaultValueConverterTests
    {
        private readonly DefaultValueConverter converter = new DefaultValueConverter();

        private static ICell Cell(object current)
        {
            var mock = new Mock<ICell>();
            mock.SetupGet(c => c.Value).Returns(current);
            return mock.Object;
        }

        // ---------------- decimal target ----------------

        [Fact]
        public void DecimalInRangeReturnsDecimal()
        {
            var result = this.converter.Convert(Cell(1m), 2.5d);
            Assert.IsType<decimal>(result);
            Assert.Equal(2.5m, result);
        }

        [Fact]
        public void DecimalOverflowClampsToDecimalMaxKeepingType()
        {
            var result = this.converter.Convert(Cell(1m), 1e30);
            Assert.IsType<decimal>(result);
            Assert.Equal(decimal.MaxValue, result);
        }

        [Fact]
        public void DecimalUnderflowClampsToDecimalMin()
        {
            var result = this.converter.Convert(Cell(1m), -1e30);
            Assert.IsType<decimal>(result);
            Assert.Equal(decimal.MinValue, result);
        }

        [Fact]
        public void DecimalBoundaryDoubleDoesNotThrow()
        {
            // (double)decimal.MaxValue is GREATER than decimal.MaxValue, so Convert.ToDecimal
            // on it throws OverflowException; the converter must clamp instead.
            var boundary = (double)decimal.MaxValue;
            var result = this.converter.Convert(Cell(1m), boundary);
            Assert.IsType<decimal>(result);
            Assert.Equal(decimal.MaxValue, result);
        }

        [Fact]
        public void DecimalNaNKeepsCurrent()
        {
            Assert.Equal(7m, this.converter.Convert(Cell(7m), double.NaN));
        }

        [Fact]
        public void DecimalFromBoolCoerces()
        {
            Assert.Equal(1m, this.converter.Convert(Cell(0m), true));
        }

        [Fact]
        public void DecimalFromUnparsableStringKeepsCurrent()
        {
            Assert.Equal(5m, this.converter.Convert(Cell(5m), "abc"));
        }

        // ---------------- int target ----------------

        [Theory]
        [InlineData(2.5, 3)]
        [InlineData(3.5, 4)]
        [InlineData(-2.5, -3)]
        [InlineData(2.4, 2)]
        public void IntRoundsAwayFromZero(double input, int expected)
        {
            var result = this.converter.Convert(Cell(0), input);
            Assert.IsType<int>(result);
            Assert.Equal(expected, result);
        }

        [Fact]
        public void IntOverflowClampsToIntMaxKeepingType()
        {
            var result = this.converter.Convert(Cell(0), 1e30);
            Assert.IsType<int>(result);
            Assert.Equal(int.MaxValue, result);
        }

        [Fact]
        public void IntNaNKeepsCurrent()
        {
            Assert.Equal(9, this.converter.Convert(Cell(9), double.NaN));
        }

        // ---------------- DateTime target ----------------

        [Fact]
        public void DateTimeFromOADateRoundTrips()
        {
            var oaDate = new DateTime(2021, 1, 1).ToOADate();
            var result = this.converter.Convert(Cell(DateTime.MinValue), oaDate);
            Assert.IsType<DateTime>(result);
            Assert.Equal(new DateTime(2021, 1, 1), result);
        }

        [Fact]
        public void DateTimeOutOfRangeKeepsCurrentWithoutThrowing()
        {
            var keep = new DateTime(2020, 1, 1);
            // 1e30 is outside the valid OADate range; FromOADate would throw ArgumentException.
            var result = this.converter.Convert(Cell(keep), 1e30);
            Assert.Equal(keep, result);
        }

        [Fact]
        public void DateTimePassesThroughDateTime()
        {
            var dateTime = new DateTime(2021, 5, 5);
            Assert.Equal(dateTime, this.converter.Convert(Cell(DateTime.MinValue), dateTime));
        }

        // ---------------- string target ----------------

        [Fact]
        public void StringFromDoubleUsesInvariantCulture()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("nl-NL");
                Assert.Equal("1.5", this.converter.Convert(Cell(string.Empty), 1.5d));
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        [Fact]
        public void StringFromNullIsEmpty()
        {
            Assert.Equal(string.Empty, this.converter.Convert(Cell("x"), null));
        }

        [Fact]
        public void StringPassesThroughString()
        {
            Assert.Equal("hello", this.converter.Convert(Cell("x"), "hello"));
        }

        [Fact]
        public void StringFromBool()
        {
            Assert.Equal("True", this.converter.Convert(Cell(string.Empty), true));
        }

        // ---------------- bool target ----------------

        [Fact]
        public void BoolFromDouble()
        {
            Assert.Equal(true, this.converter.Convert(Cell(false), 1.0));
            Assert.Equal(false, this.converter.Convert(Cell(true), 0.0));
        }

        // ---------------- no type hint ----------------

        [Fact]
        public void NullCurrentPassesThrough()
        {
            Assert.Equal(42d, this.converter.Convert(Cell(null), 42d));
        }

        // ---------------- configurability ----------------

        [Fact]
        public void RoundingModeIsConfigurable()
        {
            var bankers = new DefaultValueConverter(CultureInfo.InvariantCulture, MidpointRounding.ToEven);
            Assert.Equal(2, bankers.Convert(Cell(0), 2.5d));
        }
    }
}
