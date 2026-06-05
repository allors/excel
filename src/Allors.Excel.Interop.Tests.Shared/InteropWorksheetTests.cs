// <copyright file="InteropWorksheetTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop.Tests.Shared
{
    using Xunit;

    /// <summary>
    /// Base for the interop WorksheetTests suites: runs the shared worksheet facts
    /// against a real Excel instance hosted by <see cref="InteropExcelFixture"/>.
    /// </summary>
    public abstract class InteropWorksheetTests : Allors.Excel.Tests.WorksheetTests
    {
        private readonly InteropExcelFixture fixture;

        protected InteropWorksheetTests()
        {
            // Created here, not in a field initializer, so Excel starts after the
            // base constructor has run (field initializers precede base constructors).
            this.fixture = new InteropExcelFixture();
        }

        public override void Dispose() => this.fixture.Dispose();

        protected override IAddIn NewAddIn() => this.fixture.NewAddIn();

        protected override void AddWorkbook() => this.fixture.AddWorkbook();

        // FreezePanes(Range(0, 5)): cell (row 0, column 5) is the top-left of the scrollable
        // region, i.e. freeze 0 rows and 5 columns. A row index of 0 must NOT be treated as
        // the special "freeze top row" case (which would offset to 1 row / 6 columns).
        [Fact]
        public void FreezePanesAtCellWithZeroRowDoesNotOffset()
        {
            var addIn = this.NewAddIn();
            this.AddWorkbook();

            var worksheet = (Worksheet)addIn.Workbooks[0].Worksheets[0];

            worksheet.FreezePanes(new Range(0, 5, columns: 1));

            var window = worksheet.InteropWorksheet.Application.ActiveWindow;
            Assert.Equal(0, window.SplitRow);
            Assert.Equal(5, window.SplitColumn);
        }

        [Fact]
        public void FreezePanesAtInteriorCell()
        {
            var addIn = this.NewAddIn();
            this.AddWorkbook();

            var worksheet = (Worksheet)addIn.Workbooks[0].Worksheets[0];

            worksheet.FreezePanes(new Range(3, 5, columns: 1));

            var window = worksheet.InteropWorksheet.Application.ActiveWindow;
            Assert.Equal(3, window.SplitRow);
            Assert.Equal(5, window.SplitColumn);
        }

        [Fact]
        public void FreezePanesTopRowOnly()
        {
            var addIn = this.NewAddIn();
            this.AddWorkbook();

            var worksheet = (Worksheet)addIn.Workbooks[0].Worksheets[0];

            // (Row 0, Column -1) is the documented "freeze the top row" case.
            worksheet.FreezePanes(new Range(0, -1, columns: 1));

            var window = worksheet.InteropWorksheet.Application.ActiveWindow;
            Assert.Equal(1, window.SplitRow);
            Assert.Equal(0, window.SplitColumn);
        }
    }
}
