// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>


namespace Allors.Excel.Headless.Tests
{
    using System.Threading.Tasks;
    using Allors.Excel;
    using Allors.Excel.Headless;
    using Allors.Excel.Tests;
    using Moq;
    using Xunit;

    public class WorksheetTests : Excel.Tests.WorksheetTests
    {
        private AddIn addIn;

        [Fact]
        public async Task GetRectangleForSingleDimensionNamedRange()
        {
            this.NewAddIn();
            this.AddWorkbook();

            var worksheet = await this.addIn.Workbooks[0].AddWorksheet();

            // A columns-only range (Rows == null) must not throw: the missing dimension is a
            // single row (Range.EffectiveRows), so the rectangle is 1 row high and 3 wide.
            worksheet.SetNamedRange("COLUMNS.ONLY", new Allors.Excel.Range(4, 5, columns: 3, worksheet: worksheet));
            var columnsOnly = worksheet.GetRectangle("COLUMNS.ONLY");
            Assert.Equal(3, columnsOnly.Width);
            Assert.Equal(1, columnsOnly.Height);

            // ... and a rows-only range is 1 column wide and 3 rows high.
            worksheet.SetNamedRange("ROWS.ONLY", new Allors.Excel.Range(4, 5, rows: 3, worksheet: worksheet));
            var rowsOnly = worksheet.GetRectangle("ROWS.ONLY");
            Assert.Equal(1, rowsOnly.Width);
            Assert.Equal(3, rowsOnly.Height);
        }

        public override void Dispose()
        {
            var workbooks = this.addIn.Workbooks;
            foreach (var workbook in workbooks)
            {
                workbook.Close(false);
            }


            this.addIn = null;
        }

        protected override IAddIn NewAddIn()
        {
            if (this.addIn != null)
            {
                throw new System.Exception("Only one AddIn allowed");
            }

            var ribbon = new Mock<IRibbon>().Object;
            this.addIn = AddIn.CreateAsync(new TestProgram(), ribbon).GetAwaiter().GetResult();

            return this.addIn;
        }

        protected override void AddWorkbook() => this.addIn.AddWorkbook().GetAwaiter().GetResult();
    }
}
