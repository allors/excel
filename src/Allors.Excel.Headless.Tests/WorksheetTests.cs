// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>


namespace Allors.Excel.Headless.Tests
{
    using Allors.Excel;
    using Allors.Excel.Headless;
    using Allors.Excel.Tests;
    using Moq;

    public class WorksheetTests : Excel.Tests.WorksheetTests
    {
        private AddIn addIn;

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
            this.addIn = new AddIn(new TestProgram(), ribbon);

            return this.addIn;
        }

        protected override void AddWorkbook() => this.addIn.AddWorkbook();
    }
}
