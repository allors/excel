// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelAddIn.VSTO.Tests
{
    using Allors.Excel;
    using Allors.Excel.Interop;
    using Moq;

    public class WorkbookTests : Allors.Excel.Tests.WorkbookTests
    {
        private readonly InteropApplication application;

        public WorkbookTests() => this.application = new InteropApplication { Visible = true };

        public override void Dispose()
        {
            var workbooks = this.application.Workbooks;
            foreach (InteropWorkbook workbook in workbooks)
            {
                workbook.Close(false);
            }

            this.application.Quit();
        }

        protected override IAddIn NewAddIn()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = new AddIn(this.application, program.Object, ribbon.Object);
            return addIn;
        }

        protected override void AddWorkbook() => this.application.Workbooks.Add();
    }
}
