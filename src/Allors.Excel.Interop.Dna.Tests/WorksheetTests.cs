// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using Allors.Excel;
using Moq;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Interop.Dna.Tests
{
    using Allors.Excel.Interop;

    public class WorksheetTests : Allors.Excel.Tests.WorksheetTests
    {
        private readonly InteropApplication application = new InteropApplication { Visible = true };

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
