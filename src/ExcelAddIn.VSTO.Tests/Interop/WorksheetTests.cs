// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelAddIn.VSTO.Tests
{
    public class WorksheetTests : Allors.Excel.Tests.Interop.WorksheetTests
    {
        public WorksheetTests()
        {
            this.application = new InteropApplication { Visible = true };
        }

        public override void Dispose()
        {
            var workbooks = this.application.Workbooks;
            foreach (InteropWorkbook workbook in workbooks)
            {
                workbook.Close(false);
            }


            this.application.Quit();
        }
    }
}
