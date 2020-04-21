// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using Allors.Excel.Embedded;
using Moq;
using Xunit;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Embedded
{
    public class WorkbookTests : IDisposable
    {
        private InteropApplication application;

        public WorkbookTests()
        {
            this.application = new InteropApplication { Visible = true };
        }

        public void Dispose()
        {
            foreach (InteropWorkbook workbook in this.application.Workbooks)
            {
                workbook.Close(false);
            }

            this.application.Quit();
        }

        [Fact(Skip ="Azure vmimage has no office installed")]
        public async void OnNew()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            program.Verify(mock => mock.OnNew(It.IsAny<IWorkbook>()), Times.Once());
        }
    }
}
