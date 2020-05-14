// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Diagnostics;
using System.Linq;
using Allors.Excel.Embedded;
using Moq;
using Xunit;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Embedded
{   
    public class WorkbookTests : InteropTest
    {
        [Fact(Skip=skipReason)]
        public async void OnNew()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            program.Verify(mock => mock.OnNew(It.IsAny<IWorkbook>()), Times.Once());

            await System.Threading.Tasks.Task.CompletedTask;
        }

        [Fact(Skip = skipReason)]
        public void SetNamedRangeWorkbook() 
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");          

            Range range = new Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));

            namedRanges = iWorksheet.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }

        [Fact(Skip = skipReason)]
        public void SetNamedRangeWorksheet()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            Range range = new Range(4, 5, 1, 10, iWorksheet);

            iWorksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = iWorksheet.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "'2'!MY.NAMEDRANGE"));

            namedRanges = workbook.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }       

        [Fact(Skip = skipReason)]
        public void UpdateNamedRangeWorkbook()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            Range range = new Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));

            range = new Range(8, 10, 2, 4, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRange = workbook.GetNamedRanges().First(v => string.Equals(v.Name, "MY.NAMEDRANGE"));

            Assert.Equal(8, namedRange.Row);
            Assert.Equal(10, namedRange.Column);
            Assert.Equal(2, namedRange.Rows);
            Assert.Equal(4, namedRange.Columns);
        }

        

        [Fact(Skip = skipReason)]
        public void UpdateNamedRangeWorksheet()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            Range range = new Range(4, 5, 1, 10, iWorksheet);

            iWorksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = iWorksheet.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "'2'!MY.NAMEDRANGE"));

            range = new Range(8, 10, 2, 4, iWorksheet);
            iWorksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRange = iWorksheet.GetNamedRanges().First(v => string.Equals(v.Name, "'2'!MY.NAMEDRANGE"));

            Assert.Equal(8, namedRange.Row);
            Assert.Equal(10, namedRange.Column);
            Assert.Equal(2, namedRange.Rows);
            Assert.Equal(4, namedRange.Columns);
        }

        [Fact(Skip = skipReason)]
        public void GetNamedRangeWorkbookForWorksheet()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            Range range = new Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges("2");

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }
    }
}
