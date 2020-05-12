// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Linq;
using Allors.Excel.Embedded;
using Moq;
using Xunit;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Embedded
{
    public class WorksheetTests : IDisposable
    {
        private InteropApplication application;

        public WorksheetTests()
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


        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void InsertRows()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (int i = 0; i < 10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";
            }

            await iWorksheet.Flush().ConfigureAwait(true);

            iWorksheet.InsertRows(3, 1);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            cell = iWorksheet[1, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A1");

            cell = iWorksheet[2, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A2");

            cell = iWorksheet[3, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A3");

            // newly inserted cell has no value
            cell = iWorksheet[4, 0];
            Assert.Null(cell.Value);

            // shifted cell 1 down. The cell that was at row 4 is now in Row 5
            cell = iWorksheet[5, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A4");
        }

        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void DeleteRows()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");
            
            for(int i=0; i<10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";
            }          
            await iWorksheet.Flush().ConfigureAwait(true);

            // Delete rows 3, 4 and 5
            iWorksheet.DeleteRows(2, 3); // Zero-Based!

            await iWorksheet.Flush().ConfigureAwait(true);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            cell = iWorksheet[1, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A1");

            cell = iWorksheet[2, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A5");

            cell = iWorksheet[3, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A6");
        }

        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void InsertColumn()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (int i = 0; i < 10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";

                cell = iWorksheet[i, 1];
                cell.Value = $"Cell B{i}";

                cell = iWorksheet[i, 2];
                cell.Value = $"Cell C{i}";
            }

            await iWorksheet.Flush().ConfigureAwait(true);

            iWorksheet.InsertColumns(0, 1);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            // newly inserted cell has no value
            cell = iWorksheet[0, 1];
            Assert.Null(cell.Value);

            cell = iWorksheet[0, 2];
            Assert.True(Convert.ToString(cell.Value) == "Cell B0");
        }

        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void InsertColumns()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (int i = 0; i < 10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";

                cell = iWorksheet[i, 1];
                cell.Value = $"Cell B{i}";

                cell = iWorksheet[i, 2];
                cell.Value = $"Cell C{i}";
            }

            await iWorksheet.Flush().ConfigureAwait(true);

            iWorksheet.InsertColumns(0, 2);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            // newly inserted cell has no value
            cell = iWorksheet[0, 1];
            Assert.Null(cell.Value);

            cell = iWorksheet[0, 2];
            Assert.Null(cell.Value);

            cell = iWorksheet[0, 3];
            Assert.True(Convert.ToString(cell.Value) == "Cell B0");
        }

        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void DeleteColumn()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");                     

            for (int i = 0; i < 10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";

                cell = iWorksheet[i, 1];
                cell.Value = $"Cell B{i}";

                cell = iWorksheet[i, 2];
                cell.Value = $"Cell C{i}";

                cell = iWorksheet[i, 3];
                cell.Value = $"Cell D{i}";
            }

            await iWorksheet.Flush().ConfigureAwait(true);

            // Delete column B
            iWorksheet.DeleteColumns(1, 1);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            cell = iWorksheet[1, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A1");


            // B is Gone!

            cell = iWorksheet[0, 1];
            Assert.True(Convert.ToString(cell.Value) == "Cell C0");

            cell = iWorksheet[1, 1];
            Assert.True(Convert.ToString(cell.Value) == "Cell C1");

            cell = iWorksheet[0, 2];
            Assert.True(Convert.ToString(cell.Value) == "Cell D0");

            cell = iWorksheet[1, 2];
            Assert.True(Convert.ToString(cell.Value) == "Cell D1");
        }

        [Fact(Skip = "Azure vmimage has no office installed")]
        //[Fact]
        public async void DeleteColumns()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (int i = 0; i < 10; i++)
            {
                cell = iWorksheet[i, 0];
                cell.Value = $"Cell A{i}";

                cell = iWorksheet[i, 1];
                cell.Value = $"Cell B{i}";

                cell = iWorksheet[i, 2];
                cell.Value = $"Cell C{i}";

                cell = iWorksheet[i, 3];
                cell.Value = $"Cell D{i}";
            }

            await iWorksheet.Flush().ConfigureAwait(true);

            // Delete column B
            iWorksheet.DeleteColumns(1, 2);

            cell = iWorksheet[0, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A0");

            cell = iWorksheet[1, 0];
            Assert.True(Convert.ToString(cell.Value) == "Cell A1");


            // B is Gone!
            // C is Gone!


            cell = iWorksheet[0, 1];
            Assert.True(Convert.ToString(cell.Value) == "Cell D0");

            cell = iWorksheet[1, 1];
            Assert.True(Convert.ToString(cell.Value) == "Cell D1");
        }
    }
}
