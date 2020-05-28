// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Linq;
using System.Reflection;
using Allors.Excel.Embedded;
using Application;
using Moq;
using Xunit;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Embedded
{
    public class WorksheetTests : InteropTest
    {

        [Fact(Skip = skipReason)]
        public async void IsVisible()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();           

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var worksheet = workbook.Worksheets.First();
            Assert.True(worksheet.IsVisible);

            worksheet.IsVisible = false;

            Assert.False(worksheet.IsVisible);

            worksheet.IsVisible = true;

            Assert.True(worksheet.IsVisible);
        }


        [Fact(Skip = skipReason)]
        public async void AddWorkbook()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

        }

        [Fact(Skip = skipReason)]
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
            Assert.True(cell.ValueAsString == "Cell A0");

            cell = iWorksheet[1, 0];
            Assert.True(cell.ValueAsString == "Cell A1");

            cell = iWorksheet[2, 0];
            Assert.True(cell.ValueAsString == "Cell A2");

            cell = iWorksheet[3, 0];
            Assert.True(cell.ValueAsString == "Cell A3");

            // newly inserted cell has no value
            cell = iWorksheet[4, 0];
            Assert.Null(cell.Value);

            // shifted cell 1 down. The cell that was at row 4 is now in Row 5
            cell = iWorksheet[5, 0];
            Assert.True(cell.ValueAsString == "Cell A4");
        }

        [Fact(Skip = skipReason)]
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

        [Fact(Skip = skipReason)]
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

        [Fact(Skip = skipReason)]
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

        [Fact(Skip = skipReason)]
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

        [Fact(Skip = skipReason)]
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

        [Fact(Skip = skipReason)]
        public async void SetIsActiveWorksheet()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
            ICell cell = null;

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[0];
            var sheet2 = workbook.Worksheets[1];

            sheet1.IsActive = true;
            Assert.True(sheet1.IsActive);
            Assert.False(sheet2.IsActive);

            sheet2.IsActive = true;
            Assert.True(sheet2.IsActive);
            Assert.False(sheet1.IsActive);
        }

        [Fact(Skip = skipReason)]
        public void GetRange()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();
           
            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[0];

            var range = sheet1.GetRange(null);
            Assert.Null(range);

            range = sheet1.GetRange("");
            Assert.Null(range);

            range = sheet1.GetRange("  ");
            Assert.Null(range);

            range = sheet1.GetRange("-1");
            Assert.Null(range);

            range = sheet1.GetRange("BLABLA");
            Assert.Null(range);

            range = sheet1.GetRange("A1");

            Assert.Equal(0, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(1, range.Columns);

            range = sheet1.GetRange("A1:C5");

            Assert.Equal(0, range.Row);
            Assert.Equal(5, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(3, range.Columns);

            range = sheet1.GetRange("A1", "C5");

            Assert.Equal(0, range.Row);
            Assert.Equal(5, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(3, range.Columns);

            range = sheet1.GetRange("A:A");

            Assert.Equal(0, range.Row);
            Assert.Equal(1048576, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(1, range.Columns);

            range = sheet1.GetRange("A:C");

            Assert.Equal(0, range.Row);
            Assert.Equal(1048576, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(3, range.Columns);

            range = sheet1.GetRange("A:A", "C:C");

            Assert.Equal(0, range.Row);
            Assert.Equal(1048576, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(3, range.Columns);

            range = sheet1.GetRange("C3", "D4");

            Assert.Equal(2, range.Row);
            Assert.Equal(2, range.Rows);

            Assert.Equal(2, range.Column);
            Assert.Equal(2, range.Columns);

        }

        [Fact(Skip = skipReason)]
        public async void GetUsedRange()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[1];

            var range = sheet1.GetUsedRange();
            Assert.Equal(0, range.Row);
            Assert.Equal(50, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(15, range.Columns);

            // for Column
            range = sheet1.GetUsedRange("B");
            Assert.Equal(0, range.Row);
            Assert.Equal(50, range.Rows);

            Assert.Equal(1, range.Column);
            Assert.Equal(1, range.Columns);

            range = sheet1.GetUsedRange("L");
            Assert.Equal(2, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(11, range.Column);
            Assert.Equal(1, range.Columns);

            // for Row
            range = sheet1.GetUsedRange(0);
            Assert.Equal(0, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(10, range.Columns);

            range = sheet1.GetUsedRange(2);
            Assert.Equal(2, range.Row);
            Assert.Equal(1, range.Rows);
            Assert.Equal(0, range.Column);
            Assert.Equal(12, range.Columns);

            range = sheet1.GetUsedRange(3);
            Assert.Equal(3, range.Row);
            Assert.Equal(1, range.Rows);
            Assert.Equal(0, range.Column);
            Assert.Equal(15, range.Columns);

            // Zero based row index
            sheet1[50, 2].Value = "x";
            sheet1[50, 3].Value = "y";
            sheet1[50, 4].Value = "z";

            await sheet1.Flush().ConfigureAwait(false);

            range = sheet1.GetUsedRange(50);
            Assert.Equal(50, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(2, range.Column);
            Assert.Equal(3, range.Columns);
        }

        [Fact(Skip = skipReason)]
        public async void GetUsedRangeColumn()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[1];

            var range = sheet1.GetUsedRange();
            Assert.Equal(0, range.Row);
            Assert.Equal(50, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(15, range.Columns);

            // for Column
            range = sheet1.GetUsedRange("B");
            Assert.Equal(0, range.Row);
            Assert.Equal(50, range.Rows);

            Assert.Equal(1, range.Column);
            Assert.Equal(1, range.Columns);

            range = sheet1.GetUsedRange("L");
            Assert.Equal(2, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(11, range.Column);
            Assert.Equal(1, range.Columns);                      
          
            sheet1[50, 30].Value = "x";
            sheet1[51, 30].Value = "y";
            sheet1[52, 30].Value = "z";

            await sheet1.Flush().ConfigureAwait(false);

            var columnName = Worksheet.ExcelColumnFromNumber(31);

            range = sheet1.GetUsedRange(columnName);
            Assert.Equal(50, range.Row);
            Assert.Equal(3, range.Rows);

            Assert.Equal(30, range.Column);
            Assert.Equal(1, range.Columns);

            // Blank line is still counted as a row
            sheet1[54, 30].Value = "aa";

            await sheet1.Flush().ConfigureAwait(false);

            range = sheet1.GetUsedRange(columnName);
            Assert.Equal(50, range.Row);
            Assert.Equal(5, range.Rows);

            Assert.Equal(30, range.Column);
            Assert.Equal(1, range.Columns);
        }

        [Fact(Skip = skipReason)]
        public async void GetUsedRangeRow()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[1];                       

            // for Row
            var range = sheet1.GetUsedRange(0);
            Assert.Equal(0, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(0, range.Column);
            Assert.Equal(10, range.Columns);

            range = sheet1.GetUsedRange(2);
            Assert.Equal(2, range.Row);
            Assert.Equal(1, range.Rows);
            Assert.Equal(0, range.Column);
            Assert.Equal(12, range.Columns);

            range = sheet1.GetUsedRange(3);
            Assert.Equal(3, range.Row);
            Assert.Equal(1, range.Rows);
            Assert.Equal(0, range.Column);
            Assert.Equal(15, range.Columns);

            // Zero based row index
            sheet1[50, 2].Value = "x";
            sheet1[50, 3].Value = "y";
            sheet1[50, 4].Value = "z";

            await sheet1.Flush().ConfigureAwait(false);

            range = sheet1.GetUsedRange(50);
            Assert.Equal(50, range.Row);
            Assert.Equal(1, range.Rows);

            Assert.Equal(2, range.Column);
            Assert.Equal(3, range.Columns);
        }
    }
}
