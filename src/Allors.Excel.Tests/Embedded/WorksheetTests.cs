// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
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
        private DirectoryInfo tempDirectory;

        public WorksheetTests()
        {            
            var dir = Path.Combine(Path.GetTempPath(), nameof(WorksheetTests));
            tempDirectory = new DirectoryInfo(dir);

            if (tempDirectory.Exists)
            {
                foreach (var file in tempDirectory.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {

                    }
                }
            }           
        }        

        [Fact(Skip = skipReason)]
        public void SaveAsPDFWithNullThrowsException()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            Assert.Throws<ArgumentNullException>(() => sheet2.SaveAsPDF(null));           
        }

        [Fact(Skip = skipReason)]
        public void SetPrintArea()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            sheet2.SetPrintArea(new Range(5, 1, 10, 5, sheet2));

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            // PrintArea is set, but we do not want to use it. Prints the entire sheet
            sheet2.SaveAsPDF(file, true, false, ignorePrintAreas: true);

            // use the printArea
            sheet2.SaveAsPDF(file, true, false, false);

            Assert.True(new FileInfo(file.FullName).Exists);

            // Set PrintArea to entire sheet
            sheet2.SetPrintArea(null);
            sheet2.SaveAsPDF(file, true, false, false);
            Assert.True(new FileInfo(file.FullName).Exists);

            sheet2.SaveAsPDF(file, true, false, ignorePrintAreas: true);
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact(Skip = skipReason)]
        public void SaveAsPDF()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));
         
            sheet2.SaveAsPDF(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact(Skip = skipReason)]
        public void SaveAsPDFThrowsComExceptionWhenEmpty()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");

            // There is nothing to print => exception
            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet1)}.pdf"));
            Assert.Throws<COMException>(() => sheet1.SaveAsPDF(file));
        }


        [Fact(Skip = skipReason)]
        public async void SaveAsPDFThrowsExceptionWhenFileExists()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];    
          
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            // First save with overwriteExistingFile
            sheet2.SaveAsPDF(file, true);

            var lastWriteTime = new FileInfo(file.FullName).LastWriteTimeUtc;

            // Second save with the same name will throw an exception
            // File exist and should not be overwritten.
            Assert.Throws<IOException>( () =>  sheet2.SaveAsPDF(file, false));

            Thread.Sleep(1000);

            // Third save will overwrite existingFile
            sheet2.SaveAsPDF(file, true);

            Assert.True(new FileInfo(file.FullName).LastWriteTimeUtc > lastWriteTime);          

        }

        [Fact(Skip = skipReason)]
        public void SaveAsXPS()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.xps"));

            sheet2.SaveAsXPS(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact(Skip = skipReason)]
        public void SaveAsXPSSetsExtensiontoXPS()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.AAA"));

            sheet2.SaveAsXPS(file);

            Assert.False(new FileInfo(file.FullName).Exists);

            file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.xps"));
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact(Skip = skipReason)]
        public void SaveAsPDFSetsExtensiontoXPS()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.AAA"));

            sheet2.SaveAsPDF(file);

            Assert.False(new FileInfo(file.FullName).Exists);

            file = new FileInfo(Path.Combine(tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }



        [Fact(Skip = skipReason)]
        public async void FreezePanes()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var sheet1 = workbook.Worksheets.First();
            var sheet2 = workbook.Worksheets.Last();

            // Cell B2 is topLeft
            var range = new Range(1, 1, 0, 0, sheet1);
            sheet1.FreezePanes(range);

            Assert.True(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            sheet1.UnfreezePanes();

            Assert.False(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            // First Row
            range = new Range(0, -1, 1, 1);
            sheet1.FreezePanes(range);

            Assert.True(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            // First Column
            range = new Range(-1, 0, 1, 1);
            sheet1.FreezePanes(range);

            Assert.True(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            // Entire Row 5
            range = new Range(5, -1, 0, 0, sheet1);
            sheet1.FreezePanes(range);

            Assert.True(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            sheet1.UnfreezePanes();

            Assert.False(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            // Entire Column 5
            range = new Range(-1, 5, 0, 0, sheet1);
            sheet1.FreezePanes(range);

            Assert.True(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            sheet1.UnfreezePanes();

            Assert.False(sheet1.HasFreezePanes);
            Assert.False(sheet2.HasFreezePanes);

            for (int row = 0; row < 15; row++)
            {
                range = new Range(row, -1, 0, 0, sheet1);
                sheet1.FreezePanes(range);

                Thread.Sleep(200);
            }

            for (int column = 0; column < 15; column++)
            {
                range = new Range(-1, column, 0, 0, sheet1);
                sheet1.FreezePanes(range);

                Thread.Sleep(200);
            }
        }

        [Fact(Skip = skipReason)]
        public async void AddWorksheetsBeforeAndAfter()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var sheet1 = workbook.Worksheets.First();
            var sheet2 = workbook.Worksheets.Last();
            Assert.Equal("2", sheet2.Name);

            // 1  | 2

            // Add before "1"
            var worksheet = (Worksheet)workbook.AddWorksheet(null, sheet1, null);
            worksheet.Name = "#3";

            // Expected order => #3 | 1  | 2
            Assert.Equal(3, workbook.Worksheets.Length);
            Assert.Equal(3, ((Workbook)workbook).WorksheetsByIndex.Length);
            Assert.Equal(1, worksheet.Index);

            var worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;
            Assert.Equal("#3", worksheetsByIndex[0].Name);
            Assert.Equal("1", worksheetsByIndex[1].Name);
            Assert.Equal("2", worksheetsByIndex[2].Name);

            // Add before "2"
            worksheet = (Worksheet)workbook.AddWorksheet(null, sheet2, null);
            worksheet.Name = "#4";

            // Expected order => #3 | 1  | #4 | 2
            worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;
            Assert.Equal(4, workbook.Worksheets.Length);
            Assert.Equal(4, ((Workbook)workbook).WorksheetsByIndex.Length);
            Assert.Equal(3, worksheet.Index);
            Assert.Equal("#3", worksheetsByIndex[0].Name);
            Assert.Equal("1", worksheetsByIndex[1].Name);
            Assert.Equal("#4", worksheetsByIndex[2].Name);
            Assert.Equal("2", worksheetsByIndex[3].Name);

            // Add after "1"
            worksheet = (Worksheet)workbook.AddWorksheet(null, null, sheet1);
            worksheet.Name = "#5";

            // Expected order => #3 | 1 | #5 | #4 | 2
            worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;
            Assert.Equal(5, workbook.Worksheets.Length);
            Assert.Equal(5, ((Workbook)workbook).WorksheetsByIndex.Length);
            Assert.Equal(3, worksheet.Index);
            Assert.Equal("#3", worksheetsByIndex[0].Name);
            Assert.Equal("1", worksheetsByIndex[1].Name);
            Assert.Equal("#5", worksheetsByIndex[2].Name);
            Assert.Equal("#4", worksheetsByIndex[3].Name);
            Assert.Equal("2", worksheetsByIndex[4].Name);

            // Add after "2"
            worksheet = (Worksheet)workbook.AddWorksheet(null, null, sheet2);
            worksheet.Name = "#6";

            // Expected order => #3 | 1 | #5 | #4 | 2 | #6
            worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;
            Assert.Equal(6, workbook.Worksheets.Length);
            Assert.Equal(6, ((Workbook)workbook).WorksheetsByIndex.Length);
            Assert.Equal(6, worksheet.Index);
            Assert.Equal("#3", worksheetsByIndex[0].Name);
            Assert.Equal("1", worksheetsByIndex[1].Name);
            Assert.Equal("#5", worksheetsByIndex[2].Name);
            Assert.Equal("#4", worksheetsByIndex[3].Name);
            Assert.Equal("2", worksheetsByIndex[4].Name);
            Assert.Equal("#6", worksheetsByIndex[5].Name);
        }

        [Fact(Skip = skipReason)]
        public async void AddWorksheetsByIndex()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var last = workbook.Worksheets.Last();
            Assert.Equal("2", last.Name);
            
            // At before index 1
            var worksheet = (Worksheet)workbook.AddWorksheet(1);
            var worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;
            Assert.Equal(3, workbook.Worksheets.Length);
            Assert.Equal(3, worksheetsByIndex.Length);
            Assert.Equal(1, worksheet.Index);
            
            // Expected order => Sheet3 | 1  | 2

            Assert.Equal("Sheet3", worksheetsByIndex[0].Name);
            Assert.Equal("1", worksheetsByIndex[1].Name);
            Assert.Equal("2", worksheetsByIndex[2].Name);

            // Add before index 2
            worksheet = (Worksheet)workbook.AddWorksheet(2);
            worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;

            Assert.Equal(4, workbook.Worksheets.Length);
            Assert.Equal(4, worksheetsByIndex.Length);
            Assert.Equal(2, worksheet.Index);

            // Expected order => Sheet3 | Sheet4 | 1 | 2  !! Order is determined dynamically, so it changes after the first AddWorksheet()
            Assert.Equal("Sheet3", worksheetsByIndex[0].Name);
            Assert.Equal("Sheet4", worksheetsByIndex[1].Name);
            Assert.Equal("1", worksheetsByIndex[2].Name);
            Assert.Equal("2", worksheetsByIndex[3].Name);

            // 0 Adds Before the Active Sheet (being the last added sheet here)
            worksheet = (Worksheet)workbook.AddWorksheet(0);
            worksheetsByIndex = ((Workbook)workbook).WorksheetsByIndex;

            Assert.Equal(5, workbook.Worksheets.Length);
            Assert.Equal(5, worksheetsByIndex.Length);
            Assert.Equal(2, worksheet.Index);

            // Expected order => Sheet3 | Sheet5 | Sheet4 | 1 | 2  !! Order is determined dynamically, so it changes after the first AddWorksheet()
            Assert.Equal("Sheet3", worksheetsByIndex[0].Name);
            Assert.Equal("Sheet5", worksheetsByIndex[1].Name);
            Assert.Equal("Sheet4", worksheetsByIndex[2].Name);
            Assert.Equal("1", worksheetsByIndex[3].Name);
            Assert.Equal("2", worksheetsByIndex[4].Name);
        }

        [Fact(Skip = skipReason)]
        public async void CellTagContainsCustomObject()
        {
            this.ExpectedContextTags = new List<ContextTag>();

            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var worksheet = (Worksheet) workbook.AddWorksheet(null, null, workbook.Worksheets.Last());

            worksheet.CellsChanged += Worksheet_CellsChanged;

            var tag1 = new ContextTag() { Context = "Cell00" };
            var tag2 = new ContextTag() { Context = "Cell01" };
            this.ExpectedContextTags.Add(tag1);
            this.ExpectedContextTags.Add(tag2);

            var cell00 = worksheet[0, 0];
            cell00.Tag = tag1;

            Assert.NotEmpty(this.ExpectedContextTags);

            // Change the cell will trigger the Change Event
            this.ExpectedContextTag = tag1;
            worksheet.InteropWorksheet.Cells[1, 1] = "i am cell00";                      
          
            var cell01 = worksheet[0, 1];
            cell01.Tag = tag2;

            // Change the cell will trigger the Change Event
            this.ExpectedContextTag = tag2;
            worksheet.InteropWorksheet.Cells[1, 2] = "i am cell01";          

            Assert.Empty(this.ExpectedContextTags);
        }

        private ContextTag ExpectedContextTag;
        private List<ContextTag> ExpectedContextTags;

        private class ContextTag
        {
            public string Context { get; set; }
        }

        private void Worksheet_CellsChanged(object sender, CellChangedEvent e)
        {
            var tag = (ContextTag) e.Cells[0].Tag;

            Assert.Equal(this.ExpectedContextTag, tag);

            this.ExpectedContextTags.Remove(tag);

        }

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

        [Fact(Skip = skipReason)]
        public void GetRectangle()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[1];

            sheet1.SetNamedRange("image", new Range(1, 1, 2, 2, sheet1));

            var rectangle = sheet1.GetRectangle("image");

            Assert.True(rectangle.Left > 0);
           
        }

        [Fact(Skip = skipReason)]
        public void NewSheetIsActive()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.AddWorksheet(0);                      

            Assert.True(sheet1.IsActive);

        }
    }
}
