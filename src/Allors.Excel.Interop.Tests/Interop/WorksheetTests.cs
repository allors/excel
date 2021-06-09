// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Tests.Interop
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Allors.Excel.Interop;
    using Moq;
    using Xunit;
    using Range = Excel.Range;

    public abstract class WorksheetTests : InteropTest
    {
        private DirectoryInfo tempDirectory;

        protected WorksheetTests()
        {
            var dir = Path.Combine(Path.GetTempPath(), nameof(WorksheetTests));
            this.tempDirectory = new DirectoryInfo(dir);

            if (this.tempDirectory.Exists)
            {
                foreach (var file in this.tempDirectory.GetFiles())
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

        [Fact]
        public void SheetHasIndex()
        {
            var program = new Mock<IProgram>();
            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");
            Assert.Equal(1, sheet1.Index);

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");
            Assert.Equal(2, sheet2.Index);
        }

        [Fact]
        public void NewSheetAtActiveSheetHasIndex()
        {
            var program = new Mock<IProgram>();
            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");
            Assert.Equal(1, sheet1.Index);

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");
            Assert.Equal(2, sheet2.Index);

            // Add sheet to the left of the activeSheet (== sheet 2)
            var sheet = workbook.AddWorksheet(0);
            Assert.Equal(1, sheet1.Index);
            Assert.Equal(2, sheet.Index);
            Assert.Equal(3, sheet2.Index);

        }

        [Fact]
        public void NewSheetBeforeSheetHasIndex()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");
            Assert.Equal(1, sheet1.Index);

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");
            Assert.Equal(2, sheet2.Index);

            // Add sheet to the left of the sheet 2
            var sheet = workbook.AddWorksheet(null, sheet2);
            Assert.Equal(1, sheet1.Index);
            Assert.Equal(2, sheet.Index);
            Assert.Equal(3, sheet2.Index);

        }

        [Fact]
        public void NewSheetAtIndexHasIndex()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");
            Assert.Equal(1, sheet1.Index);

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");
            Assert.Equal(2, sheet2.Index);

            // Add sheet to right of index 99 (index > sheets.Count will add sheet after last index)
            var sheet = workbook.AddWorksheet(99);
            Assert.Equal(1, sheet1.Index);
            Assert.Equal(2, sheet2.Index);
            Assert.Equal(3, sheet.Index);

            // Add sheet at index 2 will add before sheet 2 (sheet with index 2)
            // So it will get the index we want it to have.
            sheet = workbook.AddWorksheet(2);
            Assert.Equal(1, sheet1.Index);
            Assert.Equal(3, sheet2.Index);
            Assert.Equal(2, sheet.Index);

        }

        [Fact]
        public void NewSheetAfterSheetHasIndex()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets.Single(v => v.Name == "1");
            Assert.Equal(1, sheet1.Index);

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");
            Assert.Equal(2, sheet2.Index);

            // Add sheet to the right of the sheet 1
            var sheet = workbook.AddWorksheet(null, null, sheet1);
            Assert.Equal(1, sheet1.Index);
            Assert.Equal(2, sheet.Index);
            Assert.Equal(3, sheet2.Index);

        }


        [Fact]
        public void ShowInputMessage()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var cell = sheet2[5, 5];

            sheet2.SetInputMessage(cell, "this is some help message", "title", showInputMessage: false);

            sheet2.SetInputMessage(cell, "this is some help message", "title", showInputMessage: true);

            cell = sheet2[1, 5];
            sheet2.SetInputMessage(cell, "this is some other help message");
            sheet2.HideInputMessage(cell);
            sheet2.HideInputMessage(cell, clearInputMessage: true);
        }

        //[Fact]
        //public void SetCustomProperties()
        //{
        //    var program = new Mock<IProgram>();

        //    var addIn = new AddIn(this.application, program.Object);

        //    this.application.Workbooks.Add();

        //    var workbook = addIn.Workbooks[0];

        //    // Sheet with content
        //    var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

        //    var expectedDate = DateTime.Now;


        //    var dict = new OldCustomProperties();
        //    dict.Add("Showcase.IsInvoiceSheet", false);
        //    dict.Add("Showcase.IsSheet2", true);
        //    dict.Add("Showcase.Sheet2.Date", expectedDate);
        //    dict.Add("Showcase.Sheet2.Decimal", 123.45M);

        //    var nullableDecimal = new decimal?(123.45M);
        //    dict.Add("Showcase.Sheet2.NullableDecimal", nullableDecimal);

        //    dict.Add("Showcase.Sheet2.Int", 12);

        //    var nullableInt = new int?(12);
        //    dict.Add("Showcase.Sheet2.NullableInt", nullableInt);

        //    dict.Add("Company.Name", "Zonsoft.be");
        //    dict.Add("Company.Street", "Uikhoverstraat 158");

        //    // Duplicates will be overwritten
        //    dict.Add("Company.City", "3631 Maasmechelen");
        //    dict.Add("Company.City", "3631 Uikhoven");

        //    dict.Add("Company.Country", "BE België");

        //    dict.Add("Showcase.Sheet2.Null", null);

        //    sheet2.SetCustomProperties(dict);

        //    var customProperties = sheet2.GetCustomProperties();

        //    Assert.Equal(dict.Count, customProperties.Count);

        //    Assert.False(customProperties.Get<bool>("Showcase.IsInvoiceSheet"));
        //    Assert.True(customProperties.Get<bool>("Showcase.IsSheet2"));

        //    // fractions of MS are not preserved!
        //    Assert.Equal(expectedDate.Date, customProperties.Get<DateTime>("Showcase.Sheet2.Date").Date);

        //    Assert.Equal(12, customProperties.Get<int>("Showcase.Sheet2.Int"));
        //    Assert.Equal(12, customProperties.Get<int?>("Showcase.Sheet2.NullableInt"));
        //    Assert.Null(customProperties.Get<int?>("Showcase.Sheet2.Null"));

        //    Assert.Equal(123.45M, customProperties.Get<decimal>("Showcase.Sheet2.Decimal"));
        //    Assert.Equal(123.45M, customProperties.Get<decimal>("Showcase.Sheet2.NullableDecimal"));

        //    Assert.Equal("Zonsoft.be", customProperties.Get<string>("Company.Name"));
        //    Assert.Equal("BE België", customProperties.Get<string>("Company.Country"));
        //    Assert.Equal("3631 Uikhoven", customProperties.Get<string>("Company.City"));

        //}

        [Fact]
        public void SaveAsPDFWithNullThrowsException()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            Assert.Throws<ArgumentNullException>(() => sheet2.SaveAsPDF(null));
        }

        [Fact]
        public void SetPrintArea()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            sheet2.SetPrintArea(new Range(5, 1, 10, 5, sheet2));

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            // PrintArea is set, but we do not want to use it. Prints the entire sheet
            sheet2.SaveAsPDF(file, true);

            // use the printArea
            sheet2.SaveAsPDF(file, true, false, false);

            Assert.True(new FileInfo(file.FullName).Exists);

            // Set PrintArea to entire sheet
            sheet2.SetPrintArea();
            sheet2.SaveAsPDF(file, true, false, false);
            Assert.True(new FileInfo(file.FullName).Exists);

            sheet2.SaveAsPDF(file, true);
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void SaveAsPDFWithHeader()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            sheet2.SetPageSetup(new PageSetup
            {
                Orientation = 1, // Portrait https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpageorientation?view=excel-pia
                PaperSize = 11, // A5 => https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpapersize?view=excel-pia
                Header = new PageHeaderFooter
                {
                    Margin = 10.0,
                    Left = "&\"Arial\"&B&12LeftHeader&B",
                    Right = "&P of &N"
                },
                Footer = new PageHeaderFooter
                {
                    Margin = 60.0,
                    Center = "Copy presented to Walter Hesius",
                }
            });

            sheet2.SaveAsPDF(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void PageSetupOrientationDefaultsToPortrait()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            sheet2.SetPageSetup(new PageSetup
            {
                // No specific settings for PaperSize and Orientation
                Header = new PageHeaderFooter
                {
                    Margin = 10.0,
                    Left = "&\"Arial\"&B&12LeftHeader&B",
                    Right = "&P of &N"
                },
            });

            sheet2.SaveAsPDF(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void SaveAsPDF()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            sheet2.SaveAsPDF(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void SaveAsPDFThrowsComExceptionWhenEmpty()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var newSheet = workbook.AddWorksheet();

            // There is nothing to print => exception
            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(newSheet)}.pdf"));
            Assert.Throws<COMException>(() => newSheet.SaveAsPDF(file));
        }


        [Fact]
        public async void SaveAsPDFThrowsExceptionWhenFileExists()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));

            // First save with overwriteExistingFile
            sheet2.SaveAsPDF(file, true);

            var lastWriteTime = new FileInfo(file.FullName).LastWriteTimeUtc;

            // Second save with the same name will throw an exception
            // File exist and should not be overwritten.
            Assert.Throws<IOException>(() => sheet2.SaveAsPDF(file));

            Thread.Sleep(1000);

            // Third save will overwrite existingFile
            sheet2.SaveAsPDF(file, true);

            Assert.True(new FileInfo(file.FullName).LastWriteTimeUtc > lastWriteTime);

        }

        [Fact]
        public void SaveAsXPS()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.xps"));

            sheet2.SaveAsXPS(file);

            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void SaveAsXPSSetsExtensiontoXPS()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.AAA"));

            sheet2.SaveAsXPS(file);

            Assert.False(new FileInfo(file.FullName).Exists);

            file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.xps"));
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }

        [Fact]
        public void SaveAsPDFSetsExtensiontoXPS()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            // Sheet with content
            var sheet2 = workbook.Worksheets.Single(v => v.Name == "2");

            var file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.AAA"));

            sheet2.SaveAsPDF(file);

            Assert.False(new FileInfo(file.FullName).Exists);

            file = new FileInfo(Path.Combine(this.tempDirectory.FullName, $"{nameof(sheet2)}.pdf"));
            Assert.True(new FileInfo(file.FullName).Exists);

            //sheet2.SaveAsPDF(file, true, true);
        }



        [Fact]
        public async void FreezePanes()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

            for (var row = 0; row < 15; row++)
            {
                range = new Range(row, -1, 0, 0, sheet1);
                sheet1.FreezePanes(range);

                Thread.Sleep(200);
            }

            for (var column = 0; column < 15; column++)
            {
                range = new Range(-1, column, 0, 0, sheet1);
                sheet1.FreezePanes(range);

                Thread.Sleep(200);
            }
        }

        [Fact]
        public async void AddWorksheetsBeforeAndAfter()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var sheet1 = workbook.Worksheets.First();
            var sheet2 = workbook.Worksheets.Last();
            Assert.Equal("2", sheet2.Name);

            // 1  | 2

            // Add before "1"
            var worksheet = (Worksheet)workbook.AddWorksheet(null, sheet1);
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
            worksheet = (Worksheet)workbook.AddWorksheet(null, sheet2);
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

        [Fact]
        public async void AddWorksheetsByIndex()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public async void CellTagContainsCustomObject()
        {
            this.ExpectedContextTags = new List<ContextTag>();

            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var worksheet = (Worksheet)workbook.AddWorksheet(null, null, workbook.Worksheets.Last());

            worksheet.CellsChanged += this.Worksheet_CellsChanged;

            var tag1 = new ContextTag { Context = "Cell00" };
            var tag2 = new ContextTag { Context = "Cell01" };
            this.ExpectedContextTags.Add(tag1);
            this.ExpectedContextTags.Add(tag2);

            var cell00 = worksheet[(0, 0)];
            cell00.Tag = tag1;

            Assert.NotEmpty(this.ExpectedContextTags);

            // Change the cell will trigger the Change Event
            this.ExpectedContextTag = tag1;
            worksheet.InteropWorksheet.Cells[1, 1] = "i am cell00";

            var cell01 = worksheet[(0, 1)];
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
            var tag = (ContextTag)e.Cells[0].Tag;

            Assert.Equal(this.ExpectedContextTag, tag);

            this.ExpectedContextTags.Remove(tag);

        }

        [Fact]
        public async void IsVisible()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

            var worksheet = workbook.Worksheets.First();
            Assert.True(worksheet.IsVisible);

            worksheet.IsVisible = false;

            Assert.False(worksheet.IsVisible);

            worksheet.IsVisible = true;

            Assert.True(worksheet.IsVisible);
        }


        [Fact]
        public async void AddWorkbook()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            Assert.Equal(2, workbook.Worksheets.Length);

        }

        [Fact]
        public async void InsertRows()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void DeleteRows()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void InsertColumn()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void InsertColumns()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void DeleteColumn()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void DeleteColumns()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            for (var i = 0; i < 10; i++)
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

        [Fact]
        public async void SetIsActiveWorksheet()
        {
            var program = new Mock<IProgram>();
            ICell cell = null;

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public void GetRange()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public async void GetUsedRange()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public async void GetUsedRangeColumn()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public async void GetUsedRangeRow()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

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

        [Fact]
        public void GetRectangle()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.Worksheets[1];

            sheet1.SetNamedRange("image", new Range(1, 1, 2, 2, sheet1));

            var rectangle = sheet1.GetRectangle("image");

            Assert.True(rectangle.Left > 0);

        }

        [Fact]
        public void NewSheetIsActive()
        {
            var program = new Mock<IProgram>();

            var addIn = new AddIn(this.application, program.Object);

            this.application.Workbooks.Add();

            var workbook = addIn.Workbooks[0];

            var sheet1 = workbook.AddWorksheet(0);

            Assert.True(sheet1.IsActive);

        }
    }
}
