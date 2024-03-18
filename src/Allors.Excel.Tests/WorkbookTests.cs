// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Tests.Interop
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Xml;
    using Moq;
    using Xunit;
    using Range = Excel.Range;

    public abstract class WorkbookTests : ExcelTest
    {
        [Fact]
        public async void OnNew()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            program.Verify(mock => mock.OnNew(It.IsAny<IWorkbook>()), Times.Once());

            await Task.CompletedTask;
        }

        [Fact]
        public void BuiltinProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.BuiltinProperties;

            properties.Title = "MyTitle";
            Assert.Equal("MyTitle", properties.Title);

            properties.Subject = "MySubject";
            Assert.Equal("MySubject", properties.Subject);

            properties.Author = "John Doe";
            Assert.Equal("John Doe", properties.Author);

            properties.Keywords = "SomeKeywords";
            Assert.Equal("SomeKeywords", properties.Keywords);

            properties.Comments = "SomeComments";
            Assert.Equal("SomeComments", properties.Comments);

            properties.Template = "MyTemplate";
            Assert.Equal("MyTemplate", properties.Template);

            properties.LastAuthor = "Jane Doe";
            Assert.Equal("Jane Doe", properties.LastAuthor);

            // TODO: Is RevisionNumber writable?
            Assert.Null(properties.RevisionNumber);

            properties.ApplicationName = "MyApplication";
            Assert.Equal("MyApplication", properties.ApplicationName);

            properties.LastPrintDate = DateTime.Today;
            Assert.Equal(DateTime.Today, properties.LastPrintDate);

            properties.CreationDate = DateTime.Today;
            Assert.Equal(DateTime.Today, properties.CreationDate);

            properties.LastSaveTime = DateTime.Today;
            Assert.Equal(DateTime.Today, properties.LastSaveTime);

            properties.TotalEditingTime = 100;
            Assert.Equal(100, properties.TotalEditingTime);

            properties.NumberOfPages = 100;
            Assert.Equal(100, properties.NumberOfPages);

            properties.NumberOfWords = 100;
            Assert.Equal(100, properties.NumberOfWords);

            properties.NumberOfCharacters = 100;
            Assert.Equal(100, properties.NumberOfCharacters);

            properties.Security = 100;
            Assert.Equal(100, properties.Security);

            properties.Category = "John Doe";
            Assert.Equal("John Doe", properties.Category);

            properties.Format = "John Doe";
            Assert.Equal("John Doe", properties.Format);

            properties.Manager = "John Doe";
            Assert.Equal("John Doe", properties.Manager);

            properties.Company = "John Doe";
            Assert.Equal("John Doe", properties.Company);

            properties.NumberOfBytes = 1;
            Assert.Equal(1, properties.NumberOfBytes);

            properties.NumberOfLines = 2;
            Assert.Equal(2, properties.NumberOfLines);

            properties.NumberOfParagraphs = 3;
            Assert.Equal(3, properties.NumberOfParagraphs);

            properties.NumberOfSlides = 4;
            Assert.Equal(4, properties.NumberOfSlides);

            properties.NumberOfNotes = 5;
            Assert.Equal(5, properties.NumberOfNotes);

            properties.NumberOfHiddenSlides = 6;
            Assert.Equal(6, properties.NumberOfHiddenSlides);

            properties.NumberOfMultimediaClips = 7;
            Assert.Equal(7, properties.NumberOfMultimediaClips);

            properties.HyperlinkBase = "href://example.com";
            Assert.Equal("href://example.com", properties.HyperlinkBase);

            properties.NumberOfCharactersWithSpaces = 8;
            Assert.Equal(8, properties.NumberOfCharactersWithSpaces);

            properties.ContentType = "doc";
            Assert.Equal("doc", properties.ContentType);

            properties.ContentStatus = "ok";
            Assert.Equal("ok", properties.ContentStatus);

            properties.Language = "nl";
            Assert.Equal("nl", properties.Language);

            properties.DocumentVersion = "1.0";
            Assert.Equal("1.0", properties.DocumentVersion);
        }

        [Fact]
        public void CustomBooleanProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            void doRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            var removeOrNotOptions = new[] { false, true };
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyBooleanProperty";

                properties.SetBoolean(prop, true);
                Assert.Equal(true, properties.GetBoolean(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetBoolean(prop, false);
                Assert.Equal(false, properties.GetBoolean(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetBoolean(prop, false);
                Assert.Equal(false, properties.GetBoolean(prop));

                properties.SetBoolean(prop, null);
                Assert.Null(properties.GetDate(prop));
            }
        }

        [Fact]
        public void CustomDateProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            void doRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            DateTime trim(DateTime now) => new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second, now.Kind);

            var removeOrNotOptions = new[] { false, true };
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyDateProperty";

                var birthDay = new DateTime(1973, 3, 27, 0, 0, 0, DateTimeKind.Local);
                properties.SetDate(prop, birthDay);
                Assert.Equal(birthDay, properties.GetDate(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetDate(prop, trim(DateTime.MinValue));
                Assert.True(properties.GetDate(prop) < DateTime.Now.AddYears(-100));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetDate(prop, trim(DateTime.MaxValue));
                Assert.Equal(trim(DateTime.MaxValue), properties.GetDate(prop));

                properties.SetDate(prop, null);
                Assert.Null(properties.GetDate(prop));
            }
        }

        [Fact]
        public void CustomFloatProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            void doRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            var removeOrNotOptions = new[] { false, true };

            // Int32 
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyFloatProperty";

                properties.SetFloat(prop, 0);
                Assert.Equal(0, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, 1);
                Assert.Equal(1, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1);
                Assert.Equal(-1, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, int.MinValue);
                Assert.Equal(int.MinValue, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, int.MaxValue);
                Assert.Equal(int.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }

            // Float32
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyFloatProperty";

                properties.SetFloat(prop, 1.1f);
                Assert.Equal(1.1f, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1.1f);
                Assert.Equal(-1.1f, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, float.MinValue);
                Assert.Equal(float.MinValue, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, float.MaxValue);
                Assert.Equal(float.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }

            // Float64
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyFloatProperty";

                properties.SetFloat(prop, 1.1d);
                Assert.Equal(1.1d, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1.1d);
                Assert.Equal(-1.1d, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, double.MinValue);
                Assert.Equal(double.MinValue, properties.GetFloat(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, double.MaxValue);
                Assert.Equal(double.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }
        }

        [Fact]
        public void CustomNumberProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            void doRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            var removeOrNotOptions = new[] { false, true };

            // Int32 
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyNumberProperty";

                properties.SetNumber(prop, 0);
                Assert.Equal(0, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, 1);
                Assert.Equal(1, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, -1);
                Assert.Equal(-1, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, int.MinValue);
                Assert.Equal(int.MinValue, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, int.MaxValue);
                Assert.Equal(int.MaxValue, properties.GetNumber(prop));

                properties.SetNumber(prop, null);
                Assert.Null(properties.GetNumber(prop));
            }

            // Int64 
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyNumberProperty";

                properties.SetNumber(prop, 0L);
                Assert.Equal(0L, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, 1L);
                Assert.Equal(1L, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, -1L);
                Assert.Equal(-1L, properties.GetNumber(prop));

                doRemoveOrNot(removeOrNot, prop);

                var tooLarge = false;
                try
                {
                    properties.SetNumber(prop, long.MinValue);
                }
                catch
                {
                    tooLarge = true;
                }

                Assert.True(tooLarge);

                doRemoveOrNot(removeOrNot, prop);

                tooLarge = false;
                try
                {
                    properties.SetNumber(prop, long.MaxValue);
                }
                catch
                {
                    tooLarge = true;
                }

                Assert.True(tooLarge);

                properties.SetNumber(prop, null);
                Assert.Null(properties.GetNumber(prop));
            }
        }

        [Fact]
        public void CustomStringProperties()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            void doRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            var removeOrNotOptions = new[] { false, true };

            // Strings 
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyStringProperty";

                properties.SetString(prop, string.Empty);
                Assert.Equal(string.Empty, properties.GetString(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetString(prop, "a string");
                Assert.Equal("a string", properties.GetString(prop));

                doRemoveOrNot(removeOrNot, prop);

                properties.SetString(prop, null);
                Assert.Null(properties.GetString(prop));
            }

            // Large strings 
            foreach (var removeOrNot in removeOrNotOptions)
            {
                var prop = "MyStringProperty";

                var length = 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                doRemoveOrNot(removeOrNot, prop);

                length = 1000 * 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                doRemoveOrNot(removeOrNot, prop);

                length = 1000 * 1000 * 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                doRemoveOrNot(removeOrNot, prop);
            }
        }

        [Fact]
        public void SetCustomXMLParts()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXML(xmlDoc);

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXMLById(tagId);
            Assert.Equal("CATALOG", outputXmlDoc.DocumentElement.Name);

            Assert.Equal(36, outputXmlDoc.DocumentElement.ChildNodes.Count);
        }

        [Fact]
        public void DeleteCustomXMLParts()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXML(xmlDoc);


            Assert.True(workbook.TryDeleteCustomXMLById(tagId));

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXMLById(Convert.ToString(tagId));
            Assert.Null(outputXmlDoc);

        }

        [Fact]
        public void SetNamedRangeWorkbook()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));

            namedRanges = iWorksheet.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }

        [Fact]
        public void SetNamedRangeWorksheet()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, iWorksheet);

            iWorksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = iWorksheet.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "'2'!MY.NAMEDRANGE"));

            namedRanges = workbook.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }

        [Fact]
        public void UpdateNamedRangeWorkbook()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, iWorksheet);

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

        [Fact]
        public void UpdateNamedRangeWorksheet()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, iWorksheet);

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

        [Fact]
        public void GetNamedRangeWorkbookForWorksheet()
        {
            var program = new Mock<IProgram>();
            var ribbon = new Mock<IRibbon>();
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges("2");

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }
    }
}
