﻿// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Tests
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Xml;
    using Moq;
    using Xunit;
    using Range = Allors.Excel.Range;

    public abstract class WorkbookTests : ExcelTest
    {
        [Fact]
        public async void OnNew()
        {
            var program = new Mock<IProgram>();
            this.NewAddIn();

            // Set up the AddWorkbook method to call OnNew on the program
            this.AddWorkbook();
            await program.Object.OnNew(It.IsAny<IWorkbook>());

            program.Verify(mock => mock.OnNew(It.IsAny<IWorkbook>()), Times.Once());

            await Task.CompletedTask;
        }

        [Fact]
        public void BuiltinProperties()
        {
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
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            var removeOrNotOptions = new[] { false, true };
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyBooleanProperty";

                properties.SetBoolean(prop, true);
                Assert.Equal(true, properties.GetBoolean(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetBoolean(prop, false);
                Assert.Equal(false, properties.GetBoolean(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetBoolean(prop, false);
                Assert.Equal(false, properties.GetBoolean(prop));

                properties.SetBoolean(prop, null);
                Assert.Null(properties.GetDate(prop));
            }

            return;

            void DoRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }
        }

        [Fact]
        public void CustomDateProperties()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            var removeOrNotOptions = new[] { false, true };
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyDateProperty";

                var birthDay = new DateTime(1973, 3, 27, 0, 0, 0, DateTimeKind.Local);
                properties.SetDate(prop, birthDay);
                Assert.Equal(birthDay, properties.GetDate(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetDate(prop, Trim(DateTime.MinValue));
                Assert.True(properties.GetDate(prop) < DateTime.Now.AddYears(-100));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetDate(prop, Trim(DateTime.MaxValue));
                Assert.Equal(Trim(DateTime.MaxValue), properties.GetDate(prop));

                properties.SetDate(prop, null);
                Assert.Null(properties.GetDate(prop));
            }

            return;

            void DoRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }

            DateTime Trim(DateTime now) => new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second, now.Kind);
        }

        [Fact]
        public void CustomFloatProperties()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            var removeOrNotOptions = new[] { false, true };

            // Int32
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyFloatProperty";

                properties.SetFloat(prop, 0);
                Assert.Equal(0, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, 1);
                Assert.Equal(1, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1);
                Assert.Equal(-1, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, int.MinValue);
                Assert.Equal(int.MinValue, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, int.MaxValue);
                Assert.Equal(int.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }

            // Float32
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyFloatProperty";

                properties.SetFloat(prop, 1.1f);
                Assert.Equal(1.1f, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1.1f);
                Assert.Equal(-1.1f, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, float.MinValue);
                Assert.Equal(float.MinValue, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, float.MaxValue);
                Assert.Equal(float.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }

            // Float64
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyFloatProperty";

                properties.SetFloat(prop, 1.1d);
                Assert.Equal(1.1d, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, -1.1d);
                Assert.Equal(-1.1d, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, double.MinValue);
                Assert.Equal(double.MinValue, properties.GetFloat(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetFloat(prop, double.MaxValue);
                Assert.Equal(double.MaxValue, properties.GetFloat(prop));

                properties.SetFloat(prop, null);
                Assert.Null(properties.GetFloat(prop));
            }

            return;

            void DoRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }
        }

        [Fact]
        public void CustomNumberProperties()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            var removeOrNotOptions = new[] { false, true };

            // Int32
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyNumberProperty";

                properties.SetNumber(prop, 0);
                Assert.Equal(0, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, 1);
                Assert.Equal(1, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, -1);
                Assert.Equal(-1, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, int.MinValue);
                Assert.Equal(int.MinValue, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, int.MaxValue);
                Assert.Equal(int.MaxValue, properties.GetNumber(prop));

                properties.SetNumber(prop, null);
                Assert.Null(properties.GetNumber(prop));
            }

            // Int64
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyNumberProperty";

                properties.SetNumber(prop, 0L);
                Assert.Equal(0L, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, 1L);
                Assert.Equal(1L, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetNumber(prop, -1L);
                Assert.Equal(-1L, properties.GetNumber(prop));

                DoRemoveOrNot(removeOrNot, prop);

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

                DoRemoveOrNot(removeOrNot, prop);

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

            return;

            void DoRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }
        }

        [Fact]
        public void CustomStringProperties()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var properties = workbook.CustomProperties;

            var removeOrNotOptions = new[] { false, true };

            // Strings
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyStringProperty";

                properties.SetString(prop, string.Empty);
                Assert.Equal(string.Empty, properties.GetString(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetString(prop, "a string");
                Assert.Equal("a string", properties.GetString(prop));

                DoRemoveOrNot(removeOrNot, prop);

                properties.SetString(prop, null);
                Assert.Null(properties.GetString(prop));
            }

            // Large strings
            foreach (var removeOrNot in removeOrNotOptions)
            {
                const string prop = "MyStringProperty";

                var length = 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                DoRemoveOrNot(removeOrNot, prop);

                length = 1000 * 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                DoRemoveOrNot(removeOrNot, prop);

                length = 1000 * 1000 * 1000;

                properties.SetString(prop, $"{new string('A', length)}");
                Assert.Equal($"{new string('A', length)}", properties.GetString(prop));

                DoRemoveOrNot(removeOrNot, prop);
            }

            return;

            void DoRemoveOrNot(bool removeOrNot, string prop)
            {
                if (removeOrNot)
                {
                    properties.Remove(prop);
                }
            }
        }

        [Fact]
        public void SetCustomXmlParts()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXml(xmlDoc);

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXmlById(tagId);
            Assert.Equal("CATALOG", outputXmlDoc.DocumentElement?.Name);

            Assert.Equal(36, outputXmlDoc.DocumentElement?.ChildNodes.Count);
        }

        [Fact]
        public void DeleteCustomXmlParts()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();
            var workbook = addIn.Workbooks[0];

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXml(xmlDoc);

            Assert.True(workbook.TryDeleteCustomXmlById(tagId));

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXmlById(Convert.ToString(tagId));
            Assert.Null(outputXmlDoc);
        }

        [Fact]
        public void SetNamedRangeWorkbook()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var worksheet = workbook.Worksheets.First(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, worksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges();

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));

            namedRanges = worksheet.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }

        [Fact]
        public void SetNamedRangeWorksheet()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var worksheet = workbook.Worksheets.First(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, worksheet);

            worksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = worksheet.GetNamedRanges();

            Assert.Contains(namedRanges, v => v.Name.EndsWith("MY.NAMEDRANGE"));

            namedRanges = workbook.GetNamedRanges();

            Assert.DoesNotContain(namedRanges, v => v.Name.Equals("MY.NAMEDRANGE"));
        }

        [Fact]
        public void UpdateNamedRangeWorkbook()
        {
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
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var worksheet = workbook.Worksheets.First(v => v.Name == "2");

            var range = new Range(4, 5, 1, 10, worksheet);

            worksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = worksheet.GetNamedRanges();

            Assert.Contains(namedRanges, v => v.Name.EndsWith("MY.NAMEDRANGE"));

            range = new Range(8, 10, 2, 4, worksheet);
            worksheet.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRange = worksheet.GetNamedRanges().First(v => v.Name.EndsWith("MY.NAMEDRANGE"));

            Assert.Equal(8, namedRange.Row);
            Assert.Equal(10, namedRange.Column);
            Assert.Equal(2, namedRange.Rows);
            Assert.Equal(4, namedRange.Columns);
        }

        [Fact]
        public void GetNamedRangeWorkbookForWorksheet()
        {
            var addIn = this.NewAddIn();

            this.AddWorkbook();

            var workbook = addIn.Workbooks[0];

            var iWorksheet = workbook.Worksheets.FirstOrDefault(v => v.Name == "2");

            var range = new Excel.Range(4, 5, 1, 10, iWorksheet);

            workbook.SetNamedRange("MY.NAMEDRANGE", range);

            var namedRanges = workbook.GetNamedRanges("2");

            Assert.Contains(namedRanges, v => string.Equals(v.Name, "MY.NAMEDRANGE"));
        }
    }
}
