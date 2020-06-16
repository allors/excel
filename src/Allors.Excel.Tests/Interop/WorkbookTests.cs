// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Linq;
using System.Xml;
using Allors.Excel.Interop;
using Moq;
using Xunit;

namespace Allors.Excel.Tests.Interop
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
        public void SetCustomProperties()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            var theDate = DateTime.Now;

            var customerProperties = new CustomProperties();
            customerProperties.Add("Company.Name", "Zonsoft.be");
            customerProperties.Add("Company.Street", "Uikhoverstraat 158");
            customerProperties.Add("Company.City", "3631 Maasmechelen");
            customerProperties.Add("Company.Country", "BE België");

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);

            customerProperties = new CustomProperties();
            customerProperties.Add("Showcase.IsInvoiceSheet", false);
            customerProperties.Add("Showcase.Date", theDate);
            customerProperties.Add("Showcase.Decimal", 123.45M);
            
            var nullableDecimal = new Nullable<decimal>(123.45M);
            customerProperties.Add("Showcase.NullableDecimal", nullableDecimal);

            customerProperties.Add("Showcase.Int", 12);

            var nullableInt = new Nullable<int>(12);
            customerProperties.Add("Showcase.NullableInt", nullableInt);

            customerProperties.Add("Company.Name", "Zonsoft.be");
            customerProperties.Add("Company.Street", "Uikhoverstraat 158");

            // Duplicates will be overwritten
            customerProperties.Add("Company.City", "3631 Maasmechelen");
            customerProperties.Add("Company.City", "3631 Uikhoven");

            customerProperties.Add("Company.Country", "BE België");

            customerProperties.Add("Showcase.Null", null);

            workbook.SetCustomProperties(customerProperties);

            customProperties = workbook.GetCustomProperties();

            Assert.Equal(customerProperties.Count, customProperties.Count);

            Assert.False(customProperties.Get<bool>("Showcase.IsInvoiceSheet"));


            // fractions of MS are not preserved!
            Assert.Equal(theDate.Date, customProperties.Get<DateTime>("Showcase.Date").Date);

            Assert.Equal(12, customProperties.Get<int>("Showcase.Int"));
            Assert.Equal(12, customProperties.Get<int?>("Showcase.NullableInt"));
            Assert.Null(customProperties.Get<int?>("Showcase.Null"));

            Assert.Equal(123.45M, customProperties.Get<decimal>("Showcase.Decimal"));
            Assert.Equal(123.45M, customProperties.Get<decimal>("Showcase.NullableDecimal"));

            Assert.Equal("Zonsoft.be", customProperties.Get<string>("Company.Name"));
            Assert.Equal("BE België", customProperties.Get<string>("Company.Country"));
            Assert.Equal("3631 Uikhoven", customProperties.Get<string>("Company.City"));

            object res = null;

            Assert.True(workbook.TryGetCustomProperty("Company.Name", ref res));

            var del = new CustomProperties();
            del.Add("Company.Name", null);

            workbook.DeleteCustomProperties(del);

            Assert.False(workbook.TryGetCustomProperty("Company.Name", ref res));
            Assert.True(workbook.TryGetCustomProperty("Company.Country", ref res));

            customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count - 1, customProperties.Count);
        }

        [Fact(Skip = skipReason)]
        public void SetManyKeysCustomProperties()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];
                      
            var customerProperties = new CustomProperties();

            foreach (int i in Enumerable.Range(0, 64))
            {
                customerProperties.Add($"key{i}", $"value.{i}");
            }

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);
        }

        [Fact(Skip = skipReason)]
        public void SetLargeStringValueCustomPropertiesTruncatesTo255()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            var customerProperties = new CustomProperties();
                      
            customerProperties.Add($"keyA", $"{new String('A', 10)}");
            customerProperties.Add($"keyB", $"{new String('A', 255)}");
            customerProperties.Add($"keyC", $"{new String('B', 256)}");
            customerProperties.Add($"keyD", $"{new String('C', 1000)}");          

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);
            Assert.Equal(10, customProperties.Get<string>("keyA").Length);
            Assert.Equal(255, customProperties.Get<string>("keyB").Length);
            Assert.Equal(255, customProperties.Get<string>("keyC").Length);
            Assert.Equal(255, customProperties.Get<string>("keyD").Length);
        }

        [Fact(Skip = skipReason)]
        public void SetDecimalValueCustomProperties()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            var customerProperties = new CustomProperties();

            customerProperties.Add($"keyA", decimal.MaxValue);
            customerProperties.Add($"keyB", decimal.MinValue);
            customerProperties.Add($"keyC", 1.25M);
            customerProperties.Add($"keyD", 1.123456M);
            customerProperties.Add($"keyE", 1.123456789M);
            customerProperties.Add($"keyF", -1.12M);            

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);
            Assert.Equal(decimal.MaxValue, customProperties.Get<decimal>("keyA"));
            Assert.Equal(decimal.MinValue, customProperties.Get<decimal>("keyB"));
            Assert.Equal(1.25M, customProperties.Get<decimal>("keyC"));
            Assert.Equal(1.123456M, customProperties.Get<decimal>("keyD"));
            
            // value is truncated, and rounded!
            Assert.Equal(1.123457M, customProperties.Get<decimal>("keyE"));

            // Negative values
            Assert.Equal(-1.12M, customProperties.Get<decimal>("keyF"));
         
        }

        [Fact(Skip = skipReason)]
        public void SetIntValueCustomProperties()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            var customerProperties = new CustomProperties();

            customerProperties.Add($"keyA", int.MaxValue);
            customerProperties.Add($"keyB", int.MinValue);
            customerProperties.Add($"keyC", 0);
            customerProperties.Add($"keyD", 1);
            customerProperties.Add($"keyE", -1);
            customerProperties.Add($"keyF", 1000);

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);
            Assert.Equal(int.MaxValue, customProperties.Get<decimal>("keyA"));
            Assert.Equal(int.MinValue, customProperties.Get<decimal>("keyB"));
            Assert.Equal(0, customProperties.Get<decimal>("keyC"));
            Assert.Equal(1, customProperties.Get<decimal>("keyD"));
            Assert.Equal(-1, customProperties.Get<decimal>("keyE"));
            Assert.Equal(1000, customProperties.Get<decimal>("keyF"));
        }

        [Fact(Skip = skipReason)]
        public void SetDateTimeValueCustomProperties()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            var customerProperties = new CustomProperties();

            customerProperties.Add($"keyA", DateTime.MaxValue);
            customerProperties.Add($"keyB", DateTime.MinValue);

            var cDate = new DateTime(2020, 5, 15, 10, 15, 20);
            customerProperties.Add($"keyC", cDate);          

            workbook.SetCustomProperties(customerProperties);

            var customProperties = workbook.GetCustomProperties();
            Assert.Equal(customerProperties.Count, customProperties.Count);
                       
            Assert.Equal(DateTime.MaxValue, customProperties.Get<DateTime>("keyA"));

            Assert.Equal(DateTime.MinValue, customProperties.Get<DateTime>("keyB"));

            var keyCDate = customProperties.Get<DateTime>("keyC");

            Assert.Equal(cDate.AddSeconds(-1), keyCDate);          
          
        }


        [Fact(Skip = skipReason)]
        public void SetCustomXMLParts()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXML(xmlDoc);

            // store the tag in the customproperties
            workbook.TrySetCustomProperty("##CATALOG", tagId);

            // First get the tag
            object tagId2 = null;
            workbook.TryGetCustomProperty("##CATALOG", ref tagId2);

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXMLById(Convert.ToString(tagId2));
            Assert.Equal("CATALOG", outputXmlDoc.DocumentElement.Name);

            Assert.Equal(36, outputXmlDoc.DocumentElement.ChildNodes.Count);

        }

        [Fact(Skip = skipReason)]
        public void DeleteCustomXMLParts()
        {
            var program = new Mock<IProgram>();
            var office = new Mock<IOffice>();

            var addIn = new AddIn(application, program.Object, office.Object);

            application.Workbooks.Add();
            var workbook = addIn.Workbooks[0];

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(@"data\catalog.xml");

            var tagId = workbook.SetCustomXML(xmlDoc);


            Assert.True(workbook.TryDeleteCustomXMLById(tagId));

            // Then read the xml
            var outputXmlDoc = workbook.GetCustomXMLById(Convert.ToString(tagId));
            Assert.Null(outputXmlDoc);          

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
