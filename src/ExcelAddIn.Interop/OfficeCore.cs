using System;
using System.Runtime.InteropServices;
using System.Xml;
using Allors.Excel.Interop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Rectangle = System.Drawing.Rectangle;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddIn.Interop
{
    public class OfficeCore : IOfficeCore
    {
        public object MsoPropertyTypeString => MsoDocProperties.msoPropertyTypeString;

        public object MsoPropertyTypeBoolean => MsoDocProperties.msoPropertyTypeBoolean;

        public object MsoPropertyTypeDate => MsoDocProperties.msoPropertyTypeDate;

        public object MsoPropertyTypeFloat => MsoDocProperties.msoPropertyTypeFloat;

        public object MsoPropertyTypeNumber => MsoDocProperties.msoPropertyTypeNumber;

        public void AddPicture(Worksheet worksheet, string fileName, Rectangle rectangle)
        {
            worksheet.Shapes.AddPicture(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }

        public XmlDocument GetCustomXmlById(Workbook interopWorkbook, string id)
        {
            var xmlDocument = new XmlDocument();
            var customXMLPart = interopWorkbook.CustomXMLParts.SelectByID(id);

            if (customXMLPart != null)
            {
                xmlDocument.LoadXml(customXMLPart.XML);

                return xmlDocument;
            }

            return null;
        }

        public string SetCustomXmlPart(Workbook interopWorkbook, XmlDocument xmlDocument)
        {
            return interopWorkbook.CustomXMLParts.Add(xmlDocument.OuterXml, Type.Missing).Id;
        }

        public bool TryDeleteCustomXmlById(Workbook interopWorkbook, string id)
        {
            try
            {
                var customXMLPart = interopWorkbook.CustomXMLParts.SelectByID(id);
                customXMLPart.Delete();
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
