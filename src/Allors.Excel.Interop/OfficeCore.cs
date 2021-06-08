using System;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Core;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using Rectangle = System.Drawing.Rectangle;

namespace Allors.Excel.Interop
{
    public class OfficeCore 
    {
        public object MsoPropertyTypeString => MsoDocProperties.msoPropertyTypeString;

        public object MsoPropertyTypeBoolean => MsoDocProperties.msoPropertyTypeBoolean;

        public object MsoPropertyTypeDate => MsoDocProperties.msoPropertyTypeDate;

        public object MsoPropertyTypeFloat => MsoDocProperties.msoPropertyTypeFloat;

        public object MsoPropertyTypeNumber => MsoDocProperties.msoPropertyTypeNumber;

        public void AddPicture(InteropWorksheet worksheet, string fileName, Rectangle rectangle)
        {
            worksheet.Shapes.AddPicture(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }

        public XmlDocument GetCustomXmlById(InteropWorkbook interopWorkbook, string id)
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

        public string SetCustomXmlPart(InteropWorkbook interopWorkbook, XmlDocument xmlDocument)
        {
            return interopWorkbook.CustomXMLParts.Add(xmlDocument.OuterXml, Type.Missing).Id;
        }

        public bool TryDeleteCustomXmlById(InteropWorkbook interopWorkbook, string id)
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
