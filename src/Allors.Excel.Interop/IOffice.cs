using System.Xml;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Interop
{
    public interface IOffice
    {
        object MsoPropertyTypeString { get; }
        object MsoPropertyTypeBoolean { get; }
        object MsoPropertyTypeDate { get; }
        object MsoPropertyTypeFloat { get; }
        object MsoPropertyTypeNumber { get; }

        void AddPicture(InteropWorksheet interopWorksheet, string filename, System.Drawing.Rectangle rectangle);

        XmlDocument GetCustomXMLById(InteropWorkbook interopWorkbook, string id);
        string SetCustomXmlPart(InteropWorkbook interopWorkbook, XmlDocument xmlDocument);
        bool TryDeleteCustomXMLById(InteropWorkbook interopWorkbook, string id);
    }
}
