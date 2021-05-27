using System.Drawing;
using System.Xml;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Interop
{
    public interface IOfficeCore
    {
        object MsoPropertyTypeString { get; }
        object MsoPropertyTypeBoolean { get; }
        object MsoPropertyTypeDate { get; }
        object MsoPropertyTypeFloat { get; }
        object MsoPropertyTypeNumber { get; }

        void AddPicture(InteropWorksheet interopWorksheet, string filename, Rectangle rectangle);

        XmlDocument GetCustomXmlById(InteropWorkbook interopWorkbook, string id);

        string SetCustomXmlPart(InteropWorkbook interopWorkbook, XmlDocument xmlDocument);

        bool TryDeleteCustomXmlById(InteropWorkbook interopWorkbook, string id);
    }
}
