using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Allors.Excel.Interop
{
    public interface IOffice
    {
        void AddPicture(InteropWorksheet interopWorksheet, string filename, System.Drawing.Rectangle rectangle);
    }
}
