using Allors.Excel.Interop;
using Microsoft.Office.Core;

namespace ExcelAddIn.Interop
{
    public class Office : IOffice
    {
        private ThisAddIn thisAddIn;

        public Office(ThisAddIn thisAddIn)
        {
            this.thisAddIn = thisAddIn;
        }

        public void AddPicture(Microsoft.Office.Interop.Excel.Worksheet worksheet, string fileName, System.Drawing.Rectangle rectangle)
        {
             worksheet.Shapes.AddPicture(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }
    }
}
