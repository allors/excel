using Allors.Excel;
using ExcelAddInLocal;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Interop
{
    public class Office : IOffice
    {
        private ThisAddIn thisAddIn;

        public Office(ThisAddIn thisAddIn)
        {
            this.thisAddIn = thisAddIn;
        }

        public void AddPicture(Worksheet interopWorksheet, string fileName, System.Drawing.Rectangle rectangle)
        {
             interopWorksheet.Shapes.AddPicture(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }
    }
}
