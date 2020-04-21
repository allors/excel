using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Allors.Excel
{
    public interface IOffice
    {
        void AddPicture(Worksheet interopWorksheet, string filename, System.Drawing.Rectangle rectangle);
    }
}
