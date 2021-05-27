using System;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Interop
{
    public class InteropTest : IDisposable
    {
        protected const string skipReason = "Azure vmimage has no office installed";
        //protected const string skipReason = null;

        protected InteropApplication application;

        public InteropTest()
        {
            application = new InteropApplication { Visible = true };
        }

        public void Dispose()
        {
            var workbooks = application.Workbooks;
            foreach (InteropWorkbook workbook in workbooks)
            {
                workbook.Close(false);
            }


            application.Quit();
        }
    }
}
