using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Tests.Embedded
{
    public class InteropTest : IDisposable
    {
        //protected const string skipReason = "Azure vmimage has no office installed";
        protected const string skipReason = null;

        protected InteropApplication application;

        public InteropTest()
        {
            this.application = new InteropApplication { Visible = true };
        }

        public void Dispose()
        {
            var workbooks = this.application.Workbooks;
            foreach (InteropWorkbook workbook in workbooks)
            {
                workbook.Close(false);
            }


            this.application.Quit();
        }
    }
}
