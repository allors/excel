using System;

namespace Allors.Excel.Tests.Interop
{
    public abstract class ExcelTest : IDisposable
    {
        public abstract void Dispose();

        protected abstract IAddIn NewAddIn();

        protected abstract void AddWorkbook();
    }
}
