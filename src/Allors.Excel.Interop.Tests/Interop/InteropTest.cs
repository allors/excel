using System;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;

namespace Allors.Excel.Tests.Interop
{
    public abstract class InteropTest : IDisposable
    {
        protected InteropApplication application;
        public abstract void Dispose();
    }
}
