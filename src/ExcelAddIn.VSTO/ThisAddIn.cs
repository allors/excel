using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Allors.Excel.Interop;
using Application;
using Microsoft.Office.Core;

namespace ExcelAddIn.VSTO
{
    public partial class ThisAddIn
    {
        private ServiceLocator serviceLocator;
        private AddIn addIn;

        private async void ThisAddIn_Startup(object sender, EventArgs e) => await Task.Run(async () =>
        {
            this.serviceLocator = new ServiceLocator();
            var program = new Program(this.serviceLocator);

            this.addIn = new AddIn(this.Application, program, this.Ribbon);

            this.Ribbon.AddIn = this.addIn;
            await program.OnStart(this.addIn);
        });

        private async void ThisAddIn_Shutdown(object sender, EventArgs e) => await Task.Run(async () =>
        {
            await this.addIn.Program.OnStop();
        });

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            SynchronizationContext windowsFormsSynchronizationContext = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(windowsFormsSynchronizationContext);

            this.Ribbon = new Ribbon();
            return this.Ribbon;
        }

        public Ribbon Ribbon { get; set; }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(this.ThisAddIn_Startup);
            this.Shutdown += new EventHandler(this.ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
