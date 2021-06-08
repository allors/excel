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
            serviceLocator = new ServiceLocator();
            var program = new Program(serviceLocator);
            var office = new OfficeCore();

            addIn = new AddIn(Application, program, office);

            Ribbon.AddIn = addIn;
            await program.OnStart(addIn);
        });

        private async void ThisAddIn_Shutdown(object sender, EventArgs e) => await Task.Run(async () =>
        {
            await addIn.Program.OnStop();
        });

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            SynchronizationContext windowsFormsSynchronizationContext = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(windowsFormsSynchronizationContext);

            Ribbon = new Ribbon();
            return Ribbon;
        }

        public Ribbon Ribbon { get; set; }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
