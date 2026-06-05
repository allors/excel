using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Allors.Excel.Interop;
using Application;
using Microsoft.Office.Core;

namespace Allors.Excel.Interop.Vsto
{
    public partial class ThisAddIn
    {
        private ServiceLocator serviceLocator;
        private AddIn addIn;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                this.serviceLocator = new ServiceLocator();
                var program = new Program(this.serviceLocator);

                // Initialize on the add-in's STA thread (do NOT offload to the thread pool):
                // the AddIn constructor subscribes to Excel Application COM events, which must
                // happen on the STA thread that owns the Excel object model.
                this.addIn = new AddIn(this.Application, program, this.Ribbon);
                this.Ribbon.AddIn = this.addIn;

                await program.OnStart(this.addIn);
            }
            catch (Exception exception)
            {
                // ThisAddIn_Startup is an async void event handler: an escaping exception
                // would tear down the Excel host. Surface it without crashing.
                MessageBox.Show(exception.ToString());
            }
        }

        private async void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                // Startup runs asynchronously and may not have completed (or may have
                // failed), so addIn can still be null at shutdown.
                if (this.addIn?.Program != null)
                {
                    await this.addIn.Program.OnStop();
                }
            }
            catch (Exception)
            {
                // Never throw from shutdown.
            }
        }

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
