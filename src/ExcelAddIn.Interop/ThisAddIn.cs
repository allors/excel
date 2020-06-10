// <copyright file="ThisAddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.ComponentModel;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Allors.Excel.Interop;
using Application;
using AppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Threading.Tasks;
using ExcelAddIn.Interop;
using ExcelAddIn.Services;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private ServiceLocator serviceLocator;
        private AddIn addIn;

        private async void ThisAddIn_Startup(object sender, System.EventArgs e) => await Task.Run(async () =>
        {
            this.serviceLocator = new ServiceLocator();
            var program = new Program(serviceLocator);

            var office = new Office(this);

            this.addIn = new AddIn(this.Application, program, office);
            this.Ribbon.AddIn = this.addIn;
            await program.OnStart(addIn);
        });

        private async void ThisAddIn_Shutdown(object sender, System.EventArgs e) => await Task.Run(async () =>
        {
            await this.addIn.Program.OnStop();
        });

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
