// <copyright file="ThisAddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Allors.Excel.Interop;
using Application;
using ExcelAddIn.Interop;
using ExcelAddIn.Services;
using Microsoft.Office.Core;

namespace ExcelAddIn
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
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
