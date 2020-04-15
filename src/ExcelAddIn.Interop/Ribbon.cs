// <copyright file="Ribbon.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Allors.Excel.Embedded;
using Office = Microsoft.Office.Core;

namespace ExcelAddInLocal
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private string doSomethingLabel = "Do Something";

        public AddIn AddIn { get; set; }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn.Interop.Ribbon.xml");
        }

        #endregion

        #region Ribbon Labels

        public string DoSomethingLabel
        {
            get => doSomethingLabel;
            set
            {
                doSomethingLabel = value;
                this.ribbon.Invalidate();
            }
        }

        public string GetDoSomethingLabel(Office.IRibbonControl control)
        {
            return this.DoSomethingLabel;
        }
        
        #endregion

        #region Ribbon Callbacks

        public async void OnClick(Office.IRibbonControl control) => await Task.Run(async () =>
        {
            if (this.AddIn != null)
            {
                await this.AddIn.Program.OnHandle(control.Id);
            }
        });

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
