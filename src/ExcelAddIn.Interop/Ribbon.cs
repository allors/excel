// <copyright file="Ribbon.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Allors.Excel.Interop;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn
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
                ribbon.Invalidate();
            }
        }

        public string GetDoSomethingLabel(Office.IRibbonControl control)
        {
            return DoSomethingLabel;
        }

        #endregion

        #region Ribbon Callbacks

        public async void OnClick(Office.IRibbonControl control) => await Task.Run(async () =>
        {
            if (AddIn != null)
            {
                await AddIn.Program.OnHandle(control.Id);
            }
        });

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (var i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
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
