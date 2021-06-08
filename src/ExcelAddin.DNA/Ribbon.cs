using System;
using System.Diagnostics;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Threading;
using Allors.Excel.Interop;
using Application;
using ExcelAddIn.DNA;
using ExcelDna.Integration;

namespace ExcelAddin.DNA
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            var application = ExcelDnaUtil.Application;
            var serviceLocator = new ServiceLocator();
            this.Program = new Program(serviceLocator, "DNA");
            var office = new OfficeCore();

            this.AddIn = new AddIn(application, this.Program, office);

            return RibbonResources.Ribbon;
        }

        public AddIn AddIn { get; private set; }

        public Program Program { get; private set; }

        public override void OnStartupComplete(ref Array custom)
        {
            base.OnStartupComplete(ref custom);

            ExcelAsyncUtil.Run("WebSnippetAsync", null,
                async delegate
                {
                    await this.Program.OnStart(this.AddIn);
                });
        }

        public override object LoadImage(string imageId)
        {
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
    }
}
