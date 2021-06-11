using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;

namespace ExcelDNA
{
    using System;
    using Allors.Excel.Interop;
    using Application;
    using ExcelDna.Integration;

    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            var application = ExcelDnaUtil.Application;
            var serviceLocator = new ServiceLocator();
            this.Program = new Program(serviceLocator);
            this.AddIn = new AddIn((InteropApplication)application, this.Program);
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
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
    }
}
