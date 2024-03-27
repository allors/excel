using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;

namespace ExcelDNA
{
    using System;
    using System.Windows.Forms;
    using Allors.Excel;
    using Allors.Excel.Interop;
    using Application;
    using ExcelDna.Integration;

    [ComVisible(true)]
    public class Ribbon : ExcelRibbon, IRibbon
    {
        public IRibbonUI RibbonUI { get; private set; } = null!;

        public AddIn AddIn { get; private set; } = null!;

        public Program Program { get; private set; } = null!;

        public override string GetCustomUI(string _)
        {
            try
            {
                var application = ExcelDnaUtil.Application;
                var serviceLocator = new ServiceLocator();
                this.Program = new Program(serviceLocator);
                this.AddIn = new AddIn((InteropApplication)application, this.Program, this);
                return RibbonResources.Ribbon;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw;
            }
        }

        public async void OnLoad(IRibbonUI ribbon)
        {
            this.RibbonUI = ribbon;

            try
            {
                await this.Program.OnStart(this.AddIn);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw;
            }
        }

        public override object LoadImage(string imageId) =>
            // This will return the image resource with the name specified in the image='xxxx' tag
            RibbonResources.ResourceManager.GetObject(imageId);

        public void OnButtonPressed(IRibbonControl control) => System.Windows.Forms.MessageBox.Show("Hello!");

        public void Invalidate() => this.RibbonUI.Invalidate();
    }
}
