// <copyright file="WorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.IO;
using Allors.Excel;
using Moq;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Interop.Vsto.Tests
{
    using Allors.Excel.Interop;
    using Allors.Excel.Tests;

    public class WorkbookTests : Allors.Excel.Tests.WorkbookTests
    {
        private readonly InteropApplication application;

        private readonly List<string> tempWorkbookFiles = new List<string>();

        public WorkbookTests()
        {
            this.application = new InteropApplication { Visible = true };
            this.DisconnectResidentAddIn();
        }

        public override void Dispose()
        {
            var workbooks = this.application.Workbooks;
            foreach (InteropWorkbook workbook in workbooks)
            {
                workbook.Close(false);
            }

            this.application.Quit();

            foreach (var file in this.tempWorkbookFiles)
            {
                try
                {
                    File.Delete(file);
                }
                catch
                {
                    // ignored
                }
            }
        }

        protected override IAddIn NewAddIn()
        {
            var ribbon = new Mock<IRibbon>();
            var addIn = new AddIn(this.application, new TestProgram(), ribbon.Object);

            // Mirror the Headless AddIn constructor, which creates an initial workbook.
            this.AddWorkbook();

            return addIn;
        }

        // Open a copy of a known single-sheet workbook: machine-independent, unlike
        // Workbooks.Add(), whose sheet count depends on the user's default template.
        // TestProgram.OnNew(workbook) adds the second sheet; both get numeric names.
        protected override void AddWorkbook()
        {
            var source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "blank.xlsx");
            var copy = Path.Combine(Path.GetTempPath(), $"AllorsExcelTest_{Guid.NewGuid():N}.xlsx");
            File.Copy(source, copy);
            this.tempWorkbookFiles.Add(copy);

            this.application.Workbooks.Open(copy);
        }

        // The repo's own VSTO add-in may be installed on developer machines; it loads
        // into every Excel instance and renames/adds worksheets, breaking determinism.
        // Disconnect it for this instance only (the registry LoadBehavior is untouched).
        private void DisconnectResidentAddIn()
        {
            dynamic app = this.application;
            var comAddIns = app.COMAddIns;
            for (var i = 1; i <= comAddIns.Count; i++)
            {
                var comAddIn = comAddIns.Item(i);
                if ((string)comAddIn.ProgId == "Allors.Excel.Interop.Vsto" && (bool)comAddIn.Connect)
                {
                    comAddIn.Connect = false;
                }
            }
        }
    }
}
