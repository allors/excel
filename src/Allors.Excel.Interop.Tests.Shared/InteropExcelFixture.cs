// <copyright file="InteropExcelFixture.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Moq;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Allors.Excel.Interop.Tests.Shared
{
    using Allors.Excel.Tests;

    /// <summary>
    /// Hosts a real Excel instance for the interop test suites: one instance per
    /// test, torn down crash-tolerantly.
    /// </summary>
    public sealed class InteropExcelFixture : IDisposable
    {
        private readonly InteropApplication application;

        private readonly int processId;

        private readonly List<string> tempWorkbookFiles = new List<string>();

        public InteropExcelFixture()
        {
            this.application = new InteropApplication { Visible = true };

            // Capture the pid now, while Excel is healthy: resolving it from the Hwnd
            // is itself a COM call, so it cannot be done once Excel has crashed.
            GetWindowThreadProcessId((IntPtr)this.application.Hwnd, out var pid);
            this.processId = (int)pid;

            this.DisconnectResidentAddIn();
        }

        public IAddIn NewAddIn()
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
        public void AddWorkbook()
        {
            var source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "blank.xlsx");
            var copy = Path.Combine(Path.GetTempPath(), $"AllorsExcelTest_{Guid.NewGuid():N}.xlsx");
            File.Copy(source, copy);
            this.tempWorkbookFiles.Add(copy);

            this.application.Workbooks.Open(copy);
        }

        public void Dispose()
        {
            try
            {
                // Snapshot before closing: removing workbooks while enumerating the
                // live COM collection is undefined.
                var workbooks = new List<InteropWorkbook>();
                foreach (InteropWorkbook workbook in this.application.Workbooks)
                {
                    workbooks.Add(workbook);
                }

                foreach (var workbook in workbooks)
                {
                    workbook.Close(false);
                }

                this.application.Quit();
            }
            catch (COMException)
            {
                // Excel can crash mid-close (observed on CI: RPC failure 0x800706BE
                // closing a workbook that carries an input-message validation). The
                // test's assertions have already run; kill the instance so a hung
                // Excel cannot linger and destabilize later tests.
                this.KillExcelProcess();
            }

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

        // Last-resort cleanup when graceful Close/Quit failed; the process is
        // usually already dead by then.
        private void KillExcelProcess()
        {
            try
            {
                using (var process = Process.GetProcessById(this.processId))
                {
                    process.Kill();
                    process.WaitForExit(5000);
                }
            }
            catch
            {
                // already exited
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
