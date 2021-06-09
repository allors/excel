// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System.Collections.Generic;
    using System.Linq;
    using InteropApplication = Microsoft.Office.Interop.Excel.Application;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropAppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;

    public class AddIn : IAddIn
    {
        private readonly Dictionary<InteropWorkbook, Workbook> workbookByInteropWorkbook;

        public AddIn(object application, IProgram program)
        {
            this.Application = (InteropApplication)application;
            this.Program = program;

            this.workbookByInteropWorkbook = new Dictionary<InteropWorkbook, Workbook>();

            ((InteropAppEvents_Event)this.Application).NewWorkbook += async interopWorkbook =>
            {
                var workbook = this.New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                var worksheets = workbook.Worksheets;
                await this.Program.OnNew(workbook);
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }
            };

            this.Application.WorkbookOpen += async interopWorkbook =>
            {
                var workbook = this.New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                var worksheets = workbook.Worksheets;
                await this.Program.OnNew(workbook);
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }
            };


            void WorkbookBeforeClose(InteropWorkbook interopWorkbook, ref bool cancel)
            {
                if (this.WorkbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
                {
                    this.Program.OnClose(workbook, ref cancel);
                    if (!cancel)
                    {
                        this.Close(interopWorkbook);
                    }
                }
            }

            this.Application.WorkbookActivate += wb =>
            {
                if (!this.WorkbookByInteropWorkbook.TryGetValue(wb, out var workbook))
                {
                    workbook = this.New(wb);
                }

                workbook.IsActive = true;
            };


            this.Application.WorkbookDeactivate += wb =>
            {
                // Could already be gone by the WorkbookBeforeClose event
                if (this.WorkbookByInteropWorkbook.TryGetValue(wb, out _))
                {
                    this.WorkbookByInteropWorkbook[wb].IsActive = false;
                }
            };

            this.Application.WorkbookBeforeClose += WorkbookBeforeClose;
        }

        public InteropApplication Application { get; }

        public IProgram Program { get; }

        public IReadOnlyDictionary<InteropWorkbook, Workbook> WorkbookByInteropWorkbook => this.workbookByInteropWorkbook;

        public IWorkbook[] Workbooks => this.WorkbookByInteropWorkbook.Values.Cast<IWorkbook>().ToArray();

        public Workbook New(InteropWorkbook interopWorkbook)
        {
            if (!this.workbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
            {
                workbook = new Workbook(this, interopWorkbook);
                this.workbookByInteropWorkbook.Add(interopWorkbook, workbook);
            }

            return workbook;
        }

        public void Close(InteropWorkbook interopWorkbook) => this.workbookByInteropWorkbook.Remove(interopWorkbook);
    }
}
