// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Allors.Excel.Interop
{
    using InteropApplication = Application;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropAppEvents_Event = AppEvents_Event;

    public class AddIn : IAddIn
    {
        private readonly Dictionary<InteropWorkbook, Workbook> workbookByInteropWorkbook;

        public AddIn(InteropApplication application, IProgram program, IOfficeCore officeCore)
        {
            Application = application;
            Program = program;
            OfficeCore = officeCore;

            workbookByInteropWorkbook = new Dictionary<InteropWorkbook, Workbook>();

            ((InteropAppEvents_Event)Application).NewWorkbook += async interopWorkbook =>
            {
                var workbook = New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                var worksheets = workbook.Worksheets;
                await Program.OnNew(workbook);
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }
            };

            Application.WorkbookOpen += async interopWorkbook =>
            {
                var workbook = New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                var worksheets = workbook.Worksheets;
                await Program.OnNew(workbook);
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }
            };


            void WorkbookBeforeClose(InteropWorkbook interopWorkbook, ref bool cancel)
            {
                if (WorkbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
                {
                    Program.OnClose(workbook, ref cancel);
                    if (!cancel)
                    {
                        Close(interopWorkbook);
                    }
                }
            }

            Application.WorkbookActivate += wb =>
            {
                if (!WorkbookByInteropWorkbook.TryGetValue(wb, out var workbook))
                {
                    workbook = New(wb);
                }

                workbook.IsActive = true;
            };
            

            Application.WorkbookDeactivate += wb =>
            {
                // Could already be gone by the WorkbookBeforeClose event
                if (WorkbookByInteropWorkbook.TryGetValue(wb, out var workbook))
                {
                    WorkbookByInteropWorkbook[wb].IsActive = false;
                }
            };

            Application.WorkbookBeforeClose += WorkbookBeforeClose;
        }

        public InteropApplication Application { get; }

        public IProgram Program { get; }
        public IOfficeCore OfficeCore { get; }

        public IReadOnlyDictionary<InteropWorkbook, Workbook> WorkbookByInteropWorkbook => workbookByInteropWorkbook;

        public IWorkbook[] Workbooks => WorkbookByInteropWorkbook.Values.Cast<IWorkbook>().ToArray();

        public Workbook New(InteropWorkbook interopWorkbook)
        {
            if (!workbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
            {
                workbook = new Workbook(this, interopWorkbook);
                workbookByInteropWorkbook.Add(interopWorkbook, workbook);
            }

            return workbook;
        }

        public void Close(InteropWorkbook interopWorkbook) => workbookByInteropWorkbook.Remove(interopWorkbook);
    }
}
