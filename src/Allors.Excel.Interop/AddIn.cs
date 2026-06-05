// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System.Collections.Generic;
    using System.Linq;
    using InteropAppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
    using InteropApplication = Microsoft.Office.Interop.Excel.Application;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>
    /// The interop (VSTO) <see cref="IAddIn"/> implementation, wrapping the live Excel object model.
    /// </summary>
    /// <remarks>
    /// COM lifetime policy: this interop layer deliberately does NOT call
    /// <c>Marshal.ReleaseComObject</c>. Runtime Callable Wrappers are released by the CLR's
    /// finalizer/GC, which is sufficient for the in-process VSTO add-in scenario (Excel hosts the
    /// process, so every wrapper is freed when the add-in unloads). Explicit release is avoided on
    /// purpose: the virtual DOM retains some COM objects (e.g. the worksheet and workbook) for their
    /// lifetime, and the CLR caches a single wrapper per COM identity, so a stray release would risk
    /// an <c>InvalidComObjectException</c> on a still-referenced object. Callers that drive Excel
    /// out-of-process and need prompt teardown should force collection (<c>GC.Collect()</c> +
    /// <c>GC.WaitForPendingFinalizers()</c>) or quit/kill the instance, as the interop test fixture does.
    /// </remarks>
    public class AddIn : IAddIn
    {
        private readonly Dictionary<InteropWorkbook, Workbook> workbookByInteropWorkbook;

        public AddIn(InteropApplication application, IProgram program, IRibbon ribbon)
        {
            this.Application = application;
            this.Program = program;
            this.Ribbon = ribbon;

            this.workbookByInteropWorkbook = [];

            ((InteropAppEvents_Event)this.Application).NewWorkbook += async interopWorkbook =>
            {
                if (!string.IsNullOrWhiteSpace(this.ExistentialAttribute))
                {
                    var customProperties = new CustomProperties(interopWorkbook.CustomDocumentProperties);
                    if (this.ExistentialAttribute == null || !customProperties.Exist(this.ExistentialAttribute))
                    {
                        return;
                    }
                }

                var workbook = this.New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                // Notify the existing worksheets first; worksheets added during
                // OnNew(workbook) are notified by Workbook.AddWorksheet itself.
                var worksheets = workbook.Worksheets;
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }

                await this.Program.OnNew(workbook);
            };

            this.Application.WorkbookOpen += async interopWorkbook =>
            {
                if (!string.IsNullOrWhiteSpace(this.ExistentialAttribute))
                {
                    var customProperties = new CustomProperties(interopWorkbook.CustomDocumentProperties);
                    if (this.ExistentialAttribute == null || !customProperties.Exist(this.ExistentialAttribute))
                    {
                        return;
                    }
                }

                var workbook = this.New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                // Notify the existing worksheets first; worksheets added during
                // OnNew(workbook) are notified by Workbook.AddWorksheet itself.
                var worksheets = workbook.Worksheets;
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }

                await this.Program.OnNew(workbook);
            };

            this.Application.WorkbookActivate += interopWorkbook =>
            {
                if (!string.IsNullOrWhiteSpace(this.ExistentialAttribute))
                {
                    var customProperties = new CustomProperties(interopWorkbook.CustomDocumentProperties);
                    if (this.ExistentialAttribute == null || !customProperties.Exist(this.ExistentialAttribute))
                    {
                        return;
                    }
                }

                if (!this.WorkbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
                {
                    workbook = this.New(interopWorkbook);
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

            this.Application.WorkbookBeforeClose += WorkbookBeforeClose;
        }

        public string ExistentialAttribute { get; set; }

        public InteropApplication Application { get; }

        public IProgram Program { get; }

        public IReadOnlyDictionary<InteropWorkbook, Workbook> WorkbookByInteropWorkbook => this.workbookByInteropWorkbook;

        public IRibbon Ribbon { get; }

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

        public void Close(InteropWorkbook interopWorkbook)
        {
            if (this.workbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
            {
                // Detach the workbook's Application-level event handlers before dropping it,
                // otherwise the closed workbook leaks via those still-attached handlers.
                workbook.Disconnect();
            }

            this.workbookByInteropWorkbook.Remove(interopWorkbook);
        }

        public void DisplayAlerts(bool displayAlerts) => this.Application.DisplayAlerts = displayAlerts;
    }
}
