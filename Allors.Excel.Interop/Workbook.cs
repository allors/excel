// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Embedded
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public class Workbook : IWorkbook
    {
        private readonly Dictionary<InteropWorksheet, Worksheet> worksheetByInteropWorksheet;

        public Workbook(AddIn addIn, InteropWorkbook interopWorkbook)
        {
            this.AddIn = addIn;
            this.InteropWorkbook = interopWorkbook;
            this.worksheetByInteropWorksheet = new Dictionary<InteropWorksheet, Worksheet>();
            this.AddIn.Application.WorkbookNewSheet += this.ApplicationOnWorkbookNewSheet;
            this.AddIn.Application.SheetBeforeDelete += this.ApplicationOnSheetBeforeDelete;
        }

        public AddIn AddIn { get; }

        public InteropWorkbook InteropWorkbook { get; }

        /// <summary>
        /// Return a Zero-Based Row, Column NamedRanges
        /// </summary>
        /// <returns></returns>
        public Range[] GetNamedRanges()
        {
            var ranges = new List<Range>();

            foreach (Microsoft.Office.Interop.Excel.Name namedRange in this.InteropWorkbook.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {
                        ranges.Add(new Range(refersToRange.Row - 1,refersToRange.Column - 1,refersToRange.Rows.Count,refersToRange.Columns.Count, name: namedRange.Name));
                    }
                }
                catch 
                { 
                    // RefersToRange can throw exception
                }
            }

            return ranges.ToArray();
        }

        public IWorksheet AddWorksheet(int? index, IWorksheet before = null, IWorksheet after = null)
        {
            InteropWorksheet interopWorksheet;

            if (index.HasValue && index.Value == 0)
            {
                interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add();
            }
            else
            {
                if (before != null)
                {
                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(((Worksheet)before).InteropWorksheet);
                }
                else if (after != null)
                {
                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(null, ((Worksheet)after).InteropWorksheet);
                }
                else
                {
                    var sortedWorksheets = this.worksheetByInteropWorksheet.OrderBy(v => v.Value.Index).Select(v => v.Key).ToArray();
                    InteropWorksheet append = null;
                    if (sortedWorksheets.Any())
                    {
                        if (!index.HasValue || index > sortedWorksheets.Length - 1)
                        {
                            index = sortedWorksheets.Length - 1;
                        }

                        append = sortedWorksheets[index.Value];
                    }

                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(Missing.Value, append, Missing.Value, Missing.Value);
                }
            }

            return this.worksheetByInteropWorksheet[interopWorksheet];
        }

        public IWorksheet Copy(IWorksheet sourceWorksheet, IWorksheet beforeWorksheet)
        {
            var source = (Worksheet)sourceWorksheet;
            var before = (Worksheet)beforeWorksheet;
            var index = before.InteropWorksheet.Index;

            source.InteropWorksheet.Copy(before.InteropWorksheet);

            var copied = (InteropWorksheet)this.InteropWorkbook.Sheets[index];
            var copiedWorksheet = this.New(copied);

            copied.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;

            return copiedWorksheet;
        }

        public IWorksheet[] Worksheets => this.worksheetByInteropWorksheet.Values.Cast<IWorksheet>().ToArray();

        public bool IsActive { get; internal set; }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.InteropWorkbook.Close((object)saveChanges ?? Missing.Value, (object)fileName ?? Missing.Value, Missing.Value);
        }

        public Worksheet New(InteropWorksheet interopWorksheet)
        {
            var worksheet = new Worksheet(this, interopWorksheet);
            this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            
            return worksheet;
        }

        private void Workbook_Activate()
        {
            throw new NotImplementedException();
        }

        private async void ApplicationOnWorkbookNewSheet(InteropWorkbook wb, object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                if (!this.worksheetByInteropWorksheet.TryGetValue(interopWorksheet, out var worksheet))
                {
                    worksheet = new Worksheet(this, interopWorksheet);
                    this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
                }

                interopWorksheet.BeforeDelete += async () => await this.AddIn.Program.OnBeforeDelete(worksheet);

                await this.AddIn.Program.OnNew(worksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }

        private void ApplicationOnSheetBeforeDelete(object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                this.worksheetByInteropWorksheet.Remove(interopWorksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }
    }
}
