// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;
using System.Linq;

namespace Allors.Excel.Headless
{
    public class Workbook : IWorkbook
    {
        public Workbook(AddIn addIn)
        {
            this.AddIn = addIn;
            this.WorksheetList = new List<Worksheet>();
        }

        public AddIn AddIn { get; }

        public List<Worksheet> WorksheetList { get; set; }

        public IWorksheet[] Worksheets => this.WorksheetList.Cast<IWorksheet>().ToArray();

        public bool IsActive { get; private set; }

        public List<Range> NamedRanges { get; } = new List<Range>();

        public IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null)
        {
            var worksheet = new Worksheet(this);

            if (index != null)
            {
                this.WorksheetList.Insert(index.Value, worksheet);
            }
            else if (before != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(before as Worksheet), worksheet);
            }
            else if (after != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(after as Worksheet) + 1, worksheet);
            }
            else
            {
                this.WorksheetList.Add(worksheet);
            }

            worksheet.Activate();

            return worksheet;
        }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.AddIn.Remove(this);
        }

        public void Activate()
        {
            foreach (var workbook in this.AddIn.WorkbookList)
            {
                workbook.IsActive = false;
            }

            this.IsActive = true;
        }

        public Excel.Range[] GetNamedRanges(string refersToSheetName = null)
        {
            return this.NamedRanges.ToArray();
        }

        public IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet)
        {
            throw new System.NotImplementedException();
        }

        public void SetNamedRange(string name, Range range)
        {
            throw new System.NotImplementedException();
        }
    }
}
