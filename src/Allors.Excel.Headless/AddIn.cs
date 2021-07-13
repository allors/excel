// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System.Collections.Generic;
    using System.Linq;

    public class AddIn : IAddIn
    {
        public AddIn() => this.WorkbookList = new List<Workbook>();

        public IRibbon Ribbon { get; set; }

        public IWorkbook[] Workbooks => this.WorkbookList.Cast<IWorkbook>().ToArray();

        public IList<Workbook> WorkbookList { get; }

        public Workbook AddWorkbook()
        {
            var workbook = new Workbook(this);
            this.WorkbookList.Add(workbook);
            workbook.Activate();
            return workbook;
        }

        public void DisplayAlerts(bool displayAlerts) => throw new System.NotImplementedException();

        public void Remove(Workbook workbook) => this.WorkbookList.Remove(workbook);
    }
}
