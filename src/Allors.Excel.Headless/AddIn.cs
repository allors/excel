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
        public AddIn(IProgram program, IRibbon ribbon)
        {
            this.Program = program;
            this.Ribbon = ribbon;
            this.WorkbookList = new List<Workbook>();

            this.AddWorkbook();

            this.Program.OnStart(this).ConfigureAwait(false);
        }

        public IProgram Program { get; }

        public IRibbon Ribbon { get; set; }

        public IWorkbook[] Workbooks => this.WorkbookList.Cast<IWorkbook>().ToArray();

        public IList<Workbook> WorkbookList { get; }

        public string ExistentialAttribute { get; set; }

        public Workbook AddWorkbook()
        {
            var workbook = new Workbook(this);
            this.WorkbookList.Add(workbook);
            workbook.Activate();

            workbook.AddWorksheet();

            this.Program.OnNew(workbook).ConfigureAwait(false);

            return workbook;
        }

        public void DisplayAlerts(bool displayAlerts) => throw new System.NotImplementedException();

        public void Remove(Workbook workbook) => this.WorkbookList.Remove(workbook);
    }
}
