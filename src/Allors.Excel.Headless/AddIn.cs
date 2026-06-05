// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    public class AddIn : IAddIn
    {
        // Private: construction creates the initial workbook and notifies the program, which
        // is asynchronous. Use CreateAsync so OnNew/OnStart are awaited rather than fire-and-forget
        // (a constructor cannot await).
        private AddIn(IProgram program, IRibbon ribbon)
        {
            this.Program = program;
            this.Ribbon = ribbon;
            this.WorkbookList = new List<Workbook>();
        }

        public static async Task<AddIn> CreateAsync(IProgram program, IRibbon ribbon)
        {
            var addIn = new AddIn(program, ribbon);

            await addIn.AddWorkbook();
            await program.OnStart(addIn);

            return addIn;
        }

        public IProgram Program { get; }

        public IRibbon Ribbon { get; set; }

        public IWorkbook[] Workbooks => this.WorkbookList.Cast<IWorkbook>().ToArray();

        public IList<Workbook> WorkbookList { get; }

        public string ExistentialAttribute { get; set; }

        public async Task<Workbook> AddWorkbook()
        {
            var workbook = new Workbook(this);
            this.WorkbookList.Add(workbook);
            workbook.Activate();

            await workbook.AddWorksheet();

            await this.Program.OnNew(workbook);

            return workbook;
        }

        public void DisplayAlerts(bool displayAlerts) => throw new System.NotImplementedException();

        public void Remove(Workbook workbook) => this.WorkbookList.Remove(workbook);
    }
}
