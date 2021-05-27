// <copyright file="AddIn.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;
using System.Linq;

namespace Allors.Excel.Headless
{
    public class AddIn : IAddIn
    {
        public AddIn()
        {
            WorkbookList = new List<Workbook>();
        }

        public IWorkbook[] Workbooks => WorkbookList.Cast<IWorkbook>().ToArray();

        public IList<Workbook> WorkbookList { get; }

        public Workbook AddWorkbook()
        {
            var workbook = new Workbook(this);
            WorkbookList.Add(workbook);
            workbook.Activate();
            return workbook;
        }

        public void Remove(Workbook workbook)
        {
            WorkbookList.Remove(workbook);
        }
    }
}
