// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;

    public class Workbook : IWorkbook
    {
        private int counter;

        public Workbook(AddIn addIn)
        {
            this.AddIn = addIn;
            this.WorksheetList = new List<Worksheet>();
            this.BuiltinProperties = new BuiltinProperties();
            this.CustomProperties = new CustomProperties();
        }

        public AddIn AddIn { get; }

        public List<Worksheet> WorksheetList { get; set; }

        public IWorksheet[] Worksheets => this.WorksheetList.Cast<IWorksheet>().ToArray();

        public bool IsActive { get; private set; }

        public Dictionary<string, Range> NamedRangeByName { get; } = new Dictionary<string, Range>();

        public event EventHandler<Allors.Excel.Hyperlink> OnHyperlinkClicked;

        public IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null)
        {
            var worksheet = new Worksheet(this)
            {
                Name = $"Sheet{++this.counter}"
            };

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

            this.AddIn.Program.OnNew(worksheet).ConfigureAwait(false);

            return worksheet;
        }

        public void Close(bool? saveChanges = null, string fileName = null) => this.AddIn.Remove(this);

        public void Activate()
        {
            foreach (var workbook in this.AddIn.WorkbookList)
            {
                workbook.IsActive = false;
            }

            this.IsActive = true;
        }

        public Range[] GetNamedRanges(string refersToSheetName = null) => this.NamedRangeByName.Values.ToArray();

        public IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet) => throw new NotImplementedException();

        public void SetNamedRange(string name, Range range) => this.NamedRangeByName[name] = range;

        public IBuiltinProperties BuiltinProperties { get; }

        public ICustomProperties CustomProperties { get; }

        public IWorksheet[] WorksheetsByIndex => this.Worksheets;

        public string SetCustomXML(XmlDocument xmlDocument) => throw new NotImplementedException();

        public XmlDocument GetCustomXMLById(string id) => throw new NotImplementedException();

        public bool TryDeleteCustomXMLById(string id) => throw new NotImplementedException();

        public bool TrySetCustomProperty(string name, dynamic value) => throw new NotImplementedException();

        public void HyperlinkClicked(Allors.Excel.Hyperlink hyperlink) => throw new NotImplementedException();
    }
}
