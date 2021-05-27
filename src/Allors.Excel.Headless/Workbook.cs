// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace Allors.Excel.Headless
{
    public class Workbook : IWorkbook
    {
        public Workbook(AddIn addIn)
        {
            AddIn = addIn;
            WorksheetList = new List<Worksheet>();
        }

        public AddIn AddIn { get; }

        public List<Worksheet> WorksheetList { get; set; }

        public IWorksheet[] Worksheets => WorksheetList.Cast<IWorksheet>().ToArray();

        public bool IsActive { get; private set; }

        public Dictionary<string, Range> NamedRangeByName { get; } = new Dictionary<string, Range>();

        public IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null)
        {
            var worksheet = new Worksheet(this);

            if (index != null)
            {
                WorksheetList.Insert(index.Value, worksheet);
            }
            else if (before != null)
            {
                WorksheetList.Insert(WorksheetList.IndexOf(before as Worksheet), worksheet);
            }
            else if (after != null)
            {
                WorksheetList.Insert(WorksheetList.IndexOf(after as Worksheet) + 1, worksheet);
            }
            else
            {
                WorksheetList.Add(worksheet);
            }

            worksheet.Activate();

            return worksheet;
        }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            AddIn.Remove(this);
        }

        public void Activate()
        {
            foreach (var workbook in AddIn.WorkbookList)
            {
                workbook.IsActive = false;
            }

            IsActive = true;
        }

        public Range[] GetNamedRanges(string refersToSheetName = null)
        {
            return NamedRangeByName.Values.ToArray();
        }

        public IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet)
        {
            throw new NotImplementedException();
        }

        public void SetNamedRange(string name, Range range)
        {
            NamedRangeByName[name] = range;
        }

        public void SetCustomProperties(CustomProperties properties)
        {
            throw new NotImplementedException();
        }

        public void DeleteCustomProperties(CustomProperties properties)
        {
            throw new NotImplementedException();
        }

        public CustomProperties GetCustomProperties()
        {
            throw new NotImplementedException();
        }

        public bool TryGetCustomProperty(string name, ref object value)
        {
            throw new NotImplementedException();
        }

        public string SetCustomXML(XmlDocument xmlDocument)
        {
            throw new NotImplementedException();
        }

        public XmlDocument GetCustomXMLById(string id)
        {
            throw new NotImplementedException();
        }

        public bool TrySetCustomProperty(string name, dynamic value)
        {
            throw new NotImplementedException();
        }

        public bool TryDeleteCustomXMLById(string id)
        {
            throw new NotImplementedException();
        }
    }
}
