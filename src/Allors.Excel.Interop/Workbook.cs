// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Xml;
    using InteropName = Microsoft.Office.Interop.Excel.Name;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropXlSheetVisibility = Microsoft.Office.Interop.Excel.XlSheetVisibility;

    public class Workbook : IWorkbook
    {
        private readonly Dictionary<InteropWorksheet, Worksheet> worksheetByInteropWorksheet;

        public Workbook(AddIn addIn, InteropWorkbook interopWorkbook)
        {
            this.AddIn = addIn;
            this.InteropWorkbook = interopWorkbook;
            this.worksheetByInteropWorksheet = [];
            this.AddIn.Application.WorkbookNewSheet += this.ApplicationOnWorkbookNewSheet;
            this.AddIn.Application.SheetBeforeDelete += this.ApplicationOnSheetBeforeDelete;
            this.BuiltinProperties = new BuiltinProperties(this.InteropWorkbook.BuiltinDocumentProperties);
            this.CustomProperties = new CustomProperties(this.InteropWorkbook.CustomDocumentProperties);
        }

        public event EventHandler<Hyperlink> OnHyperlinkClicked;

        public AddIn AddIn { get; }

        public InteropWorkbook InteropWorkbook { get; }

        public IBuiltinProperties BuiltinProperties { get; }

        public ICustomProperties CustomProperties { get; }

        /// <inheritdoc/>
        public Excel.IWorksheet[] Worksheets => this.worksheetByInteropWorksheet.Values.Cast<Excel.IWorksheet>().ToArray();

        public Excel.IWorksheet[] WorksheetsByIndex => this.worksheetByInteropWorksheet.Values.OrderBy(v => v.Index).Cast<Excel.IWorksheet>().ToArray();

        /// <inheritdoc/>
        public bool IsActive { get; internal set; }

        /// <summary>
        /// When index = 0 => add new worksheet before the active worksheet
        /// When index &lt;= #sheets >= add new worksheet before the index worksheet
        /// When index > #sheets => add new worksheet after the last worksheet
        /// When before != null  => add new worksheet before that worksheet
        /// When after != null  => add new worksheet after that worksheet
        /// When all params are null => add new worksheet after the last worksheet
        /// </summary>
        /// <param name="index"></param>
        /// <param name="before"></param>
        /// <param name="after"></param>
        /// <returns></returns>
        public Excel.IWorksheet AddWorksheet(int? index, Excel.IWorksheet before = null, Excel.IWorksheet after = null)
        {
            try
            {
                this.AddIn.Application.WorkbookNewSheet -= this.ApplicationOnWorkbookNewSheet;

                InteropWorksheet interopWorksheet;
                if (index is 0)
                {
                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add();
                }
                else
                {
                    if (before != null)
                    {
                        interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(((Worksheet)before).InteropWorksheet, Missing.Value);
                    }
                    else if (after != null)
                    {
                        interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(Missing.Value, ((Worksheet)after).InteropWorksheet);
                    }
                    else
                    {
                        if (!index.HasValue || index > this.InteropWorkbook.Sheets.Count - 1)
                        {
                            index = this.InteropWorkbook.Sheets.Count;
                            var insertAfter = this.worksheetByInteropWorksheet.Keys.FirstOrDefault(v => v.Index == index);
                            interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(Missing.Value, insertAfter);
                        }
                        else
                        {
                            var insertBefore = this.worksheetByInteropWorksheet.Keys.FirstOrDefault(v => v.Index == index);
                            interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(insertBefore);
                        }
                    }
                }

                return this.TryAdd(interopWorksheet);
            }
            finally
            {
                this.AddIn.Application.WorkbookNewSheet += this.ApplicationOnWorkbookNewSheet;
            }
        }

        /// <inheritdoc/>
        public Excel.IWorksheet Copy(Excel.IWorksheet sourceWorksheet, Excel.IWorksheet beforeWorksheet)
        {
            var source = (Worksheet)sourceWorksheet;
            var before = (Worksheet)beforeWorksheet;
            var index = before.InteropWorksheet.Index;

            source.InteropWorksheet.Copy(before.InteropWorksheet);

            var copied = (InteropWorksheet)this.InteropWorkbook.Sheets[index];
            var copiedWorksheet = this.New(copied);

            copied.Visible = InteropXlSheetVisibility.xlSheetVisible;

            return copiedWorksheet;
        }

        /// <inheritdoc/>
        public void Close(bool? saveChanges = null, string fileName = null) => this.InteropWorkbook.Close((object)saveChanges ?? Missing.Value, (object)fileName ?? Missing.Value, Missing.Value);

        public Worksheet New(InteropWorksheet interopWorksheet) => this.TryAdd(interopWorksheet);

        /// <inheritdoc/>
        public Range[] GetNamedRanges(string refersToSheetName = null)
        {
            var ranges = new List<Range>();

            foreach (InteropName namedRange in this.InteropWorkbook.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {
                        var iworkSheet = this.worksheetByInteropWorksheet.FirstOrDefault(v => string.Equals(v.Key.Name, refersToRange.Worksheet.Name)).Value;

                        if (string.IsNullOrEmpty(refersToSheetName) || refersToSheetName!.Equals(iworkSheet?.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            ranges.Add(new Range(refersToRange.Row - 1, refersToRange.Column - 1, refersToRange.Rows.Count, refersToRange.Columns.Count, worksheet: iworkSheet, name: namedRange.Name));
                        }
                    }
                }
                catch
                {
                    // RefersToRange can throw exception
                }
            }

            return ranges.ToArray();
        }

        /// <inheritdoc/>
        public void SetNamedRange(string name, Range range)
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    var interopWorksheet = ((Worksheet)range.Worksheet!).InteropWorksheet;

                    var topLeft = interopWorksheet.Cells[range.Row + 1, range.Column + 1];
                    var bottomRight = interopWorksheet.Cells[range.Row + range.Rows, range.Column + range.Columns];

                    var refersTo = interopWorksheet.Range[topLeft, bottomRight];

                    // When it does not exist, add it, else we update the range.
                    if (this.InteropWorkbook.Names
                            .Cast<InteropName>()
                            .Any(v => string.Equals(v.Name, name)))
                    {
                        this.InteropWorkbook.Names.Item(name).RefersTo = refersTo;
                    }
                    else
                    {
                        this.InteropWorkbook.Names.Add(name, refersTo);
                    }
                }
                catch
                {
                    // can throw exception, we don't care.
                }
            }
        }

        /// <inheritdoc/>
        public XmlDocument GetCustomXmlById(string id)
        {
            var customXmlPart = this.InteropWorkbook.CustomXMLParts.SelectByID(id);

            if (customXmlPart != null)
            {
                var xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(customXmlPart.XML);
                return xmlDocument;
            }

            return null;
        }

        /// <inheritdoc/>
        public string SetCustomXml(XmlDocument xmlDocument) => this.InteropWorkbook.CustomXMLParts.Add(xmlDocument.OuterXml, Type.Missing).Id;

        /// <inheritdoc/>
        public bool TryDeleteCustomXmlById(string id)
        {
            try
            {
                var customXmlPart = this.InteropWorkbook.CustomXMLParts.SelectByID(id);
                customXmlPart.Delete();
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public void HyperlinkClicked(Hyperlink hyperlink) => this.OnHyperlinkClicked?.Invoke(this, hyperlink);

        private void ApplicationOnWorkbookNewSheet(InteropWorkbook wb, object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                var worksheet = this.TryAdd(interopWorksheet);

                interopWorksheet.BeforeDelete += async () => await this.AddIn.Program.OnBeforeDelete(worksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }

        private Worksheet TryAdd(InteropWorksheet interopWorksheet)
        {
            if (!this.worksheetByInteropWorksheet.TryGetValue(interopWorksheet, out var worksheet))
            {
                worksheet = new Worksheet(this, interopWorksheet);
                this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            }

            worksheet.IsActive = true;
            return worksheet;
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
