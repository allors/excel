// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;

namespace Allors.Excel.Interop
{
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropName = Name;
    using InteropXlSheetVisibility = XlSheetVisibility;

    public class Workbook : IWorkbook
    {
        private readonly Dictionary<InteropWorksheet, Worksheet> worksheetByInteropWorksheet;

        public Workbook(AddIn addIn, InteropWorkbook interopWorkbook)
        {
            AddIn = addIn;
            InteropWorkbook = interopWorkbook;
            worksheetByInteropWorksheet = new Dictionary<InteropWorksheet, Worksheet>();
            AddIn.Application.WorkbookNewSheet += ApplicationOnWorkbookNewSheet;
            AddIn.Application.SheetBeforeDelete += ApplicationOnSheetBeforeDelete;
        }

        public AddIn AddIn { get; }

        public InteropWorkbook InteropWorkbook { get; }

        /// <summary>
        /// When index = 0 => add new worksheet before the active worksheet
        /// When index =< #sheets => add new worksheet before the index worksheet
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
            InteropWorksheet interopWorksheet;

            try
            {
                AddIn.Application.WorkbookNewSheet -= ApplicationOnWorkbookNewSheet;

                if (index.HasValue && index.Value == 0)
                {
                    interopWorksheet = (InteropWorksheet)InteropWorkbook.Sheets.Add();
                }
                else
                {
                    if (before != null)
                    {
                        interopWorksheet = (InteropWorksheet)InteropWorkbook.Sheets.Add(((Worksheet)before).InteropWorksheet, Missing.Value);

                    }
                    else if (after != null)
                    {
                        interopWorksheet = (InteropWorksheet)InteropWorkbook.Sheets.Add(Missing.Value, ((Worksheet)after).InteropWorksheet);
                    }
                    else
                    {
                        if (!index.HasValue || index > InteropWorkbook.Sheets.Count - 1)
                        {
                            index = InteropWorkbook.Sheets.Count;
                            var insertAfter = worksheetByInteropWorksheet.Keys.FirstOrDefault(v => v.Index == index);
                            interopWorksheet = (InteropWorksheet)InteropWorkbook.Sheets.Add(Missing.Value, insertAfter);
                        }
                        else
                        {
                            var insertBefore = worksheetByInteropWorksheet.Keys.FirstOrDefault(v => v.Index == index);
                            interopWorksheet = (InteropWorksheet)InteropWorkbook.Sheets.Add(insertBefore);
                        }
                    }
                }

                return TryAdd(interopWorksheet);
            }
            finally
            {
                AddIn.Application.WorkbookNewSheet += ApplicationOnWorkbookNewSheet;
            }
        }

        /// <inheritdoc/>
        public Excel.IWorksheet Copy(Excel.IWorksheet sourceWorksheet, Excel.IWorksheet beforeWorksheet)
        {
            var source = (Worksheet)sourceWorksheet;
            var before = (Worksheet)beforeWorksheet;
            var index = before.InteropWorksheet.Index;

            source.InteropWorksheet.Copy(before.InteropWorksheet);

            var copied = (InteropWorksheet)InteropWorkbook.Sheets[index];
            var copiedWorksheet = New(copied);

            copied.Visible = InteropXlSheetVisibility.xlSheetVisible;

            return copiedWorksheet;
        }

        /// <inheritdoc/>

        public Excel.IWorksheet[] Worksheets => worksheetByInteropWorksheet.Values.Cast<IWorksheet>().ToArray();

        public Worksheet[] WorksheetsByIndex => worksheetByInteropWorksheet.Values.OrderBy(v => v.Index).ToArray();

        /// <inheritdoc/>
        public bool IsActive { get; internal set; }

        /// <inheritdoc/>
        public void Close(bool? saveChanges = null, string fileName = null)
        {
            InteropWorkbook.Close((object)saveChanges ?? Missing.Value, (object)fileName ?? Missing.Value, Missing.Value);
        }

        public Worksheet New(InteropWorksheet interopWorksheet)
        {
            return TryAdd(interopWorksheet);
        }


        private void ApplicationOnWorkbookNewSheet(InteropWorkbook wb, object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                var worksheet = TryAdd(interopWorksheet);

                interopWorksheet.BeforeDelete += async () => await AddIn.Program.OnBeforeDelete(worksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }

        private Worksheet TryAdd(InteropWorksheet interopWorksheet)
        {
            if (!worksheetByInteropWorksheet.TryGetValue(interopWorksheet, out var worksheet))
            {
                worksheet = new Worksheet(this, interopWorksheet);
                worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            }

            worksheet.IsActive = true;
            return worksheet;
        }

        private void ApplicationOnSheetBeforeDelete(object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                worksheetByInteropWorksheet.Remove(interopWorksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }

        /// <inheritdoc/>
        public Range[] GetNamedRanges(string refersToSheetName = null)
        {
            var ranges = new List<Range>();

            foreach (InteropName namedRange in InteropWorkbook.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {
                        var iworkSheet = worksheetByInteropWorksheet.FirstOrDefault(v => string.Equals(v.Key.Name, refersToRange.Worksheet.Name)).Value;

                        if (string.IsNullOrEmpty(refersToSheetName) || refersToSheetName.Equals(iworkSheet?.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            ranges.Add(new Range(refersToRange.Row - 1, refersToRange.Column - 1, refersToRange.Rows.Count, refersToRange.Columns.Count, worksheet: iworkSheet, name: namedRange.Name));
                        }
                    }
                }
                catch (Exception)
                {
                    // RefersToRange can throw exception
                }
            }

            return ranges.ToArray();
        }


        /// <inheritdoc/>
        public void SetNamedRange(string name, Range range)
        {
            if (!string.IsNullOrWhiteSpace(name) && range != null)
            {
                try
                {

                    var interopWorksheet = ((Worksheet)range.Worksheet).InteropWorksheet;

                    if (interopWorksheet != null)
                    {
                        var topLeft = interopWorksheet.Cells[range.Row + 1, range.Column + 1];
                        var bottomRight = interopWorksheet.Cells[range.Row + range.Rows, range.Column + range.Columns];

                        var refersTo = interopWorksheet.Range[topLeft, bottomRight];

                        // When it does not exist, add it, else we update the range.
                        if (InteropWorkbook.Names
                                .Cast<InteropName>()
                                .Any(v => string.Equals(v.Name, name)))
                        {
                            InteropWorkbook.Names.Item(name).RefersTo = refersTo;
                        }
                        else
                        {
                            InteropWorkbook.Names.Add(name, refersTo);
                        }
                    }
                }
                catch
                {
                    // can throw exception, we dont care.
                }
            }
        }

        /// <inheritdoc/>
        public void SetCustomProperties(CustomProperties properties)
        {
            if (properties == null || !properties.Any())
            {
                return;
            }

            var customDocumentProperties = InteropWorkbook.CustomDocumentProperties;
            var typeCustomDocumentProperties = customDocumentProperties.GetType();

            foreach (var kvp in properties)
            {
                if (!string.IsNullOrWhiteSpace(kvp.Key))
                {
                    TrySet(customDocumentProperties, typeCustomDocumentProperties, kvp.Key, kvp.Value);
                }
            }
        }

        /// <inheritdoc/>
        public bool TrySetCustomProperty(string name, dynamic value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            var customDocumentProperties = InteropWorkbook.CustomDocumentProperties;
            var typeCustomDocumentProperties = customDocumentProperties.GetType();

            try
            {
                this.TrySet(customDocumentProperties, typeCustomDocumentProperties, name, value);

                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        /// <inheritdoc/>
        public void DeleteCustomProperties(CustomProperties properties)
        {
            if (properties == null || !properties.Any())
            {
                return;
            }

            var customDocumentProperties = InteropWorkbook.CustomDocumentProperties;
            var typeCustomDocumentProperties = customDocumentProperties.GetType();

            foreach (var kvp in properties)
            {
                if (!string.IsNullOrWhiteSpace(kvp.Key))
                {
                    try
                    {
                        var nrProps = typeCustomDocumentProperties.InvokeMember("Count",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, customDocumentProperties, new object[] { });

                        for (var counter = 1; counter <= ((int)nrProps); counter++)
                        {
                            var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                                BindingFlags.GetProperty | BindingFlags.Default,
                                null, customDocumentProperties, new object[] { counter });

                            var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                                BindingFlags.GetProperty | BindingFlags.Default,
                                null, itemProp, new object[] { });

                            if (Equals(kvp.Key, oPropName))
                            {
                                typeCustomDocumentProperties.InvokeMember("Delete",
                                 BindingFlags.InvokeMethod | BindingFlags.Default,
                                 null, itemProp, new object[] { });

                                break;
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // Blank
                    }

                }
            }
        }

        /// <inheritdoc/>
        public CustomProperties GetCustomProperties()
        {
            var dict = new CustomProperties();

            var customProperties = InteropWorkbook.CustomDocumentProperties;
            var docPropsType = customProperties.GetType();
            object nrProps;

            nrProps = docPropsType.InvokeMember("Count",
                BindingFlags.GetProperty | BindingFlags.Default,
                null, customProperties, new object[] { });

            for (var counter = 1; counter <= ((int)nrProps); counter++)
            {
                var itemProp = docPropsType.InvokeMember("Item",
                    BindingFlags.GetProperty | BindingFlags.Default,
                    null, customProperties, new object[] { counter });

                var oPropName = docPropsType.InvokeMember("Name",
                    BindingFlags.GetProperty | BindingFlags.Default,
                    null, itemProp, new object[] { });

                var oPropVal = docPropsType.InvokeMember("Value",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, itemProp, new object[] { });

                if (CustomProperties.MagicNull.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, null);
                }
                else if (CustomProperties.MagicDecimalMaxValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, Decimal.MaxValue);
                }
                else if (CustomProperties.MagicDecimalMinValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, Decimal.MinValue);
                }
                else if (CustomProperties.MagicDateTimeMaxValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, DateTime.MaxValue);
                }
                else if (CustomProperties.MagicDateTimeMinValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, DateTime.MinValue);
                }
                else
                {
                    dict.Add((string)oPropName, oPropVal);
                }
            }

            return dict;
        }

        /// <inheritdoc/>
        public bool TryGetCustomProperty(string name, ref object value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            try
            {
                object result = null;

                var customDocumentProperties = InteropWorkbook.CustomDocumentProperties;
                var typeCustomDocumentProperties = customDocumentProperties.GetType();

                if (TryGet(customDocumentProperties, typeCustomDocumentProperties, name, ref result))
                {
                    if (CustomProperties.MagicNull.Equals(result))
                    {
                        value = null;
                    }
                    else if (CustomProperties.MagicDecimalMaxValue.Equals(result))
                    {
                        value = Decimal.MaxValue;
                    }
                    else if (CustomProperties.MagicDecimalMinValue.Equals(result))
                    {
                        value = Decimal.MinValue;
                    }
                    else if (CustomProperties.MagicDateTimeMaxValue.Equals(result))
                    {
                        value = DateTime.MaxValue;
                    }
                    else if (CustomProperties.MagicDateTimeMinValue.Equals(result))
                    {
                        value = DateTime.MinValue;
                    }
                    else
                    {
                        value = result;
                    }

                    return true;
                }

                return false;
            }
            catch (COMException)
            {
                return false;
            }
        }


        private bool TryGet(object customDocumentProperties, Type typeCustomDocumentProperties, string key, ref dynamic result)
        {
            try
            {

                var nrProps = typeCustomDocumentProperties.InvokeMember("Count",
                    BindingFlags.GetProperty | BindingFlags.Default,
                    null, customDocumentProperties, new object[] { });

                for (var counter = 1; counter <= ((int)nrProps); counter++)
                {
                    var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, customDocumentProperties, new object[] { counter });

                    var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, itemProp, new object[] { });

                    if (Equals(key, oPropName))
                    {
                        result = typeCustomDocumentProperties.InvokeMember("Value",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, itemProp, new object[] { });

                        return true;
                    }
                }

                return false;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private void TrySet(object customDocumentProperties, Type typeCustomDocumentProperties, string key, dynamic value)
        {
            try
            {

                var setValue = value;

                if (setValue == null)
                {
                    setValue = CustomProperties.MagicNull;
                }

                object result = null;

                if (!TryGet(customDocumentProperties, typeCustomDocumentProperties, key, ref result))
                {
                    var office = AddIn.OfficeCore;

                    var type = office.MsoPropertyTypeString;

                    if (value is bool || value is bool?)
                    {
                        type = office.MsoPropertyTypeBoolean;
                    }

                    if (value is DateTime || value is DateTime?)
                    {
                        if ((DateTime)value == DateTime.MaxValue)
                        {
                            setValue = CustomProperties.MagicDateTimeMaxValue;
                            type = office.MsoPropertyTypeString;
                        }
                        else if ((DateTime)value == DateTime.MinValue)
                        {
                            setValue = CustomProperties.MagicDateTimeMinValue;
                            type = office.MsoPropertyTypeString;
                        }
                        else
                        {
                            setValue = ((DateTime?)setValue)?.ToOADate();
                            type = office.MsoPropertyTypeDate;
                        }
                    }

                    if (value is float || value is float?)
                    {
                        type = office.MsoPropertyTypeFloat;
                    }

                    if (value is decimal || value is decimal?)
                    {
                        if ((decimal)value == decimal.MaxValue)
                        {
                            setValue = CustomProperties.MagicDecimalMaxValue;
                            type = office.MsoPropertyTypeString;
                        }
                        else if ((decimal)value == decimal.MinValue)
                        {
                            setValue = CustomProperties.MagicDecimalMinValue;
                            type = office.MsoPropertyTypeString;
                        }
                        else
                        {
                            if (decimal.TryParse(Convert.ToString(setValue), out decimal decimalResult))
                            {
                                var parts = decimal.GetBits(decimalResult);
                                var scale = (byte)((parts[3] >> 16) & 0x7F);

                                if (scale > 6) // float can onlu hold 6 precision
                                {
                                    setValue = Convert.ToSingle(value);
                                }
                            }

                            type = office.MsoPropertyTypeFloat;
                        }
                    }

                    if (value is int || value is int?)
                    {
                        type = office.MsoPropertyTypeNumber;
                    }

                    object[] oArgs = { key, false, type, setValue };

                    typeCustomDocumentProperties.InvokeMember("Add", BindingFlags.Default |
                                               BindingFlags.InvokeMethod, null,
                                               customDocumentProperties, oArgs);
                }
                else
                {
                    var nrProps = typeCustomDocumentProperties.InvokeMember("Count",
                    BindingFlags.GetProperty | BindingFlags.Default,
                    null, customDocumentProperties, new object[] { });

                    for (var counter = 1; counter <= ((int)nrProps); counter++)
                    {
                        var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, customDocumentProperties, new object[] { counter });

                        var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, itemProp, new object[] { });

                        if (Equals(key, oPropName))
                        {
                            typeCustomDocumentProperties.InvokeMember("Value",
                            BindingFlags.SetProperty | BindingFlags.Default,
                            null, itemProp, new object[] { setValue });

                            break;
                        }
                    }
                }
            }
            catch (COMException)
            {

            }
        }

        /// <inheritdoc/>
        public XmlDocument GetCustomXMLById(string id) => AddIn.OfficeCore.GetCustomXmlById(InteropWorkbook, id);

        /// <inheritdoc/>
        public string SetCustomXML(XmlDocument xmlDocument) => AddIn.OfficeCore.SetCustomXmlPart(InteropWorkbook, xmlDocument);

        /// <inheritdoc/>
        public bool TryDeleteCustomXMLById(string id) => AddIn.OfficeCore.TryDeleteCustomXmlById(InteropWorkbook, id);
    }
}
