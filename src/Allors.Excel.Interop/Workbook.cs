// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop
{
    using Allors.Excel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using InteropName = Microsoft.Office.Interop.Excel.Name;
    using InteropXlSheetVisibility = Microsoft.Office.Interop.Excel.XlSheetVisibility;
    using System.Xml;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;

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
                this.AddIn.Application.WorkbookNewSheet -= this.ApplicationOnWorkbookNewSheet;

                if (index.HasValue && index.Value == 0)
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

        public Excel.IWorksheet[] Worksheets => this.worksheetByInteropWorksheet.Values.Cast<IWorksheet>().ToArray();

        public Worksheet[] WorksheetsByIndex => this.worksheetByInteropWorksheet.Values.Cast<Worksheet>().OrderBy(v => v.Index).ToArray();


        public bool IsActive { get; internal set; }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.InteropWorkbook.Close((object)saveChanges ?? Missing.Value, (object)fileName ?? Missing.Value, Missing.Value);
        }

        public Worksheet New(InteropWorksheet interopWorksheet)
        {
            return this.TryAdd(interopWorksheet);
        }
             
        private void ApplicationOnWorkbookNewSheet(InteropWorkbook wb, object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                Worksheet worksheet = this.TryAdd(interopWorksheet);

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

        /// <summary>
        /// Return a Zero-Based Row, Column NamedRanges
        /// </summary>
        /// <returns></returns>
        public Excel.Range[] GetNamedRanges(string refersToSheetName = null)
        {
            var ranges = new List<Excel.Range>();

            foreach (InteropName namedRange in this.InteropWorkbook.Names)
            {
                try
                {
                    var refersToRange = namedRange.RefersToRange;
                    if (refersToRange != null)
                    {                   
                        var iworkSheet = this.worksheetByInteropWorksheet.FirstOrDefault(v => string.Equals(v.Key.Name, refersToRange.Worksheet.Name)).Value;

                        if(string.IsNullOrEmpty(refersToSheetName) || refersToSheetName.Equals(iworkSheet?.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            ranges.Add(new Excel.Range(refersToRange.Row - 1, refersToRange.Column - 1, refersToRange.Rows.Count, refersToRange.Columns.Count, worksheet: iworkSheet, name: namedRange.Name));
                        }
                    }
                }
                catch(Exception ex)
                {
                    // RefersToRange can throw exception
                }
            }

            return ranges.ToArray();
        }


        /// <summary>
        /// Adds a NamedRange that has its scope on the Workbook
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        public void SetNamedRange(string name, Excel.Range range)
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
                }
                catch
                {
                    // can throw exception, we dont care.
                }
            }
        }

        public void SetCustomProperties(Excel.CustomProperties properties)
        {
            if (properties == null || !properties.Any())
            {
                return;
            }

            var customDocumentProperties = this.InteropWorkbook.CustomDocumentProperties;
            Type typeCustomDocumentProperties = customDocumentProperties.GetType();

            foreach (var kvp in properties)
            {
                if (!string.IsNullOrWhiteSpace(kvp.Key))
                {
                    this.TrySet(customDocumentProperties, typeCustomDocumentProperties, kvp.Key, kvp.Value);
                }
            }
        }

        public bool TrySetCustomProperty(string name, dynamic value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            var customDocumentProperties = this.InteropWorkbook.CustomDocumentProperties;
            Type typeCustomDocumentProperties = customDocumentProperties.GetType();

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

        public void DeleteCustomProperties(Excel.CustomProperties properties)
        {
            if (properties == null || !properties.Any())
            {
                return;
            }

            var customDocumentProperties = this.InteropWorkbook.CustomDocumentProperties;
            Type typeCustomDocumentProperties = customDocumentProperties.GetType();

            foreach (var kvp in properties)
            {
                if (!string.IsNullOrWhiteSpace(kvp.Key))
                {
                    try
                    {
                        var nrProps = typeCustomDocumentProperties.InvokeMember("Count",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, customDocumentProperties, new object[] { });

                        for (int counter = 1; counter <= ((int)nrProps); counter++)
                        {
                            var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                                BindingFlags.GetProperty | BindingFlags.Default,
                                null, customDocumentProperties, new object[] { counter });

                            var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                                BindingFlags.GetProperty | BindingFlags.Default,
                                null, itemProp, new object[] { });

                            if (string.Equals(kvp.Key, oPropName))
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

        public Excel.CustomProperties GetCustomProperties()
        {
            var dict = new Excel.CustomProperties();

            object customProperties = this.InteropWorkbook.CustomDocumentProperties;
            Type docPropsType = customProperties.GetType();
            object nrProps;

            nrProps = docPropsType.InvokeMember("Count",
                BindingFlags.GetProperty | BindingFlags.Default,
                null, customProperties, new object[] { });

            for (int counter = 1; counter <= ((int)nrProps); counter++)
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

                if (Excel.CustomProperties.MagicNull.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, null);
                }
                else if (Excel.CustomProperties.MagicDecimalMaxValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, Decimal.MaxValue);
                }
                else if (Excel.CustomProperties.MagicDecimalMinValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, Decimal.MinValue);
                }
                else if (Excel.CustomProperties.MagicDateTimeMaxValue.Equals(oPropVal))
                {
                    dict.Add((string)oPropName, DateTime.MaxValue);
                }
                else if (Excel.CustomProperties.MagicDateTimeMinValue.Equals(oPropVal))
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

        public bool TryGetCustomProperty(string name, ref object value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            try
            {
                object result = null;

                var customDocumentProperties = this.InteropWorkbook.CustomDocumentProperties;
                Type typeCustomDocumentProperties = customDocumentProperties.GetType();

                if (this.TryGet(customDocumentProperties, typeCustomDocumentProperties, name, ref result))
                {
                    if (Excel.CustomProperties.MagicNull.Equals(result))
                    {
                        value = null;
                    }
                    else if (Excel.CustomProperties.MagicDecimalMaxValue.Equals(result))
                    {
                        value = Decimal.MaxValue;
                    }
                    else if (Excel.CustomProperties.MagicDecimalMinValue.Equals(result))
                    {
                        value = Decimal.MinValue;
                    }
                    else if (Excel.CustomProperties.MagicDateTimeMaxValue.Equals(result))
                    {
                        value = DateTime.MaxValue;
                    }
                    else if (Excel.CustomProperties.MagicDateTimeMinValue.Equals(result))
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

                for (int counter = 1; counter <= ((int)nrProps); counter++)
                {
                    var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, customDocumentProperties, new object[] { counter });

                    var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, itemProp, new object[] { });

                    if (string.Equals(key, oPropName))
                    {
                        result = typeCustomDocumentProperties.InvokeMember("Value",
                        BindingFlags.GetProperty | BindingFlags.Default,
                        null, itemProp, new object[] { });

                        return true;
                    }
                }

                return false;
            }
            catch (COMException ex)
            {
                return false;
            }
        }

        private void TrySet(object customDocumentProperties, Type typeCustomDocumentProperties, string key, dynamic value)
        {
            try
            {

                dynamic setValue = value;

                if ( setValue == null)
                {
                    setValue = Excel.CustomProperties.MagicNull;
                }                               

                object result = null;

                if (!this.TryGet(customDocumentProperties, typeCustomDocumentProperties, key, ref result))
                {
                    var type = MsoDocProperties.msoPropertyTypeString;

                    if (value is bool || value is bool?)
                    {
                        type = MsoDocProperties.msoPropertyTypeBoolean;
                    }

                    if (value is DateTime || value is DateTime?)
                    {
                        if ((DateTime)value == DateTime.MaxValue)
                        {
                            setValue = Excel.CustomProperties.MagicDateTimeMaxValue;
                            type = MsoDocProperties.msoPropertyTypeString;
                        }
                        else if ((DateTime)value == DateTime.MinValue)
                        {
                            setValue = Excel.CustomProperties.MagicDateTimeMinValue;
                            type = MsoDocProperties.msoPropertyTypeString;
                        }
                        else
                        {
                           setValue = ((DateTime?)setValue)?.ToOADate();
                           type = MsoDocProperties.msoPropertyTypeDate;
                        }
                    }

                    if (value is float || value is float?)
                    {
                        type = MsoDocProperties.msoPropertyTypeFloat;
                    }

                    if (value is decimal || value is decimal?)
                    {
                        if ((decimal)value == decimal.MaxValue)
                        {
                            setValue = Excel.CustomProperties.MagicDecimalMaxValue;
                            type = MsoDocProperties.msoPropertyTypeString;
                        } 
                        else if ((decimal)value == decimal.MinValue)
                        {
                            setValue = Excel.CustomProperties.MagicDecimalMinValue;
                            type = MsoDocProperties.msoPropertyTypeString;
                        }
                        else
                        {
                            if(decimal.TryParse(Convert.ToString(setValue), out decimal decimalResult))
                            {
                                var parts = decimal.GetBits(decimalResult);
                                byte scale = (byte)((parts[3] >> 16) & 0x7F);

                                if(scale > 6) // float can onlu hold 6 precision
                                {
                                    setValue = Convert.ToSingle(value);
                                }
                            }                            
                            
                            type = MsoDocProperties.msoPropertyTypeFloat;
                        }
                    }

                    if (value is int || value is int?)
                    {
                        type = MsoDocProperties.msoPropertyTypeNumber;
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

                    for (int counter = 1; counter <= ((int)nrProps); counter++)
                    {
                        var itemProp = typeCustomDocumentProperties.InvokeMember("Item",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, customDocumentProperties, new object[] { counter });

                        var oPropName = typeCustomDocumentProperties.InvokeMember("Name",
                            BindingFlags.GetProperty | BindingFlags.Default,
                            null, itemProp, new object[] { });

                        if (string.Equals(key, oPropName))
                        {
                            typeCustomDocumentProperties.InvokeMember("Value",
                            BindingFlags.SetProperty | BindingFlags.Default,
                            null, itemProp, new object[] { setValue });

                            break;
                        }
                    }
                }
            }
            catch (COMException ex)
            {

            }
        }

        public string SetCustomXML(XmlDocument xmlDocument)
        {
            var xmlPart = this.InteropWorkbook.CustomXMLParts.Add(xmlDocument.OuterXml, Type.Missing);

            return xmlPart.Id;
        }

        public XmlDocument GetCustomXMLById(string id)
        {
            var xmlDocument = new XmlDocument();
            var customXMLPart = this.InteropWorkbook.CustomXMLParts.SelectByID(id);

            if(customXMLPart != null)
            {
                xmlDocument.LoadXml(customXMLPart.XML);

                return xmlDocument;
            }

            return null;            
        }

        public bool TryDeleteCustomXMLById(string id)
        {
            try
            {
                var customXMLPart = this.InteropWorkbook.CustomXMLParts.SelectByID(id);

                customXMLPart.Delete();

                return true;
            }
            catch (COMException)
            {
                return false;
            }
           
        }
    }
}
