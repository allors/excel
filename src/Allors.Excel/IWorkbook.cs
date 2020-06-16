// <copyright file="IWorkbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;
using System.Xml;

namespace Allors.Excel
{
    public interface IWorkbook
    {
        /// <summary>
        /// Gets the IsActive value.
        /// </summary>
        bool IsActive { get; }

        /// <summary>
        /// Gets the List of worksheets in this workbook
        /// </summary>
        IWorksheet[] Worksheets { get; }

        /// <summary>
        /// Closed the workbook with the given parameters
        /// </summary>
        /// <param name="saveChanges"></param>
        /// <param name="fileName"></param>
        void Close(bool? saveChanges = null, string fileName = null);

        /// <summary>
        /// Add a worksheet at the given location
        /// </summary>
        /// <param name="index"></param>
        /// <param name="before"></param>
        /// <param name="after"></param>
        /// <returns></returns>
        IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null);

        /// <summary>
        /// Copies a worksheet at the given beforeWorksheet
        /// </summary>
        /// <param name="source"></param>
        /// <param name="beforeWorksheet"></param>
        /// <returns></returns>
        IWorksheet Copy(IWorksheet source, IWorksheet beforeWorksheet);

        /// <summary>
        /// Gets all named Ranges in this workbook scope.
        /// </summary>
        /// <returns></returns>
        Range[] GetNamedRanges(string refersToSheetName = null);

        /// <summary>
        /// Adds a NamedRange scoped to the Workbook
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        void SetNamedRange(string name, Range range);

        /// <summary>
        /// Adds or updates the (interop) CustomDocumentProperties in this workbook
        /// </summary>
        /// <param name="properties"></param>
        void SetCustomProperties(Excel.CustomProperties properties);

        /// <summary>
        /// Removes the Custom properties
        /// </summary>
        /// <param name="properties"></param>
        void DeleteCustomProperties(Excel.CustomProperties properties);

        /// <summary>
        /// Gets the Custom Properties from this workbook
        /// </summary>
        /// <returns></returns>
        Excel.CustomProperties GetCustomProperties();

        /// <summary>
        /// Gets the customProperty. returns true if the name exists. ref value will contain the value of the property
        /// </summary>
        /// <param name="name">the name of the customproperty</param>
        /// <param name="value">the value of the customproperty</param>
        /// <returns>true when the customproperty exists.</returns>
        bool TryGetCustomProperty(string name, ref object value);

        /// <summary>
        /// Add or updates the value of the customproperty.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        bool TrySetCustomProperty(string name, dynamic value);


        /// <summary>
        /// Sets the content of the xmldocument in the customxmlparts, and returns the id of the xmlpart.
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <returns></returns>
        string SetCustomXML(XmlDocument xmlDocument);

        /// <summary>
        /// Gets the content of the xmlpart (find by the given id), and returns as a xml document.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        XmlDocument GetCustomXMLById(string id);

        /// <summary>
        /// Delete the custom xml part.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        bool TryDeleteCustomXMLById(string id);
    }
}
