// <copyright file="IWorkbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;
    using System.Xml;

    public interface IWorkbook
    {

        event EventHandler<Allors.Excel.Hyperlink> OnHyperlinkClicked;

        /// <summary>
        /// Event FollowHyperLink triggers to method call.
        /// </summary>
        /// <param name="textToDisplay">the textpart of the hyperlink</param>
        void HyperlinkClicked(Hyperlink hyperlink);

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
        void Close(bool? saveChanges = null, string? fileName = null);

        /// <summary>
        /// Add a worksheet at the given location
        /// </summary>
        /// <param name="index"></param>
        /// <param name="before"></param>
        /// <param name="after"></param>
        /// <returns></returns>
        IWorksheet AddWorksheet(int? index = null, IWorksheet? before = null, IWorksheet? after = null);

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
        Range[] GetNamedRanges(string? refersToSheetName = null);

        /// <summary>
        /// Adds a NamedRange scoped to the Workbook
        /// </summary>
        /// <param name="name"></param>
        /// <param name="range"></param>
        void SetNamedRange(string name, Range range);

        IBuiltinProperties BuiltinProperties { get; }

        ICustomProperties CustomProperties { get; }

        IWorksheet[] WorksheetsByIndex { get; }

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
        XmlDocument? GetCustomXMLById(string id);

        /// <summary>
        /// Delete the custom xml part.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        bool TryDeleteCustomXMLById(string id);
    }
}
