// <copyright file="BuiltinProperties.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>


namespace Allors.Excel.Headless
{
    using System;
    using Allors.Excel;

    internal class BuiltinProperties : IBuiltinProperties
    {
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Author { get; set; }
        public string Keywords { get; set; }
        public string Comments { get; set; }
        public string Template { get; set; }
        public string LastAuthor { get; set; }
        public string RevisionNumber { get; }
        public string ApplicationName { get; set; }
        public DateTime? LastPrintDate { get; set; }
        public DateTime? CreationDate { get; set; }
        public DateTime? LastSaveTime { get; set; }
        public long? TotalEditingTime { get; set; }
        public long? NumberOfPages { get; set; }
        public long? NumberOfWords { get; set; }
        public long? NumberOfCharacters { get; set; }
        public long? Security { get; set; }
        public string Category { get; set; }
        public string Format { get; set; }
        public string Manager { get; set; }
        public string Company { get; set; }
        public long? NumberOfBytes { get; set; }
        public long? NumberOfLines { get; set; }
        public long? NumberOfParagraphs { get; set; }
        public long? NumberOfSlides { get; set; }
        public long? NumberOfNotes { get; set; }
        public long? NumberOfHiddenSlides { get; set; }
        public long? NumberOfMultimediaClips { get; set; }
        public string HyperlinkBase { get; set; }
        public long? NumberOfCharactersWithSpaces { get; set; }
        public string ContentType { get; set; }
        public string ContentStatus { get; set; }
        public string Language { get; set; }
        public string DocumentVersion { get; set; }
    }
}
