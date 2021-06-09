namespace Allors.Excel
{
    using System;

    public interface IBuiltinProperties
    {
        string Title { get; set; }

        string Subject { get; set; }

        string Author { get; set; }

        string Keywords { get; set; }

        string Comments { get; set; }

        string Template { get; set; }

        string LastAuthor { get; set; }

        string RevisionNumber { get; }

        string ApplicationName { get; set; }

        DateTime? LastPrintDate { get; set; }

        DateTime? CreationDate { get; set; }

        DateTime? LastSaveTime { get; set; }

        long? TotalEditingTime { get; set; }

        long? NumberOfPages { get; set; }

        long? NumberOfWords { get; set; }

        long? NumberOfCharacters { get; set; }

        long? Security { get; set; }

        string Category { get; set; }

        string Format { get; set; }

        string Manager { get; set; }

        string Company { get; set; }

        long? NumberOfBytes { get; set; }

        long? NumberOfLines { get; set; }

        long? NumberOfParagraphs { get; set; }

        long? NumberOfSlides { get; set; }

        long? NumberOfNotes { get; set; }

        long? NumberOfHiddenSlides { get; set; }

        long? NumberOfMultimediaClips { get; set; }

        string HyperlinkBase { get; set; }

        long? NumberOfCharactersWithSpaces { get; set; }

        string ContentType { get; set; }

        string ContentStatus { get; set; }

        string Language { get; set; }

        string DocumentVersion { get; set; }
    }
}
