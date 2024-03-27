namespace Allors.Excel
{
    using System;

    public class BuiltinProperties(object properties) : PropertiesBase(properties), IBuiltinProperties
    {
        private const string TitleKey = "Title";
        private const string SubjectKey = "Subject";
        private const string AuthorKey = "Author";
        private const string KeywordsKey = "Keywords";
        private const string CommentsKey = "Comments";
        private const string TemplateKey = "Template";
        private const string LastAuthorKey = "Last author";
        private const string RevisionNumberKey = "Revision number";
        private const string ApplicationNameKey = "Application name";
        private const string LastPrintDateKey = "Last print date";
        private const string CreationDateKey = "Creation date";
        private const string LastSaveTimeKey = "Last save time";
        private const string TotalEditingTimeKey = "Total editing time";
        private const string NumberOfPagesKey = "Number of pages";
        private const string NumberOfWordsKey = "Number of words";
        private const string NumberOfCharactersKey = "Number of characters";
        private const string SecurityKey = "Security";
        private const string CategoryKey = "Category";
        private const string FormatKey = "Format";
        private const string ManagerKey = "Manager";
        private const string CompanyKey = "Company";
        private const string NumberOfBytesKey = "Number of bytes";
        private const string NumberOfLinesKey = "Number of lines";
        private const string NumberOfParagraphsKey = "Number of paragraphs";
        private const string NumberOfSlidesKey = "Number of slides";
        private const string NumberOfNotesKey = "Number of notes";
        private const string NumberOfHiddenSlidesKey = "Number of hidden Slides";
        private const string NumberOfMultimediaClipsKey = "Number of multimedia clips";
        private const string HyperlinkBaseKey = "Hyperlink base";
        private const string NumberOfCharactersWithSpacesKey = "Number of characters (with spaces)";
        private const string ContentTypeKey = "Content type";
        private const string ContentStatusKey = "Content status";
        private const string LanguageKey = "Language";
        private const string DocumentVersionKey = "Document version";

        public string Title { get => this.GetString(TitleKey); set => this.SetString(TitleKey, value); }

        public string Subject { get => this.GetString(SubjectKey); set => this.SetString(SubjectKey, value); }

        public string Author { get => this.GetString(AuthorKey); set => this.SetString(AuthorKey, value); }

        public string Keywords { get => this.GetString(KeywordsKey); set => this.SetString(KeywordsKey, value); }

        public string Comments { get => this.GetString(CommentsKey); set => this.SetString(CommentsKey, value); }

        public string Template { get => this.GetString(TemplateKey); set => this.SetString(TemplateKey, value); }

        public string LastAuthor { get => this.GetString(LastAuthorKey); set => this.SetString(LastAuthorKey, value); }

        public string RevisionNumber => this.GetString(RevisionNumberKey);

        public string ApplicationName { get => this.GetString(ApplicationNameKey); set => this.SetString(ApplicationNameKey, value); }

        public DateTime? LastPrintDate { get => this.GetDate(LastPrintDateKey); set => this.SetDate(LastPrintDateKey, value); }

        public DateTime? CreationDate { get => this.GetDate(CreationDateKey); set => this.SetDate(CreationDateKey, value); }

        public DateTime? LastSaveTime { get => this.GetDate(LastSaveTimeKey); set => this.SetDate(LastSaveTimeKey, value); }

        public long? TotalEditingTime { get => this.GetNumber(TotalEditingTimeKey); set => this.SetNumber(TotalEditingTimeKey, value); }

        public long? NumberOfPages { get => this.GetNumber(NumberOfPagesKey); set => this.SetNumber(NumberOfPagesKey, value); }

        public long? NumberOfWords { get => this.GetNumber(NumberOfWordsKey); set => this.SetNumber(NumberOfWordsKey, value); }

        public long? NumberOfCharacters { get => this.GetNumber(NumberOfCharactersKey); set => this.SetNumber(NumberOfCharactersKey, value); }

        public long? Security { get => this.GetNumber(SecurityKey); set => this.SetNumber(SecurityKey, value); }

        public string Category { get => this.GetString(CategoryKey); set => this.SetString(CategoryKey, value); }

        public string Format { get => this.GetString(FormatKey); set => this.SetString(FormatKey, value); }

        public string Manager { get => this.GetString(ManagerKey); set => this.SetString(ManagerKey, value); }

        public string Company { get => this.GetString(CompanyKey); set => this.SetString(CompanyKey, value); }

        public long? NumberOfBytes { get => this.GetNumber(NumberOfBytesKey); set => this.SetNumber(NumberOfBytesKey, value); }

        public long? NumberOfLines { get => this.GetNumber(NumberOfLinesKey); set => this.SetNumber(NumberOfLinesKey, value); }

        public long? NumberOfParagraphs { get => this.GetNumber(NumberOfParagraphsKey); set => this.SetNumber(NumberOfParagraphsKey, value); }

        public long? NumberOfSlides { get => this.GetNumber(NumberOfSlidesKey); set => this.SetNumber(NumberOfSlidesKey, value); }

        public long? NumberOfNotes { get => this.GetNumber(NumberOfNotesKey); set => this.SetNumber(NumberOfNotesKey, value); }

        public long? NumberOfHiddenSlides { get => this.GetNumber(NumberOfHiddenSlidesKey); set => this.SetNumber(NumberOfHiddenSlidesKey, value); }

        public long? NumberOfMultimediaClips { get => this.GetNumber(NumberOfMultimediaClipsKey); set => this.SetNumber(NumberOfMultimediaClipsKey, value); }

        public string HyperlinkBase { get => this.GetString(HyperlinkBaseKey); set => this.SetString(HyperlinkBaseKey, value); }

        public long? NumberOfCharactersWithSpaces { get => this.GetNumber(NumberOfCharactersWithSpacesKey); set => this.SetNumber(NumberOfCharactersWithSpacesKey, value); }

        public string ContentType { get => this.GetString(ContentTypeKey); set => this.SetString(ContentTypeKey, value); }

        public string ContentStatus { get => this.GetString(ContentStatusKey); set => this.SetString(ContentStatusKey, value); }

        public string Language { get => this.GetString(LanguageKey); set => this.SetString(LanguageKey, value); }

        public string DocumentVersion { get => this.GetString(DocumentVersionKey); set => this.SetString(DocumentVersionKey, value); }

        public void SetBoolean(string key, bool? value) => this.Set(key, value);

        public void SetDate(string key, DateTime? value) => this.Set(key, value);

        public void SetFloat(string key, double? value) => this.Set(key, value);

        public void SetNumber(string key, long? value) => this.Set(key, value);

        public void SetString(string key, string value) => this.Set(key, value);

        private void Set(string key, object? value)
        {
            if (value == null)
            {
                this.Remove(key);
                return;
            }

            var property = this.Get(key);
            if (property != null)
            {
                property.Value = value;
            }
        }
    }
}
