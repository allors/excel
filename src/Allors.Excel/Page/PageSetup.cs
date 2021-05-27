namespace Allors.Excel
{
    /// <summary>
    /// sss
    /// </summary>
    public class PageSetup
    {
        /// <summary>
        /// Orientation Portrait (1) or Landscape (2). Defaults to Portrait
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpageorientation?view=excel-pia
        /// </summary>
        public int Orientation { get; set; } = 1; // Portrait

        /// <summary>
        /// PaperSize of the printed worksheet.
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpapersize?view=excel-pia
        /// </summary>
        public int PaperSize { get; set; } = 9; // A4

        public PageHeaderFooter Header { get; set; }
             
        public PageHeaderFooter Footer { get; set; }
    }
}
