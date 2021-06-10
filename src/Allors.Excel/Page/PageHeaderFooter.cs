namespace Allors.Excel
{
    /// <summary>
    /// PageHeaderFooter defines the header and footer style.
    /// </summary>
    public class PageHeaderFooter
    {
        /// <summary>
        /// Margin in Points from either the top (Header) or bottom (Footer)
        /// </summary>
        public double Margin { get; set; }

        /// <summary>
        /// string contains the text on the left.
        /// </summary>       
        public string Left { get; set; }

        /// <summary>
        /// string contains the text in the center      
        /// </summary>
        public string Center { get; set; }

        /// <summary>
        /// string contains the text on the right.       
        /// </summary>
        public string Right { get; set; }
    }
}
