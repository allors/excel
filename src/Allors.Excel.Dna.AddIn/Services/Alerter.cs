namespace ExcelDNA
{
    using System.Windows.Forms;
    using Application;

    internal class Alerter : IAlerter
    {
        public void Alert(string message) => MessageBox.Show(message);
    }
}
