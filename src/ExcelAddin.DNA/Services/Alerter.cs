using System.Windows.Forms;
using Application;

namespace ExcelAddIn.DNA
{
    internal class Alerter : IAlerter
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}