using Application;
using System.Windows.Forms;

namespace ExcelAddIn.Services
{
    internal class Alerter : IAlerter
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}