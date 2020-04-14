using Application;
using System.Windows.Forms;

namespace ExcelAddInLocal
{
    internal class Alerter : IAlerter
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}