using System.Windows.Forms;
using Application;

namespace Allors.Excel.Interop.Vsto
{
    internal class Alerter : IAlerter
    {
        public void Alert(string message) => MessageBox.Show(message);
    }
}
