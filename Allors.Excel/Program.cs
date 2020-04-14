using System.Linq;
using System.Windows.Forms;

namespace Application
{
    using System.Drawing;
    using System.Threading.Tasks;
    using Dipu.Excel;

    public class Program : IProgram
    {
        public IAddIn AddIn { get; private set; }

        public async Task OnStart(IAddIn addIn)
        {
            this.AddIn = addIn;

            this.AddIn.Register(Actions.Dosomething, () => MessageBox.Show("Boom!!!!"));
        }

        public async Task OnStop()
        {
        }

        public async Task OnNew(IWorkbook workbook)
        {
            var style = new Style(Color.Aqua, Color.Blue);

            var sheet = workbook.CreateSheet();
            sheet.Name = "2";
            sheet.CellChanged += (sender, v) => { MessageBox.Show($"Cells changed: {string.Join(",", v.Cells.Select(w => $"{w.Row}:{w.Column}"))}"); };

            for (var i = 0; i < 100; i++)
            {
                for (var j = 0; j < 10; j++)
                {
                    sheet[i, j].Value = decimal.Parse($"{i},{j}");
                    if (j == 0 || j == 2)
                    {
                        sheet[i, j].Style = style;
                        sheet[i, j].NumberFormat = "#.###,00";
                    }
                }
            }

            await sheet.Flush();

            sheet[0, 0].Value = "Whoppa!";

            await sheet.Flush();
        }

        public void OnClose(IWorkbook workbook, ref bool cancel)
        {
        }

        public async Task OnNew(IWorksheet worksheet)
        {
            worksheet.Name = "1";
        }
    }
}