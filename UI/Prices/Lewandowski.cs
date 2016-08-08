using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class Lewandowski : GeneralMethods
    {
        public Lewandowski(object sender, DoWorkEventArgs e)
        {
            Worker = sender as BackgroundWorker;
            E = e;
        }

        public void Analyze(int row, Range range)
        {
            if (Worker.CancellationPending)
            {
                E.Cancel = true;
                return;
            }
            var startTable = false;
            var category1 = string.Empty;
            ResultingPrice.Clear();
            for (var i = 7; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                var line = new OutputPriceLine();
                var str = ConverterToString(range.Cells[i, 1] as Range);
                if (string.IsNullOrWhiteSpace(str.Trim()))
                {
                    startTable = false;
                    category1 = ConverterToString(range.Cells[i, 2] as Range);
                }
                if (str.Contains("Наименование"))
                {
                    startTable = true;
                    continue;
                }
                // todo: нужно вычислить артикуль
                //line.VendorCode = ;
                if (startTable)
                {
                    line.Name = ConverterToString(range.Cells[i, 2] as Range);
                    line.Category1 = category1;
                    line.Cost = ConverterToString(range.Cells[i, 4] as Range);
                    if (!string.IsNullOrEmpty(line.Name))
                        ResultingPrice.Add(line);
                }
                if (string.IsNullOrWhiteSpace(category1)) break;
            }
        }

    }
}
