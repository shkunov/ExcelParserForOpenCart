using System.ComponentModel;
using System.Text.RegularExpressions;
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
            var j = 1;
            ResultingPrice.Clear();
            const string pattern = "[0-9]+";
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
                    j++;
                    continue;
                }
                if (startTable)
                {
                    var prefix = Regex.Match(category1, pattern).Value;
                    line.VendorCode = string.Format("{0}-{1}-{2}", string.IsNullOrWhiteSpace(prefix) ? category1 : prefix, str, j);
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
