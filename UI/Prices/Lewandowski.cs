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
            var j = 0;
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
                var str = ConverterToString(range.Cells[i, 1] as Range);
                var name = ConverterToString(range.Cells[i, 2] as Range);
                if (string.IsNullOrWhiteSpace(name)) break;
                if (string.IsNullOrWhiteSpace(str.Trim()))
                {
                    startTable = false;
                    category1 = name.Replace(':', ' ').Trim();
                    continue;
                }
                if (str.Contains("№"))
                {
                    startTable = true;
                    j++;
                    continue;
                }
                if (!startTable) continue;
                var postfix = Regex.Match(category1, pattern).Value;
                var line = new OutputPriceLine
                {
                    VendorCode = string.Format("A-{0}-{1}-{2}", j, str, string.IsNullOrWhiteSpace(postfix) ? category1 : postfix),
                    Name = name.Trim(),
                    Category1 = category1,
                    Cost = ConverterToString(range.Cells[i, 4] as Range)
                };
                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
            }
        }

    }
}
