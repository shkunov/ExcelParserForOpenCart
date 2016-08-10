using System.ComponentModel;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    class PTGrupp : GeneralMethods
    {
        public PTGrupp(object sender, DoWorkEventArgs e)
        {
            Worker = sender as BackgroundWorker;
            E = e;
        }

        public void Analyze(int row, Range range)
        {
            if (Worker.CancellationPending)
            {
                E.Cancel = true;
                ResultingPrice.Clear();
                return;
            }
            var category1 = string.Empty;
            var category2 = string.Empty;
            ResultingPrice.Clear();
            const string pattern = "(\\d+\\.\\s?)";
            for (var i = 11; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }
                var theRange = range.Cells[i, 3] as Range;
                if (theRange != null)
                {
                    string str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "13816530") // 1 категория
                    {
                        category1 = Regex.Replace(str, pattern, string.Empty);
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "15132390") // 2 категория
                    {
                        category2 = str;
                        continue;
                    }
                }
                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };
                var vendorCode = ConverterToString(range.Cells[i, 4] as Range);
                line.Name = ConverterToString(range.Cells[i, 3] as Range);
                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без артикля
                line.Cost = ConverterToString(range.Cells[i, 7] as Range);                               
                line.VendorCode = vendorCode;

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(line.Name))
                    break; // выходить из цикла
      
                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
            }
        }
    }
}

