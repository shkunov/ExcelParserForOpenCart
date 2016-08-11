using System;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class Autoventuri : GeneralMethods
    {
        public Autoventuri(object sender, DoWorkEventArgs e)
        {
            Worker = sender as BackgroundWorker;
            E = e;
        }

        public void Analyze(int row, Range range)
        {
            if(Worker.CancellationPending)
            {
                E.Cancel = true;
                ResultingPrice.Clear();
                return;
            }
            var category1 = string.Empty;
            var category2 = string.Empty;
            ResultingPrice.Clear();
            //const string pattern = "(\\d+\\.\\s?)";
            // цикл для обработки прайс листа \\начинаем с 11-й строки
            for (var i = 11; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }

                var theRange = range.Cells[i, 2] as Range; //берем из 2 столбца
                if (theRange != null)
                {
                    string str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "11842740") // 1 категория
                    {
                        category1 = str.TrimStart(new Char[] {' '});//Regex.Replace(str, pattern, string.Empty);
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "12829635") // 2 категория
                    {
                        category2 = str.TrimStart(new Char[] { ' ' });
                        continue;
                    }
                }

                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };
                var vendorCode = ConverterToString(range.Cells[i, 3] as Range);
                line.Name = ConverterToString(range.Cells[i, 2] as Range).TrimStart(new Char[] { ' ' });//тримим пробелы вначале строки
                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без артикля
                line.Cost = ConverterToString(range.Cells[i, 6] as Range);
                line.VendorCode = vendorCode;

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(line.Name))
                    break; // выходить из цикла

                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
                // выйти из цикла необходимо с помощью оператора break
            }
        }
    }
}
