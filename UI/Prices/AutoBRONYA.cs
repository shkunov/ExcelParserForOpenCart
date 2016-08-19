using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class AutoBronya : GeneralMethods
    {
        public AutoBronya(object sender, DoWorkEventArgs e)
        {
            Worker = sender as BackgroundWorker;
            E = e;
        }
        /// <summary>
        /// Обработка ЕКБ_Прайс АвтоБРОНЯ_Игорь.xls
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
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
            var secMarket = string.Empty;
            var countEmptyRow = 0;
            var compareCategory1 = string.Empty;
            ResultingPrice.Clear();
            // цикл для обработки прайс листа
            for (var i = 6; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }

                var theRange = range.Cells[i, 1] as Range; //1 категория
                var secRange = range.Cells[i, 4] as Range; //2 категория
                

                if (theRange != null)
                {
                    string str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();



                    if (sc == "15986394") // 1 категория
                    {
                        compareCategory1 = category1 = str;
                        //category2 = string.Empty;
                        secMarket = string.Empty;
                        continue;
                    }
                    else
                    { category1 = str; }                    
                }

                if (secRange != null)
                {
                    string str = ConverterToString(secRange.Value2);
                    {
                        category2 = str;
                        //continue;
                    }
                }

                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };

                var vendorCode = ConverterToString(range.Cells[i, 5] as Range);

                line.Name = ConverterToString(range.Cells[i, 2] as Range);

                if (line.Category1.Trim().ToUpper() != compareCategory1.Trim().ToUpper() && secMarket != "")
                {
                    secMarket = string.Empty;
                }

                line.Name = line.Name + ((secMarket != "") ? " (" + secMarket + ")" : "");

                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                {
                    if (line.Name != secMarket) { secMarket = line.Name; };
                    continue; // игнорировать строки без артикля
                }
                line.Cost = ConverterToString(range.Cells[i, 17] as Range);
                line.VendorCode = vendorCode;

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(line.Name))
                { countEmptyRow++; }

                if (countEmptyRow >= 2) { break; } // выходить из цикла, после 2-й пустой строки
                
                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);

            }
        }
    }
}
