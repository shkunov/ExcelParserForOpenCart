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

            var countEmptyRow = 0;
            var icount = 0;
            var compareCategory2 = string.Empty;
            var compareVendorCode = string.Empty;
            var unionDescription = string.Empty;
            ResultingPrice.Clear();

            var category1 = ConverterToString(range.Cells[4, 1] as Range); //1 категория
            var category2 = string.Empty;
            // цикл для обработки прайс листа
            for (var i = 6; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    ResultingPrice.Clear();
                    break;
                }

                var secRange = range.Cells[i, 1] as Range; //2 категория


                if (secRange != null)
                {
                    string str = ConverterToString(secRange.Value2);
                    var color = secRange.Interior.Color;
                    var sc = color.ToString();

                    if (sc == "15986394") // 2 категория
                    {
                        compareCategory2 = category2 = str;
                        countEmptyRow = 0; //идет новая категория 2, зануляем счет на пустые строки
                        continue;
                    }
                    else
                    { category2 = str; }
                }

                if (secRange != null)
                {
                    string str = ConverterToString(secRange.Value2);
                    {
                        category2 = str;
                    }
                }

                var line = new OutputPriceLine
                {
                    Category1 = category1,
                    Category2 = category2
                };


                var vendorCode = ConverterToString(range.Cells[i, 5] as Range);

                if (compareVendorCode != vendorCode)
                {
                    compareVendorCode = vendorCode;
                    unionDescription = "<p>" + ConverterToString(range.Cells[i, 2] as Range) + "(" + ConverterToString(range.Cells[i, 3] as Range) + ")" + "</p>";
                    line.ProductDescription = unionDescription;
                }
                else if (compareVendorCode == vendorCode)
                {
                    unionDescription += "<p>" + ConverterToString(range.Cells[i, 2] as Range) + "(" + ConverterToString(range.Cells[i, 3] as Range) + ")" + "</p>";                    
                    ResultingPrice[icount-1].ProductDescription = unionDescription; //модифицируем
                    if (unionDescription == "<p>()</p><p>()</p><p>()</p><p>()</p>") { break; }// выйти из цикла, при пустых 4-х строк
                    else 
                    continue; // пропускаем строку
                }


                line.Name = ConverterToString(range.Cells[i, 4] as Range);


                if (string.IsNullOrEmpty(vendorCode) && !string.IsNullOrEmpty(line.Name))
                {
                    continue; // игнорировать строки без артикля
                }
                line.Cost = ConverterToString(range.Cells[i, 17] as Range);
                line.VendorCode = vendorCode;

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(line.Name))
                { countEmptyRow++; }

                if (countEmptyRow >= 3) { break; } // выходить из цикла, после 3-й пустой строки

                if (!string.IsNullOrEmpty(line.Name))
                { ResultingPrice.Add(line); icount++;}

            }
            //здесь выполнить "схлопывание" записей по одинаковому артикулу, цене, объединить значения в ProductDescription
        }
    }
}
