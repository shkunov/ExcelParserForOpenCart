using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
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
            var code = string.Empty;
            var vendorCode = string.Empty;
            var pair = false;
            ResultingPrice.Clear();
            // список имён с одинаковым артикулем
            var list = new List<PairProductAndCost>();
            const string pattern = "(\\d+\\.\\s?)";
            for (var i = 11; i < row; i++)
            {
                if (Worker.CancellationPending)
                {
                    E.Cancel = true;
                    break;
                }
                var line = new OutputPriceLine();
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
                    //if (sc == "12710911") continue; // пока предлагаю эту графу пропускать, это как бы 3 категория
                }
                line.Category1 = category1;
                line.Category2 = category2;
                vendorCode = ConverterToString(range.Cells[i, 4] as Range);
                // получить артикул следующей строки для сравнения
                var tempVendorCode = ConverterToString(range.Cells[i + 1, 4] as Range);

                var tempName = ConverterToString(range.Cells[i, 3] as Range);
                var tempName2 = ConverterToString(range.Cells[i + 1, 3] as Range);
                var cost1 = ConverterToDecimal(range.Cells[i, 7] as Range);
                var cost2 = ConverterToDecimal(range.Cells[i + 1, 7] as Range);
                if (!pair)
                {
                    list.Clear();
                    list.Add(new PairProductAndCost
                    {
                        Name = tempName,
                        Cost = cost1
                    });
                }
                if (string.IsNullOrWhiteSpace(vendorCode) && string.IsNullOrWhiteSpace(tempVendorCode) && string.IsNullOrWhiteSpace(tempName))
                    break;
                //todo: протестировать случай если артикуль не заполнен
                if (tempVendorCode == vendorCode && !string.IsNullOrWhiteSpace(vendorCode))
                {
                    // дублирование
                    list.Add(new PairProductAndCost
                    {
                        Name = tempName2,
                        Cost = cost2
                    });
                    pair = true;
                    continue;
                }
                // получить из списка опции и имя
                /*if (list.Count >= 2)
                {
                    string name, options, costs;
                    GetNameAndOptionFromAutogur73(list, out name, out options, out costs);
                    line.Name = name;
                    line.Option = options;
                    line.Cost = list[0].Cost.ToString(CultureInfo.CurrentCulture);
                    line.PlusThePrice = costs;
                }
                else*/
                //опций нет
                //под опциями понимается дополнительные характеристики изделия, при том же артикуле см. AutogurPrice.cs
                
                line.Name = tempName;
                line.Cost = cost1.ToString(CultureInfo.CurrentCulture);                
                list.Clear();
                pair = false;
                line.VendorCode = string.IsNullOrEmpty(vendorCode) ? code : vendorCode;

                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && string.IsNullOrEmpty(line.Name))
                    break; // выходить из цикла
                if (string.IsNullOrEmpty(vendorCode) && string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(line.Name))
                    continue; // игнорировать строки без кода и артикля
                if (!string.IsNullOrEmpty(line.Name))
                    ResultingPrice.Add(line);
            }
        }

    }
}

