using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class For2Union : GeneralMethods
    {
        /// <summary>
        /// Обработка для прайса 2 союза
        /// </summary>
        /// <param name="row"></param>
        /// <param name="range"></param>
        public void Analyze(int row, Range range)
        {
            var category1 = string.Empty;
            var category2 = string.Empty;
            List = new List<OutputPriceLine>();
            for (var i = 9; i < row; i++)
            {
                var line = new OutputPriceLine();
                var str = string.Empty;
                var theRange = range.Cells[i, 1] as Range;
                if (theRange != null)
                {
                    str = ConverterToString(theRange.Value2);
                    var color = theRange.Interior.Color;
                    var sc = color.ToString();
                    if (sc == "0") // чёрный
                    {
                        category1 = str;
                        category2 = string.Empty;
                        continue;
                    }
                    if (sc == "8421504")
                    {
                        category2 = str;
                        continue;
                    }
                }
                line.Category1 = category1;
                line.Category2 = category2;
                line.VendorCode = ConverterToString(range.Cells[i, 3] as Range);
                line.Name = ConverterToString(range.Cells[i, 4] as Range);
                line.Qt = ConverterToString(range.Cells[i, 5] as Range);
                // todo: цена в USD может стоит её как-то обработать?
                line.Cost = ConverterToString(range.Cells[i, 6] as Range);

                if (!string.IsNullOrEmpty(line.VendorCode))
                    List.Add(line);
                if (string.IsNullOrEmpty(str)) break;
            }
        }
    }
}
