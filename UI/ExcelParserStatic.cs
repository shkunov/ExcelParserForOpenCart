using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    public partial class ExcelParser
    {
        private static bool IsExcelInstall()
        {
            var hkcr = Registry.ClassesRoot;
            var excelKey = hkcr.OpenSubKey("Excel.Application");
            return excelKey != null;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static string ConverterToString(dynamic obj)
        {
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }

        private static decimal ConverterToDecimal(Range range)
        {
            if (range == null)
                return 0;
            var obj = range.Value2;
            if (obj == null)
                return 0;
            decimal d;
            try
            {
                d = Convert.ToDecimal(obj);
            }
            catch
            {
                d = 0;
            }
            return d;
        }

        private static string ConverterToString(Range range)
        {
            if (range == null)
                return string.Empty;
            var obj = range.Value2;
            if (obj == null)
                return string.Empty;
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }
        /// <summary>
        /// Поиск опции, прайс Каталог OJ 2016_06_01 вер. 6
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string OjOptionParser(string s)
        {
            string input;
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;
            if (s.Contains("-"))
                return string.Empty;
            if (s.Contains("(опция"))
            {
                input = s.Replace("опция", string.Empty)
                    .Replace(")", string.Empty)
                    .Replace("(", ";")
                    .Replace(",", string.Empty);
                return input;
            }
            input = s.Replace("опция", string.Empty).Trim();
            if (input.Length < 1)
                return string.Empty;
            if (input[0] == '(') input = input.Replace("(", string.Empty);
            input = input.Replace(",", ";").Replace("(", ";").Replace(")", string.Empty);
            return input;
        }
        /// <summary>
        /// Парсинг опции для прайса ИП Пьянов С.Г. Autogur73.ru
        /// </summary>
        /// <param name="list"></param>
        /// <param name="name"></param>
        /// <param name="options"></param>
        /// <param name="costs"></param>
        private static void GetNameAndOptionFromAutogur73(IReadOnlyList<PairProductAndCost> list, 
            out string name, out string options, out string costs)
        {
            var i = 0;
            var maxStr = "";
            var minStr = "";
            var @case = 1;
            decimal cost = 0;
            costs = "";
            options = "";
            foreach (var s in list)
            {
                if (maxStr.Length < s.Name.Length)
                    maxStr = s.Name;
            }
            foreach (var s in list)
            {
                if (i == 0)
                {
                    minStr = s.Name;
                    i++;
                    continue;
                }
                if (s.Name.Length < minStr.Length)
                    minStr = s.Name;
            }
            if (maxStr.Length - minStr.Length < 5) @case = 3;            
            name = minStr;
            if (@case == 1)
            {
                options = string.Empty;
                i = 0;
                foreach (var item in list)
                {
                    var option = item.Name.Replace(minStr, string.Empty).Replace(",", "").Trim();

                    if (option.Length > 19)
                    {
                        @case = 2;
                        break;
                    }

                    if (string.IsNullOrEmpty(option)) continue;
                    if (i == 0)
                    {
                        options = option.Trim();
                        cost = item.Cost;
                        costs = "";
                    }
                    else
                    {
                        options += " ; " + option.Trim();
                        var diff = item.Cost - cost;
                        costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    }
                    i++;
                }
                options = options.Trim();
            }
            if (@case == 2)
            {
                options = string.Empty;
                var words = minStr.Split(new[] { ' ', ',', ':', '?', '!', ')' }, StringSplitOptions.RemoveEmptyEntries);
                i = 0;
                foreach (var item in list)
                {
                    if (item.Name == minStr) continue;
                    var option = item.Name.Replace(")", "");
                    foreach (var word in words)
                    {
                        if (word.Length == 1)
                            continue;
                        option = option.Replace(word, "");
                    }
                    option = option.Replace(",", "").Replace("(", "");
                    if (i == 0)
                    {
                        options = option.Trim();
                        cost = item.Cost;
                        costs = "";
                    }
                    else
                    {
                        options += " ; " + option.Trim();
                        var diff = item.Cost - cost;
                        costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    }
                    i++;
                }
                options = options.Trim();
            }
            if (@case != 3) return;
            i = 0;
            foreach (var item in list)
            {
                var j = item.Name.LastIndexOf(',') + 1;
                var option = "";
                for (var k = j; k < item.Name.Length; k++)
                {
                    option += item.Name[k];
                }
                if (i == 0)
                {
                    options = option.Trim();
                    cost = item.Cost;
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                else
                {
                    options += " ; " + option.Trim();
                    var diff = item.Cost - cost;
                    costs += " ; " + diff.ToString(CultureInfo.CurrentCulture);
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                i++;
            }
        }
        /// <summary>
        /// Определение типа прайс листа
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private static EnumPrices DetermineTypeOfPriceList(Range range)
        {
            var str = ConverterToString(range.Cells[2, 3] as Range);
            if (str.Contains("Два Союза"))
                return EnumPrices.ДваСоюза;

            var str1 = ConverterToString(range.Cells[1, 1] as Range);
            var str2 = ConverterToString(range.Cells[1, 4] as Range);
            if (str1.Contains("Рисунок") && str2.Contains("Марка и модель автомобиля"))
                return EnumPrices.OJ;

            str1 = ConverterToString(range.Cells[9, 3] as Range);
            str2 = ConverterToString(range.Cells[11, 3] as Range);

            if (str1.Contains("Прайс-лист") && str2.Contains("Наименование товаров"))
                return EnumPrices.Autogur73;

            return EnumPrices.Неизвестный;
        }
    }
}
