using System;
using System.Collections.Generic;
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

        private static string OptionParser(string s)
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

        private static void GetNameAndOption(IReadOnlyList<string> list, out string name, out string options)
        {
            var i = 0;
            var minLen = 0;
            var indexMinLen = 0;
            foreach (var str in list)
            {
                if (i == 0)
                {
                    minLen = str.Length;
                    i++;
                    continue;
                }
                if (minLen > str.Length)
                {
                    minLen = str.Length;
                    indexMinLen = i;
                }
                i++;
            }
            var minStr = list[indexMinLen];
            options = string.Empty;
            i = 0;
            foreach (var str in list)
            {
                var option = str.Replace(minStr, string.Empty).Replace(",", "").Trim();
                if (string.IsNullOrEmpty(option)) continue;
                if (i == 0)
                    options = option;
                else
                    options += " ; " + option;
                i++;
            }
            name = minStr;
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
