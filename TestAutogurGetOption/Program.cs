using System;
using System.Collections.Generic;

namespace TestAutogurGetOption
{
    class Program
    {
        static void Main()
        {
            var list = new List<string>
            {
                //"ГУР 452 (YuBei) дв.ЗМЗ-402, 410 с механизмом Газель, Соболь",
                //"ГУР 452 (YuBei) дв.ЗМЗ-402, 410 с механизмом Газель, Соболь, Лифт (50-100)",
                //"ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF",
                //"ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF, Лифт (50-100)",
                //"ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF Люкс"
                // такой случай возможен придётся править его ручками
                "Шланг ГУР сливной УАЗ-469,  Хантер",
                "Шланг ГУР сливной УАЗ-469, Хантер (импорт)"
            };
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
            Console.WriteLine("Min len: {0}", minStr);
            var options = string.Empty;
            i = 0;
            foreach (var str in list)
            {
                var option = str.Replace(minStr, string.Empty).Replace(",", "").Trim();
                if (string.IsNullOrEmpty(option)) continue;
                Console.WriteLine("Option: {0}", option);
                if (i == 0)
                    options = option;
                else
                    options += " ; " + option;
                i++;
            }
            Console.WriteLine("Options: {0}", options);
            Console.ReadLine();
        }
    }
}
