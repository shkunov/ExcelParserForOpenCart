﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace TestAutogurGetOption
{
    class Program
    {
        private static void GetNameAndOption(IReadOnlyList<string> list, out string name, out string options)
        {
            var i = 0;
            var maxStr = "";
            var minStr = "";
            var @case = 1;
            options = "";
            foreach (var s in list)
            {
                if (maxStr.Length < s.Length)
                    maxStr = s;
            }
            foreach (var s in list)
            {
                if (i == 0)
                {
                    minStr = s;
                    i++;
                    continue;
                }
                if (s.Length < minStr.Length)
                    minStr = s;
            }
            if (maxStr.Length - minStr.Length < 5) @case = 2;
            name = minStr;
            if (@case == 1)
            {
                options = string.Empty;
                var words = minStr.Split(new[] { ' ', ',', ':', '?', '!', ')', '(' }, StringSplitOptions.RemoveEmptyEntries);
                i = 0;
                foreach (var str in list)
                {
                    if (str == minStr) continue;
                    var tmpWords = str.Split(new[] { ' ', ',', ':', '?', '!', ')', '(' }, StringSplitOptions.RemoveEmptyEntries);
                    var option = tmpWords.Except(words).Aggregate("", (current, w) => current + (w + " "));
                    if (option.Length > 19)
                    {
                        @case = 2;
                        break;
                    }
                    if (i == 0)
                        options = option;
                    else
                        options += " ; " + option;
                    i++;
                }
                options = options.Trim();
            }
            if (@case != 2) return;
            i = 0;
            foreach (var str in list)
            {
                var j = str.LastIndexOf(',') + 1;
                var option = "";
                for (var k = j; k < str.Length; k++)
                {
                    option += str[k];
                }
                if (i == 0)
                {
                    options = option.Trim();
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                else
                {
                    options += " ; " + option.Trim();
                    name = name.Replace(option, "").Replace(",", "").Trim();
                }
                i++;
            }
        }

        private static void Test(IEnumerable<string> list)
        {
            var i = 0;
            var options = "";
            foreach (var str in list)
            {
                var j = str.LastIndexOf(',') + 1;
                var option = "";
                for (var k = j; k < str.Length; k++)
                {
                    option += str[k];
                }
                if (i == 0)
                    options = option.Trim();
                else
                    options += " ; " + option.Trim();
                i++;
            }
        }

        private static void ListToConsole(IEnumerable<string> list)
        {
            foreach (var s in list)
            {
                Console.WriteLine(s);
            }
        }

        static void Main()
        {
            var list = new List<string>
            {
                "Карданчик (шарнир) УАЗ-31512 рулевого управления (с ГУР Борисов)",
                "Карданчик (шарнир) УАЗ-31512 рулевого управления (с ГУР Борисов мелкий шлиц)"
            };
            var name = "";
            var options = "";
            Console.WriteLine("Case 1");
            ListToConsole(list);
            GetNameAndOption(list, out name, out options);
            Console.WriteLine("Options: {0}", options);

            list = new List<string>
            {
                "ГУР 452 (YuBei) дв.ЗМЗ-402, 410 с механизмом Газель, Соболь",
                "ГУР 452 (YuBei) дв.ЗМЗ-402, 410 с механизмом Газель, Соболь, Лифт (50-100)"
            };
            Console.WriteLine("Case 2");
            ListToConsole(list);
            GetNameAndOption(list, out name, out options);
            Console.WriteLine("Options: {0}", options);

            list = new List<string>
            {
                "ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF",
                "ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF, Лифт (50-100)",
                "ГУР 452 (г. Борисов) дв. УМЗ-421 с насосом ZF Люкс"
            };
            Console.WriteLine("Case 3");
            ListToConsole(list);
            GetNameAndOption(list, out name, out options);
            Console.WriteLine("Options: {0}", options);

            list = new List<string>
            {
                "Шланг ГУР сливной УАЗ-469,  Хантер",
                "Шланг ГУР сливной УАЗ-469, Хантер (импорт)"
            };
            Console.WriteLine("Case 4");
            ListToConsole(list);
            GetNameAndOption(list, out name, out options);
            Console.WriteLine("Options: {0}", options);

            list = new List<string>
            {
                "Дифференциал самоблок. ВАЗ-1111 винтовой \"ValRacing\" , Стандарт",
                "Дифференциал самоблок. ВАЗ-1111 винтовой \"ValRacing\" , Туризм",
                "Дифференциал самоблок. ВАЗ-1111 винтовой \"ValRacing\" , Спорт"
            };
            ListToConsole(list);
            //Test(list);
            GetNameAndOption(list, out name, out options);
            Console.WriteLine("Options: {0}", options);
            Console.ReadLine();
        }
    }
}
