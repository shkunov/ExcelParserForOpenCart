using System;
using System.Collections.Generic;
using System.Linq;

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
                "Карданчик (шарнир) УАЗ-31512 рулевого управления (с ГУР Борисов)",
                "Карданчик (шарнир) УАЗ-31512 рулевого управления (с ГУР Борисов мелкий шлиц)",
                //"Шланг ГУР сливной УАЗ-469,  Хантер",
                //"Шланг ГУР сливной УАЗ-469, Хантер (импорт)"
            };

            var i = 0;
            var maxStr = "";
            var minStr = "";

            Console.WriteLine("Test 1");

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

            var words = minStr.Split(new[] { ' ', ',', ':', '?', '!', ')'}, StringSplitOptions.RemoveEmptyEntries);
            var options = "";
            i = 0;
            foreach (var str in list)
            {
                if (str == minStr) continue;
                var option = str.Replace(")", "");
                foreach (var word in words)
                {
                    if (word.Length == 1)
                        continue;
                    option = option.Replace(word, "");
                }
                option = option.Replace(",", "").Replace("(", "");
                if (i == 0)
                    options = option;
                else
                    options += " ; " + option;
                i++;
            }
            Console.WriteLine("Options: {0}", options.Trim());

            Console.WriteLine("Test 2");

            options = string.Empty;
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
