using System;
using System.Text.RegularExpressions;

namespace TestParsingOptionOJ
{
    class Program
    {
        static void Main()
        {
            string[] test = {
            "опция (стандарт, лифт 35, лифт 50)"
            , "стандарт, лифт 65 (опция лифт 50)"
            , "стандарт, лифт 30...50"
            };

            foreach (var s in test)
            {
                var input = s.Replace("опция", string.Empty).Trim();
                if (input[0] == '(') input = input.Replace("(", string.Empty);
                input = input.Replace(",", ";").Replace("(", ";").Replace(")", string.Empty);
                Console.WriteLine(input);
            }

            //foreach (var s in test)
            //{
            //    if (s.Contains("(опция"))
            //    {
                    
            //    }
            //    var input = s.Replace("опция", string.Empty).Trim();
            //    if (input[0] == '(') input = input.Replace("(", string.Empty);
            //    input = input.Replace(",", ";").Replace("(", ";").Replace(")", string.Empty);
            //    Console.WriteLine(input);
            //}

            foreach (var s in test)
            {
                var input = s.Replace("опция", string.Empty).Trim();
                var pattern = "((\\w?стандарт)|(\\w+\\s[0-9-]+))";
                var regex = new Regex(pattern);

                // Получаем совпадения в экземпляре класса Match
                var match = regex.Match(input);

                // отображаем все совпадения
                while (match.Success)
                {
                    // Т.к. мы выделили в шаблоне одну группу (одни круглые скобки),
                    // ссылаемся на найденное значение через свойство Groups класса Match
                    var str = match.Groups[1].Value;
                    Console.WriteLine(match.Groups[1].Value);

                    // Переходим к следующему совпадению
                    match = match.NextMatch();
                }
            }

            Console.ReadLine();
        }
    }
}
