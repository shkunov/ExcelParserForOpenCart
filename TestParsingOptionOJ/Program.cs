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
                string input;
                if (s.Contains("(опция"))
                {
                    input = s.Replace("опция", string.Empty)
                        .Replace(")", string.Empty)
                        .Replace("(", ";")
                        .Replace(",", string.Empty);
                    Console.WriteLine(input);
                    continue;
                }
                input = s.Replace("опция", string.Empty).Trim();
                if (input[0] == '(') input = input.Replace("(", string.Empty);
                input = input.Replace(",", ";").Replace("(", ";").Replace(")", string.Empty);
                Console.WriteLine(input);
            }

            Console.ReadLine();
        }
    }
}
