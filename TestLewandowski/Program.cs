using System;
using System.Text.RegularExpressions;

namespace TestLewandowski
{
    class Program
    {
        static void Main()
        {
            var input = "";
            const string pattern = "[0-9]+";
            input = Regex.Match("Стеклопластиковые изделия на УАЗ-31512:", pattern).Value;
            Console.WriteLine(input);
            Console.ReadLine();
        }
    }
}
