using System.IO;
using System.Reflection;

namespace ExcelParserForOpenCart
{
    public static class Global
    {
        public static bool SearchFoto;
        public static bool SaveOnlyWithFoto;
        public static readonly string[] IgnorOptions = {"завод", "импорт"};
        /// <summary>
        /// Получить путь к файлу шаблона Excel
        /// </summary>
        /// <returns>Путь к файлу</returns>
        public static string GetTemplate()
        {
            const string fileName = "template.xls";
            var pathExe = Assembly.GetExecutingAssembly().Location;
            var dir = Path.GetDirectoryName(pathExe);
            if (string.IsNullOrEmpty(dir)) return null;
            var template = Path.Combine(dir, fileName);
            return template;
        }
    }
}
