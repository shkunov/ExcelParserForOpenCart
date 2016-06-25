using System.IO;
using System.Reflection;

namespace ExcelParserForOpenCart
{
    public static class Global
    {
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
