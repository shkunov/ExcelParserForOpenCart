using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserForOpenCart
{
    static class Global
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
