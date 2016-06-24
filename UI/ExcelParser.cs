using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart
{
    class ExcelParser
    {

        public void OpenExcel(string fileName)
        {
            // Open Excel and get first worksheet.
            var application = new Application();
            var workbook = application.Workbooks.Open(fileName);
        }

        public void SaveResult(string fileName)
        {
            var list = new List<OutputPriceLine>();
            var line = new OutputPriceLine
            {
                Name = "Имя",
                VendorCode = "1234567"
            };
            list.Add(line);
            var template = Global.GetTemplate();
            if (template == null)
            {
                //обработать ошибку
                return;
            }
            if (!File.Exists(template))
            {
                //обработать ошибку отсутствия шаблона
                return;
            }
            var application = new Application();
            var workbook = application.Workbooks.Open(template);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            // действия по заполнению шаблона
            if (worksheet != null) worksheet.SaveAs(fileName);
        }
    }
}
