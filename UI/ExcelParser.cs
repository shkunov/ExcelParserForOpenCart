using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    class ExcelParser
    {
        public event Action<string> OnParserAction;

        public ExcelParser()
        {
            if (!IsExcelInstall())
            {
                SendMessage("Excel не установлен!");
            }
        }

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
                SendMessage("Ошибка! Не могу получить путь к шаблону!");
                return;
            }
            if (!File.Exists(template))
            {
                SendMessage("Ошибка! Отсутствует шабон!");
                return;
            }
            var application = new Application();
            var workbook = application.Workbooks.Open(template);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            // действия по заполнению шаблона
            var i = 2;
            foreach (var obj in list)
            {
                // заносить полученную линию в шаблон
                worksheet.Cells[i, 1] = obj.VendorCode;
                worksheet.Cells[i, 2] = obj.Name;
                i++;
            }
            worksheet.SaveAs(fileName);
            SendMessage("Сохраняю как: " + fileName);
            application.Quit();
            SendMessage("Прайс создан!");
        }

        private void SendMessage(string message)
        {
            if (OnParserAction != null) OnParserAction(message);
        }

        private static bool IsExcelInstall()
        {
            var hkcr = Registry.ClassesRoot;
            var excelKey = hkcr.OpenSubKey("Excel.Application");
            return excelKey != null;
        }
    }
}
