using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    class ExcelParser
    {
        public event Action<string> OnParserAction;
        public event Action<int> OnProgressBarAction;

        private readonly bool _isExcelInstal;        
        private readonly BackgroundWorker _workerSave;
        private readonly BackgroundWorker _workerOpen;
        private List<OutputPriceLine> _list;
        private string _template;
        private string _fileNameForSave;

        public ExcelParser()
        {
            _isExcelInstal = true;
            if (!IsExcelInstall())
            {
                SendMessage("Excel не установлен!");
                _isExcelInstal = false;
                return;
            }
            _workerSave = new BackgroundWorker {WorkerReportsProgress = true};
            _workerSave.DoWork += _workerSave_DoWork;
            _workerSave.RunWorkerCompleted += _workerSave_RunWorkerCompleted;
            _workerSave.ProgressChanged += _workerSave_ProgressChanged;

            _workerOpen = new BackgroundWorker { WorkerReportsProgress = true };

        }

        private void _workerSave_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SendProgressBarInfo(e.ProgressPercentage);
        }

        private void _workerSave_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SendMessage("Сохраняю как: " + _fileNameForSave);
            SendMessage("Прайс создан!");
        }

        private void _workerSave_DoWork(object sender, DoWorkEventArgs e)
        {
            var application = new Application();
            var workbook = application.Workbooks.Open(_template);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            // действия по заполнению шаблона
            var i = 3;
            foreach (var obj in _list)
            {
                // заносить полученную линию в шаблон
                worksheet.Cells[i, 1] = obj.VendorCode;
                worksheet.Cells[i, 2] = obj.Name;
                worksheet.Cells[i, 3] = obj.Category1;
                worksheet.Cells[i, 4] = obj.Category2;
                worksheet.Cells[i, 5] = obj.ProductDescription;
                worksheet.Cells[i, 6] = obj.Cost;
                worksheet.Cells[i, 7] = obj.Foto;
                worksheet.Cells[i, 8] = obj.Option;
                worksheet.Cells[i, 9] = obj.Qt;
                worksheet.Cells[i, 10] = obj.PlusThePrice;
                i++;
            }
            worksheet.SaveAs(_fileNameForSave);
            application.Quit();
            _workerSave.ReportProgress(100);
        }

        public void OpenExcel(string fileName)
        {
            if (_isExcelInstal == false)
                return;
            // Open Excel and get first worksheet.
            var application = new Application();
            var workbook = application.Workbooks.Open(fileName);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            if (worksheet == null) return;
            var row = worksheet.Rows.Count;
            var column = worksheet.Columns.Count;
            for (var i = 0; i > row; i++)
            {
                
            }
        }

        public void SaveResult(string fileName)
        {
            if (_isExcelInstal == false)
                return;
            var list = new List<OutputPriceLine>();
            var line = new OutputPriceLine
            {
                Name = "Имя",
                VendorCode = "1234567",
                Category1 = "Багажники на Suzuki",
                Category2 = "Багажники",
                ProductDescription = "Огромный багажник, вместительностью до 300 кг",
                Cost = "3000",
                Foto = "http://img.yandex.ru/2034.jpg",
                Option = "Да",
                Qt = "1000",
                PlusThePrice = "100"
            };
            list.Add(line);
            var line2 = new OutputPriceLine
            {
                Name = "Имя",
                VendorCode = "7654321",
                Category1 = "Багажники на Suzuki",
                Category2 = "Багажники",
                ProductDescription = "Огромный багажник, вместительностью до 300 кг",
                Cost = "3000",
                Foto = "http://img.yandex.ru/2034.jpg",
                Option = "Нет",
                Qt = "2000",
                PlusThePrice = "200"
            };
            list.Add(line2);
            _list = list;
            _fileNameForSave = fileName;
            _template = Global.GetTemplate();
            if (_template == null)
            {
                SendMessage("Ошибка! Не могу получить путь к шаблону!");
                return;
            }
            if (!File.Exists(_template))
            {
                SendMessage("Ошибка! Отсутствует шаблон!");
                return;
            }
            _workerSave.RunWorkerAsync();
        }

        private void SendMessage(string message)
        {
            if (OnParserAction != null) OnParserAction(message);
        }

        private void SendProgressBarInfo(int i)
        {
            if (OnProgressBarAction != null) OnProgressBarAction(i);
        }

        private static bool IsExcelInstall()
        {
            var hkcr = Registry.ClassesRoot;
            var excelKey = hkcr.OpenSubKey("Excel.Application");
            return excelKey != null;
        }
    }
}
