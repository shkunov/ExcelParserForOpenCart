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
    }
}
