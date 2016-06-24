using System.Windows;
using Microsoft.Win32;

namespace ExcelParserForOpenCart
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly ExcelParser _excelParser;

        public MainWindow()
        {
            InitializeComponent();
            _excelParser = new ExcelParser();
        }

        private string CreateOpenFileDialog()
        {
            var filename = string.Empty;
            var dlg = new OpenFileDialog { Filter = "Excel files|;*.xlsx;*.xls" };
            dlg.FileOk += delegate
            {
                filename = dlg.FileName;
            };
            dlg.ShowDialog(this);
            return filename;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var filename = CreateOpenFileDialog();
            if (string.IsNullOrEmpty(filename)) return;
            _excelParser.OpenExcel(filename);
            BtnSave.IsEnabled = true;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var filename = string.Empty;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel files|;*.xlsx;*.xls"
            };
            dlg.FileOk += delegate
            {
                filename = dlg.FileName;
            };
            dlg.ShowDialog(this);
            if (string.IsNullOrEmpty(filename)) return;
            _excelParser.SaveResult(filename);
        }
    }
}
