using System;
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
            _excelParser = new ExcelParser();
            InitializeComponent();
            _excelParser.OnParserAction += OnParserAction;
            _excelParser.OnProgressBarAction += OnProgressBarAction;
            _excelParser.OnOpenDocument += OnOpenDocument;
            _excelParser.OnSaveDocument += OnSaveDocument;
        }

        private void OnSaveDocument(object sender, EventArgs eventArgs)
        {
            BtnOpen.IsEnabled = true;
            BtnSave.IsEnabled = false; 
        }

        private void OnOpenDocument(object sender, EventArgs e)
        {
            BtnOpen.IsEnabled = true;
            BtnSave.IsEnabled = true;            
        }

        private void OnProgressBarAction(int obj)
        {
            ProgressBar.Value = obj;
        }

        private void OnParserAction(string message)
        {
            MessageList.Items.Add(message);
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
            BtnOpen.IsEnabled = false;
            BtnSave.IsEnabled = false;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var filename = string.Empty;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel files|*.xls"
            };
            dlg.FileOk += delegate
            {
                filename = dlg.FileName;
            };
            dlg.ShowDialog(this);
            if (string.IsNullOrEmpty(filename)) return;
            _excelParser.SaveResult(filename);
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            switch (ComBoxSelectPrice.SelectedIndex)
            {
                case (int)EnumPrices.ДваСоюза:
                    _excelParser.PriceType = EnumPrices.ДваСоюза;
                    break;
                case (int)EnumPrices.OJ:
                    _excelParser.PriceType = EnumPrices.OJ;
                    break;
                case (int)EnumPrices.ПТГрупп:
                    _excelParser.PriceType = EnumPrices.ПТГрупп;
                    break;
                case (int)EnumPrices.Autogur73:
                    _excelParser.PriceType = EnumPrices.Autogur73;
                    break;
                case (int)EnumPrices.Композит:
                    _excelParser.PriceType = EnumPrices.Композит;
                    break;
                case (int)EnumPrices.Риваль:
                    _excelParser.PriceType = EnumPrices.Риваль;
                    break;
                default:
                    MessageBox.Show("Неверный индекс", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;
            }
        }
    }
}
