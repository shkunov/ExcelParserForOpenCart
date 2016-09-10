// Для написания кода использовалась статья:
//https://almostcode.wordpress.com/2015/09/16/simple-parser/
using System;
using System.Collections.Generic;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HtmlAgilityPack;

namespace ParserPhotoTesr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private int _count;
        private readonly List<Product> _list; 

        public MainWindow()
        {
            InitializeComponent();
            _list = new List<Product>();
        }

        private static byte[] DownloadImage(string imageUrl)
        {
            var webClient = new WebClient();
            return webClient.DownloadData(imageUrl);
        }

        private void GetImage(string url)
        {
            var myuri = new Uri(url);
            var pathQuery = myuri.PathAndQuery;
            var hostName = myuri.ToString().Replace(pathQuery, "");

            var doc = new HtmlWeb().Load(url.Trim());
            //получаем список всех постов по нашему поиску, все остальное барахло мимо
            var posters =
                doc.DocumentNode.SelectNodes("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/*");
            //получаемссылку на первый пост из нашего списка постов
            var i = 1;
            foreach (var poster in posters)
            {
                var num =
                    poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[1]").InnerText;
                var urlImg = poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[3]/a/img")
                    .GetAttributeValue("src", string.Empty);
                num = num.Replace("Арт.", "").Trim();
                var filename = System.IO.Path.GetFileName(urlImg);
                if (filename != null)
                {
                    var s = filename[0].ToString() + filename[1] + filename[2];
                    MessagesBox.Items.Add(num);
                    MessagesBox.Items.Add(hostName + urlImg);
                    // картинка в максимальном расширении
                    var imgUrl = string.Format("{0}/upload/iblock/{1}/{2}", hostName, s, filename);
                    MessagesBox.Items.Add(imgUrl);
                    _list.Add(new Product
                    {
                        Num = num,
                        ImgUrl = imgUrl
                    });
                    _count++;
                }
                i++;
            }
            //if (!string.IsNullOrEmpty(imgUrl))
            //{
            //    //создаем поток для byte[] скачанного рисунка
            //    var memoryStream = new MemoryStream(DownloadImage(imgUrl));
            //    //растягиваем рисунок по размеру пикчер бокса, тут уж можно поступать как угодно
            //    //pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            //    ////конвертируем стрим в имейдж 
            //    //pictureBox1.Image = Image.FromStream(memoryStream);
            //}
        }

        private void BtnParse_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextSearch.Text)) return;
            _count = 0;
            _list.Clear();
            var hostName = TextSearch.Text;
            MessagesBox.Items.Clear();
            //получаем html страницу со всем барахлом включая результаты нашего поиска
            var doc = new HtmlWeb().Load(hostName.Trim());
            var catalogs = doc.DocumentNode.SelectNodes("//*[@id=\"market\"]/div/div[2]/div[1]/div/div[2]/div/*/ul/*/a");
            foreach (var catalog in catalogs)
            {
                var uri = catalog.GetAttributeValue("href", "");
                GetImage(hostName + uri);
            }
            MessagesBox.Items.Add(string.Format("Всего картинок: {0}", _count));
        }

        private void CtrlCCopyCmdExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            var lb = sender as ListBox;
            if (lb == null) return;
            var selected = lb.SelectedItem;
            if (selected != null) Clipboard.SetText(selected.ToString());
        }

        private void CtrlCCopyCmdCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void RightClickCopyCmdExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            var mi = sender as MenuItem;
            if (mi == null) return;
            var selected = mi.DataContext;
            if (selected != null) Clipboard.SetText(selected.ToString());
        }

        private void RightClickCopyCmdCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }
    }
}
