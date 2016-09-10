//https://almostcode.wordpress.com/2015/09/16/simple-parser/

using System;
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
        public MainWindow()
        {
            InitializeComponent();
        }

        private static byte[] DownloadImage(string imageURL)
        {
            var webClient = new WebClient();
            return webClient.DownloadData(imageURL);
        }

        private void BtnParse_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextSearch.Text)) return;

            var myuri = new Uri(TextSearch.Text);
            var pathQuery = myuri.PathAndQuery;
            var hostName = myuri.ToString().Replace(pathQuery, "");

            MessagesBox.Items.Clear();
            //получаем html страницу со всем барахлом включая результаты нашего поиска
            var doc = new HtmlWeb().Load(TextSearch.Text.Trim());
            //получаем список всех постов по нашему поиску, все остальное барахло мимо
            var posters =
                doc.DocumentNode.SelectNodes("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/*");
            //получаемссылку на первый пост из нашего списка постов
            var i = 1;
            foreach (var poster in posters)

            {
                //*[@id="wrap"]/div/section/div[2]/div[6]/div[1]/div/div[1]
                //*[@id="wrap"]/div/section/div[2]/div[6]/div[2]/div/div[1]
                var num =
                    poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[1]").InnerText;
                var url = poster.SelectSingleNode("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[" + i + "]/div/div[3]/a/img")
                    .GetAttributeValue("src", string.Empty);
                //MessagesBox.Items.Add(string.Format("{0} - {1}", num, hostName + url));
                var filename = System.IO.Path.GetFileName(url);
                //http://www.autoventuri.ru/upload/iblock/9de/9de154c2a03c8f9438ffb286070b5fcf.jpeg
                if (filename != null)
                {
                    var s = filename[0].ToString() + filename[1] + filename[2]; 
                    MessagesBox.Items.Add(num);
                    MessagesBox.Items.Add(hostName + url);
                    // картинка в максимальном расширении
                    MessagesBox.Items.Add(string.Format("{0}/upload/iblock/{1}/{2}", hostName, s, filename));
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
