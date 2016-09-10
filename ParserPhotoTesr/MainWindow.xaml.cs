//https://almostcode.wordpress.com/2015/09/16/simple-parser/
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
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
                MessagesBox.Items.Add(string.Format("{0} - {1}", num,  url));
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
    }
}
