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

            //получаем html страницу со всем барахлом включая результаты нашего поиска
            var doc = new HtmlWeb().Load("http://www.autoventuri.ru/catalog-akkumulyatory/");
            //получаем список всех постов по нашему поиску, все остальное барахло мимо
            var posters =
                doc.DocumentNode.SelectNodes("//*[@id=\"wrap\"]/div/section/div[2]/div[6]/div[1]/div/div[3]/a/img");
            //получаемссылку на первый пост из нашего списка постов
            var htmlNode = posters.FirstOrDefault();
            if (htmlNode == null) return;
             var iUrl = htmlNode.GetAttributeValue("src", string.Empty);

            //var gagUrl = htmlNode.GetAttributeValue("data-entry-url", string.Empty);
            ////загружаем страницу с нашим постом со всем барахлом
            //doc = new HtmlWeb().Load(string.Format("{0}{1}", "http://9gag.com", gagUrl));
            ////из всей страницы выделяем конкретно тот html элемент который содержит рисунок поста
            //var poster = doc.DocumentNode.SelectNodes("//*[contains(@class, 'badge-item-img')]").FirstOrDefault();
            ////получаем из аттрибута html тега img ссылку, т.е. значение аттрибута src
            //if (poster == null) return;
            //var imgUrl = poster.GetAttributeValue("src", string.Empty);
            ////если ссылка не нул и не пустая качаем этот рисунок по ссылке
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
