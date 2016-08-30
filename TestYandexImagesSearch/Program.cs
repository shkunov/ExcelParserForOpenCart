using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TestYandexImagesSearch
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.Expect100Continue = false;

            /* Адрес для совершения запроса, полученный при регистрации IP,
            в него уже забит логин и ключ API.*/
            const string url = @"https://yandex.ru/search/xml?user=kozintsev@neosoft.su&key=03.1130000021344244:518a664a01a989f13c01396f66273c68";

            // Текст запроса в формате XML
            const string command = @"<?xml version=""1.0"" encoding=""UTF-8""?>  
          <request>  
           <query>Багажник SUZUKI JIMNY (алюминиево магниевый сплав) JIMNY Багажник</query>
           <groupings>
             <groupby attr=""d""
                    mode=""deep""
                    groups-on-page=""10""
                    docs-in-group=""1"" />  
           </groupings>  
          </request>";

            var bytes = Encoding.UTF8.GetBytes(command);
            // Объект, с помощью которого будем отсылать запрос и получать ответ.
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentLength = bytes.Length;
            request.ContentType = "text/xml";
            // Пишем наш XML-запрос в поток
            using (var requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }

            // Получаем ответ
            var response = request.GetResponse() as HttpWebResponse;

            if (response == null)
                return;

            var input = response.GetResponseStream();

            if (input == null) return;

            var xmlReader = XmlReader.Create(input);

            var xmlResponse = XDocument.Load(xmlReader);
        }
    }
}
