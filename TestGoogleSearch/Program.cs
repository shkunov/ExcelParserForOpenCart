using System;
using Google.Apis.Customsearch.v1;

namespace TestGoogleSearch
{
    class Program
    {
        static void Main()
        {
            const string apiKey = "AIzaSyAgVugs7fZxeQ_VRuJwrQf-7JEEs7Im6Eo";
            const string searchEngineId = "011217440955194719924:trc1fbor6og";
            const string query = @"Багажник SUZUKI JIMNY (алюминиево магниевый сплав) JIMNY Багажник";
            var customSearchService = new CustomsearchService(new Google.Apis.Services.BaseClientService.Initializer() { ApiKey = apiKey });
            var listRequest = customSearchService.Cse.List(query);
            listRequest.Cx = searchEngineId;
            var search = listRequest.Execute();
            foreach (var item in search.Items)
            {
                Console.WriteLine("Title : " + item.Title + Environment.NewLine + "Link : " + item.Link + Environment.NewLine + Environment.NewLine);
            }
            Console.ReadLine();
        }
    }
}
