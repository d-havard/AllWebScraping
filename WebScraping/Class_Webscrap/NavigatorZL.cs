using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class NavigatorZL
    {
        /// <summary>
        /// Launch the emulated navigator to get the JSON file with the information that are on the website, 
        /// and download the JSON file.
        /// </summary>
        /// <param name="url"></param>
        public void DownloadJsonFile(string url, string nomSheet)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.Method = WebRequestMethods.Http.Get;
            httpWebRequest.Accept = "application/json; charset=utf-8";
            string file;
            var response = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var sr = new StreamReader(response.GetResponseStream()))
            {
                file = sr.ReadToEnd();
                File.WriteAllText($"{nomSheet}.json", file);
                Console.WriteLine($"{nomSheet}.json");
            }
        }
    }
}
