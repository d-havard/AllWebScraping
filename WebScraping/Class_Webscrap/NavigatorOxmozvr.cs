using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class NavigatorOxmozvr
    {
        /// <summary>
        /// Launch the emulated navigator to get the JSON file with the information that are on the website, 
        /// and download the JSON file.
        /// </summary>
        /// <param name="url"></param>
        public void LaunchNavigatorAndGetJsonFile(string url)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.Method = WebRequestMethods.Http.Get;
            httpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange";

            string file;
            var response = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var sr = new StreamReader(response.GetResponseStream()))
            {
                file = sr.ReadToEnd();
                File.WriteAllText("Oxmoz.json", file);
            }
        }
    }
}
