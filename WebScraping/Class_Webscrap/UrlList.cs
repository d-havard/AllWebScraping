using OpenQA.Selenium.DevTools.V123.LayerTree;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class UrlList
    {
        List<string> Urllist = new List<string>();
        string date = DateTime.Now.ToString("yyyy-MM-dd");

        /// <summary>
        /// Make the lists of the url of all the websites we want to webscrap
        /// 18 url where there is one arena
        /// </summary>
        public void MakeUrlList()
        {
            Urllist.Add($"https://booking.zerolatencyvr.com/sessions/143/{date}/?experienceId=&players=8&packageTypeId=1&priceCode=");
            Urllist.Add($"https://booking.zerolatencyvr.com/sessions/114/{date}/?experienceId=&players=8&packageTypeId=1&priceCode=");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=2&gameId=1&currentDate={date}");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=3&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=4&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=5&gameId=1&currentDate={date}");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=6&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=7&gameId=1&currentDate={date}");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=10&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=11&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=12&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=13&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=14&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=15&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=17&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=20&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=21&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=22&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=23&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate={date}");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=25&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=26&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=27&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=28&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=29&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=30&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=31&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=32&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=34&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=36&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=37&gameId=1&currentDate=2024-06-24");
            
        }

        /// <summary>
        /// Return the list of Url
        /// </summary>
        /// <returns>the list of Url</returns>
        public List<string> GetUrlList()
        {
            return Urllist;
        }
    }
}
