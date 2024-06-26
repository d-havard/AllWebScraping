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
        /// </summary>
        public void MakeUrlList()
        {
            Urllist.Add($"https://booking.zerolatencyvr.com/sessions/143/{date}/?experienceId=&players=8&packageTypeId=1&priceCode=");
            Urllist.Add("https://oxmozvr.fr/wp-admin/admin-ajax.php?action=bookactiReloadBookingSystem&attributes:{\"id\":\"bookacti-wc-form-fields-product-3653\",\"class\":\"bookacti-woocommerce-product-booking-system\",\"hide_availability\":0,\"calendars\":[3],\"activities\":[8,9,11,10,12],\"group_categories\":[\"none\"],\"groups_only\":0,\"groups_single_events\":0,\"multiple_bookings\":0,\"bookings_only\":0,\"tooltip_booking_list\":0,\"tooltip_booking_list_columns\":[],\"status\":[],\"user_id\":0,\"method\":\"calendar\",\"auto_load\":0,\"start\":\"2024-06-23%2011:26:15\",\"end\":\"2026-12-10%2010:56:15\",\"trim\":1,\"past_events\":0,\"past_events_bookable\":0,\"days_off\":[{\"from\":\"2023-10-01\",\"to\":\"2023-10-01\"},{\"from\":\"2023-07-31\",\"to\":\"2023-08-04\"},{\"from\":\"2023-06-13\",\"to\":\"2023-06-23\"}],\"check_roles\":1,\"picked_events\":[],\"form_id\":7,\"form_action\":\"default\",\"when_perform_form_action\":\"on_submit\",\"redirect_url_by_activity\":{\"8\":\"https://oxmozvr.fr/commander/\"},\"redirect_url_by_group_category\":[],\"display_data\":{\"slotMinTime\":\"10:00\",\"slotMaxTime\":\"23:00\"},\"hide_events_price\":1,\"product_by_activity\":[],\"product_by_group_category\":[],\"products_page_url\":[]}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=2&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=3&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=4&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=5&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=6&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=7&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=10&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=11&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=12&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=13&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=14&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=15&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=17&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=20&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=21&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=22&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=23&gameId=1&currentDate=2024-06-24");
            Urllist.Add($"https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate={date}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=25&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=26&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=27&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=28&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=29&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=30&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=31&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=32&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=34&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=36&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=37&gameId=1&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=38&currentDate=2024-06-24");
            
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
