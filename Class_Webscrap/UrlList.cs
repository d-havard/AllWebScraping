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

        public void MakeUrlList()
        {
            Urllist.Add($"https://booking.zerolatencyvr.com/sessions/143/{date}/?experienceId=&players=8&packageTypeId=1&priceCode=");
            Urllist.Add("https://oxmozvr.fr/wp-admin/admin-ajax.php?action=bookactiReloadBookingSystem&attributes:{\"id\":\"bookacti-wc-form-fields-product-3653\",\"class\":\"bookacti-woocommerce-product-booking-system\",\"hide_availability\":0,\"calendars\":[3],\"activities\":[8,9,11,10,12],\"group_categories\":[\"none\"],\"groups_only\":0,\"groups_single_events\":0,\"multiple_bookings\":0,\"bookings_only\":0,\"tooltip_booking_list\":0,\"tooltip_booking_list_columns\":[],\"status\":[],\"user_id\":0,\"method\":\"calendar\",\"auto_load\":0,\"start\":\"2024-06-23%2011:26:15\",\"end\":\"2026-12-10%2010:56:15\",\"trim\":1,\"past_events\":0,\"past_events_bookable\":0,\"days_off\":[{\"from\":\"2023-10-01\",\"to\":\"2023-10-01\"},{\"from\":\"2023-07-31\",\"to\":\"2023-08-04\"},{\"from\":\"2023-06-13\",\"to\":\"2023-06-23\"}],\"check_roles\":1,\"picked_events\":[],\"form_id\":7,\"form_action\":\"default\",\"when_perform_form_action\":\"on_submit\",\"redirect_url_by_activity\":{\"8\":\"https://oxmozvr.fr/commander/\"},\"redirect_url_by_group_category\":[],\"display_data\":{\"slotMinTime\":\"10:00\",\"slotMaxTime\":\"23:00\"},\"hide_events_price\":1,\"product_by_activity\":[],\"product_by_group_category\":[],\"products_page_url\":[]}");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=1&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=2&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=3&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=4&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=5&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=6&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=7&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=10&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=11&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=12&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=13&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=14&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=15&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=17&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=20&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=21&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=22&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=23&currentDate=2024-06-24");
            Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=24&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=25&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=26&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=27&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=28&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=29&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=30&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=31&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=32&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=34&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=36&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=37&currentDate=2024-06-24");
            //Urllist.Add("https://www.eva.gg/fr-FR/booking?locationId=1&gameId=38&currentDate=2024-06-24");
            
        }

        public List<string> GetUrlList()
        {
            return Urllist;
        }
    }
}
