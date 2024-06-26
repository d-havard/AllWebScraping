using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class VerificationUrl
    {
        /// <summary>
        /// Check the URL put in the function and determine what website it is.
        /// </summary>
        /// <param name="url"></param>
        /// <returns>The name of the excel sheet</returns>
        public string NameUrl(string url)
        {
            string nomSheet = "";
            switch(url)
            {
                //case string ur when ur.Contains("locationId=1&gameId"):
                //    nomSheet = "EVA BEAUCHAMP";
                //    break;

                case string ur when ur.Contains("locationId=2&gameId"):
                    nomSheet = "EVA NANTES";
                    break;

                case string ur when ur.Contains("locationId=3&gameId"):
                    nomSheet = "EVA LE HAVRE";
                    break;

                case string ur when ur.Contains("locationId=4&gameId"):
                    nomSheet = "EVA STRASBOURG";
                    break;

                case string ur when ur.Contains("locationId=5&gameId"):
                    nomSheet = "EVA REIMS";
                    break;

                case string ur when ur.Contains("locationId=6&gameId"):
                    nomSheet = "EVA ROUEN";
                    break;

                case string ur when ur.Contains("locationId=7&gameId"):
                    nomSheet = "EVA TOULOUSE";
                    break;

                case string ur when ur.Contains("locationId=10&gameId"):
                    nomSheet = "EVA LA RÉUNION";
                    break;

                case string ur when ur.Contains("locationId=11&gameId"):
                    nomSheet = "EVA LILLE";
                    break;

                case string ur when ur.Contains("locationId=12&gameId"):
                    nomSheet = "EVA EVA";
                    break;

                case string ur when ur.Contains("locationId=13&gameId"):
                    nomSheet = "EVA POITIERS";
                    break;

                case string ur when ur.Contains("locationId=14&gameId"):
                    nomSheet = "EVA LYON NORD";
                    break;

                case string ur when ur.Contains("locationId=15&gameId"):
                    nomSheet = "EVA TOURS";
                    break;

                case string ur when ur.Contains("locationId=17&gameId"):
                    nomSheet = "EVA GRENOBLE";
                    break;

                case string ur when ur.Contains("locationId=20&gameId"):
                    nomSheet = "EVA LIÈGE";
                    break;

                case string ur when ur.Contains("locationId=21&gameId"):
                    nomSheet = "EVA LA ROCHELLE";
                    break;

                case string ur when ur.Contains("locationId=22&gameId"):
                    nomSheet = "EVA BORDEAUX-LAC";
                    break;

                case string ur when ur.Contains("locationId=23&gameId"):
                    nomSheet = "EVA LES MURAUX";
                    break;

                case string ur when ur.Contains("locationId=24&gameId"):
                    nomSheet = "EVA RENNES";
                    break;

                case string ur when ur.Contains("locationId=25&gameId"):
                    nomSheet = "EVA AMIENS";
                    break;

                case string ur when ur.Contains("locationId=26&gameId"):
                    nomSheet = "EVA MARSEILLE";
                    break;

                case string ur when ur.Contains("locationId=27&gameId"):
                    nomSheet = "EVA DIJON";
                    break;

                case string ur when ur.Contains("locationId=28&gameId"):
                    nomSheet = "EVA TROYES";
                    break;

                case string ur when ur.Contains("locationId=29&gameId"):
                    nomSheet = "EVA CLERMONT-FERRAND";
                    break;

                case string ur when ur.Contains("locationId=30&gameId"):
                    nomSheet = "EVA EVREUX";
                    break;

                case string ur when ur.Contains("locationId=31&gameId"):
                    nomSheet = "EVA THIONVILLE";
                    break;

                case string ur when ur.Contains("locationId=32&gameId"):
                    nomSheet = "EVA ORLÉANS";
                    break;

                case string ur when ur.Contains("locationId=34&gameId"):
                    nomSheet = "EVA PERPIGNAN";
                    break;

                case string ur when ur.Contains("locationId=36&gameId"):
                    nomSheet = "EVA BAYONNE";
                    break;

                case string ur when ur.Contains("locationId=37&gameId"):
                    nomSheet = "EVA METZ";
                    break;

                case string ur when ur.Contains("locationId=38&gameId"):
                    nomSheet = "EVA NANCY";
                    break;

                case string ur when ur.Contains("oxmozvr.fr"):
                    nomSheet = "OXMOZ TOULOUSE";
                    break;

                case string ur when ur.Contains("booking.zerolatencyvr.com"):
                    nomSheet = "ZEROLATENCY LA ROCHELLE";
                    break;
            }
            return nomSheet;
        }
    }
}
