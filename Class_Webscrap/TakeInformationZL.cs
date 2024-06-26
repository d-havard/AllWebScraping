using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class TakeInformationZL
    {
        JsonElement.ArrayEnumerator root;
        JsonElement sessionCost;
        JsonElement StartTime;
        JsonElement BookedSlot;
        JsonElement Package;
        JsonElement MaximumPlayer;
        JsonElement PriceTierType;
        JsonElement namePriceTierType;
        JsonElement.ArrayEnumerator PackageTierPrices;
        JsonElement namePackageTierPrice;
        JsonElement costPackageTierPrice;
        JsonElement PackagePriceTierType;
        JsonElement PricePackageTierPrice;
        List<string> StartedHourList = new();
        string date = "";
        List<int> MaximumPlayerList = new();
        List<int> NumberPlayerList = new();
        List<decimal> PriceList = new();
        string jsonContent = "";

        /// <summary>
        /// Read the JSON file
        /// </summary>
        /// <param name="path"></param>
        /// <returns>Nothing</returns>
        public async Task ReadJson(string path)
        {
            jsonContent = await File.ReadAllTextAsync(path);
        }

        /// <summary>
        /// Deserialize the json file to use it later
        /// </summary>
        /// <returns>the JSONDocument where there is all the informations</returns>
        public JsonDocument DeserializeJson()
        {
            var request = JsonConvert.DeserializeObject(jsonContent);
            JsonDocument document = JsonDocument.Parse(jsonContent);
            
            return document;
        }


        /// <summary>
        /// Get every useful element to put in the excel and add them to several lists
        /// </summary>
        /// <param name="document"></param>
        public void GetElements(JsonDocument document)
        {
            root = document.RootElement.EnumerateArray();
            string comparedHour = "";
            foreach (var element in root)
            {
                StartTime = element.GetProperty("startTime");
                string[] StartTimeArray = StartTime.ToString().Split('T');
                string StartHour = StartTimeArray[1].Remove(StartTimeArray[1].Length - 1);
                date = StartTimeArray[0];
                if (StartHour != comparedHour)
                {
                    StartedHourList.Add(StartHour);
                    comparedHour = StartHour;
                    BookedSlot = element.GetProperty("bookedSlots");
                    NumberPlayerList.Add(BookedSlot.GetInt32());
                    Package = element.GetProperty("package");
                    MaximumPlayer = element.GetProperty("maximumSlots");
                    MaximumPlayerList.Add(MaximumPlayer.GetInt32());
                    PriceTierType = element.GetProperty("priceTierType");
                    namePriceTierType = PriceTierType.GetProperty("name");
                    PackageTierPrices = Package.GetProperty("packageTierPrices").EnumerateArray();

                    foreach (var PackageTierPrice in PackageTierPrices)
                    {
                        PackagePriceTierType = PackageTierPrice.GetProperty("priceTierType");
                        namePackageTierPrice = PackagePriceTierType.GetProperty("name");
                        if (namePackageTierPrice.ToString() == namePriceTierType.ToString())
                        {
                            PricePackageTierPrice = PackageTierPrice.GetProperty("price");
                            costPackageTierPrice = PricePackageTierPrice.GetProperty("cost");
                            PriceList.Add(costPackageTierPrice.GetDecimal());
                        }
                    }
                }
            }
        }

        //public void rearrangedate()
        //{
        //    string[] splitDate = date.Split('-');
        //    date = splitDate[2] + "/" + splitDate[1] + "/" + splitDate[0];
        //}

        public string Getdate()
        {
            return date;
        }

        public List<string> GetHourList()
        {
            return StartedHourList;
        }


        public List<int> GetMaximumPlayerList()
        {
            return MaximumPlayerList;
        }


        public List<int> GetNumberPlayerList()
        {
            return NumberPlayerList;
        }


        public List<decimal> GetPriceList()
        {
            return PriceList;
        }
    }
}