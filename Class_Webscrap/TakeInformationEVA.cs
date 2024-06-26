using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL.Drawing;
using OpenQA.Selenium.Interactions;
using System.Text.Json;
using Newtonsoft.Json;

namespace Class_Webscrap
{
    public class TakeInformationEVA
    {
        
        JsonElement root;
        JsonElement data;
        JsonElement calendar;
        JsonElement sessionList;
        JsonElement.ArrayEnumerator list;
        JsonElement slot;
        string hour = "";
        bool peakHour;
        JsonElement.ArrayEnumerator availabilities;
        int maximumPlayerAvailable;
        int numberPlayer;
        bool battlepassPlayer;
        List<string> hourList = new();
        List<bool> peakHourList = new();
        List<int> MaximumPlayerAvailableList = new();
        List<int> numberPlayerList = new();
        List<bool> BattlepassPlayerList = new();


        /// <summary>
        /// Read all the json file that was previously generate, find the one who have the information we want,
        /// and delete the other files.
        /// </summary>
        /// <param name="jsonResponses"></param>
        /// <returns>the path to the good file</returns>
        public async Task<string> GetJsonFile(List<string> jsonResponses, string nomsheet)
        {
            string PathJsonFile = "..\\..\\..\\..\\Json_Files\\response_10.json";
            for (int i = 1; i < jsonResponses.Count; i++)
            {
                var jsonFilePath = $"..\\..\\..\\..\\Json_Files\\response_{nomsheet}.json";
                string jsonContent = await File.ReadAllTextAsync(jsonFilePath);

                var request = JsonConvert.DeserializeObject(jsonContent);
                JsonDocument document = JsonDocument.Parse(jsonContent);
                root = document.RootElement;
                if (root.TryGetProperty("data", out data))
                {
                    if (data.TryGetProperty("calendar", out calendar))
                    {
                        PathJsonFile = jsonFilePath;
                    }
                    else
                    {
                        //File.Delete(jsonFilePath);
                    }
                }
                else
                {
                    File.Delete(jsonFilePath);
                }
            }
            return PathJsonFile;
        }

        /// <summary>
        /// Rearrange the date found in the json and format it to be as the same format as the date in the excel
        /// </summary>
        /// <param name="JsonFilePath"></param>
        /// <returns>the new date format</returns>
        public async Task<string> RearrangeDate(string JsonFilePath)
        {
            string dateFormat = "";
            string JsonContent = await File.ReadAllTextAsync(JsonFilePath);
            JsonDocument document = JsonDocument.Parse(JsonContent);
            root = document.RootElement;
            data = root.GetProperty("data");
            calendar = data.GetProperty("calendar");
            sessionList = calendar.GetProperty("sessionList");
            list = sessionList.GetProperty("list").EnumerateArray();

            foreach (var item in list)
            {
                slot = item.GetProperty("slot");
                string date = slot.GetProperty("date").GetString();
                string[] dataArray = date.Split('-');
                dateFormat = dataArray[2] + "/" + dataArray[1] + "/" + dataArray[0];
            }
            return dateFormat;
        }

        /// <summary>
        /// Get the information and place them in local variables to save them
        /// </summary>
        /// <param name="JsonFilePath"></param>
        /// <returns></returns>
        public async Task GetInformation(string JsonFilePath)
        {
            string JsonContent = await File.ReadAllTextAsync(JsonFilePath);
            JsonDocument document = JsonDocument.Parse(JsonContent);
            root = document.RootElement;
            JsonElement data = root.GetProperty("data");
            
            calendar = data.GetProperty("calendar");
            sessionList = calendar.GetProperty("sessionList");
            list = sessionList.GetProperty("list").EnumerateArray();

            foreach (var item in list)
            {
                slot = item.GetProperty("slot");
                string hour = slot.GetProperty("startTime").GetString();
                hourList.Add(hour);
                peakHourList.Add(item.GetProperty("isPeakHour").GetBoolean());

                availabilities = item.GetProperty("availabilities").EnumerateArray();
                foreach (var availabilitie in availabilities)
                {
                    MaximumPlayerAvailableList.Add(availabilitie.GetProperty("total").GetInt32());
                    numberPlayerList.Add(availabilitie.GetProperty("taken").GetInt32());
                    BattlepassPlayerList.Add(availabilitie.GetProperty("hasBattlepassPlayer").GetBoolean());
                }
            }

        }

        /// <summary>
        /// Get the list of the start hour of every sessions
        /// </summary>
        /// <returns>the hour list</returns>
        public List<string> GetStartedHour()
        {
            return hourList;
        }

        /// <summary>
        /// Get the list of maximum player of every sessions
        /// </summary>
        /// <returns>The list of maximum player</returns>
        public List<int> GetMaximumPlayer()
        {
            return MaximumPlayerAvailableList;
        }

        /// <summary>
        /// Get the list of the actual number of player of every sessions
        /// </summary>
        /// <returns></returns>
        public List<int> GetNumberPlayer()
        {
            return numberPlayerList;
        }
        /// <summary>
        /// Get the list of the battlepass player of every sessions
        /// </summary>
        /// <returns>the list of battlepass player</returns>
        public List<bool> GetBattlePassPlayer()
        {
            return BattlepassPlayerList;
        }
        /// <summary>
        /// Get the list of Peak hour of every sessions
        /// </summary>
        /// <returns>The Peak Hour List</returns>
        public List<bool> GetPeakHour()
        {
            return peakHourList;
        }

    }
}
