using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using Newtonsoft.Json;
using System.Reflection.Metadata;

namespace Class_Webscrap
{
    public class TakeInformationOxmozvr
    {
        JsonElement root;
        JsonElement Booking_system_data;
        JsonElement start;
        JsonElement title;
        JsonElement.ObjectEnumerator Bookings;
        JsonElement Number;
        JsonElement totalAvailability;
        JsonElement NumberPlayer;
        List<string> startList = new List<string>();
        List<int> NumberPlayerList = new List<int>();
        List<int> MaximumPlayerList = new List<int>();
        string dateTime = "";

        /// <summary>
        /// Read the JSON and deserialize it to use the informations in the JSON file
        /// </summary>
        /// <returns>The JsonDocument</returns>
        public async Task<JsonDocument> GetDeserializedDocument()
        {
            string jsonContent = await File.ReadAllTextAsync("Oxmoz.json");

            var request = JsonConvert.DeserializeObject(jsonContent);
            JsonDocument document = JsonDocument.Parse(jsonContent);
            
            return document;
        }

        /// <summary>
        /// Get all the informations needed in the json to put them afterward in multiple lists
        /// </summary>
        /// <param name="document"></param>
        public void GetInformationFromJson(JsonDocument document)
        {
            root = document.RootElement;
            Booking_system_data = root.GetProperty("booking_system_data");

            string startString = "";

            JsonElement.ArrayEnumerator Events = Booking_system_data.GetProperty("events").EnumerateArray();
            foreach (var Event in Events)
            {
                start = Event.GetProperty("start");
                title = Event.GetProperty("title");
                startString = start.ToString();
                string[] startSplit = startString.Split(' ');
                dateTime = DateTime.Now.ToString("yyyy-MM-dd");
                if (startSplit[0] == dateTime)
                {
                    
                    if (title.ToString() == "Arena")
                    {
                        startList.Add(startString);
                    }

                }

            }
            Bookings = Booking_system_data.GetProperty("bookings").EnumerateObject();
            //int nb = 0;
            foreach (string start in startList)
            {
                foreach (var booking in Bookings)
                {
                    //Console.WriteLine(booking.Name.ToString());
                    string ObjectNumber = booking.Value.ToString();
                    if (ObjectNumber.Contains(start))
                    {
                        Number = booking.Value.GetProperty(start);
                        //Console.WriteLine(Number.ToString());
                        totalAvailability = Number.GetProperty("total_availability");
                        NumberPlayer = Number.GetProperty("quantity");
                        NumberPlayerList.Add(NumberPlayer.GetInt32());
                        MaximumPlayerList.Add(totalAvailability.GetInt32());
                        //Console.WriteLine($"Joueurs sur cette session : {NumberPlayer}/{totalAvailability}");
                    }
                    //Console.WriteLine(ObjectNumber.ToString());
                    //Console.WriteLine(booking);

                }
            }
        }

        public string GetDate()
        {
            return dateTime;
        }

        /// <summary>
        /// Get the List of the started hours of each sessions
        /// </summary>
        /// <returns>the List of the started hours of each sessions</returns>
        public List<string> GetStartList() 
        {
            return startList;
        }

        public List<int> GetNumberPlayerList() 
        {
            return NumberPlayerList;
        }

        public List<int> GetMaximumPlayerList()
        {
            return MaximumPlayerList;
        }
    }
}
